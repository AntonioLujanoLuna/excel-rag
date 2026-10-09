# Retrieval

The retrieval half of excel-rag: the Elasticsearch query layer, hybrid fusion, bounded reference
expansion, and the stateless FastAPI service that exposes them. The response is **retrieval
evidence** — coordinates, values, formulas and resolved references — never a generated answer.

Everything below applies to the two real code paths: the in-memory double the tests and a laptop
demo run against, and the live adapter in `src/excel_rag/live.py`. Where they differ — latency, the
vector search — the difference is stated.

## Endpoints

All routes live under `/api/v1` except `/health`. When `ServerSettings.service_token` is set, every
`/api/v1` route requires it as `X-Service-Token: <token>` or `Authorization: Bearer <token>`.
`/health` is always open.

### `POST /api/v1/search/excel`

The primary endpoint. Body is the frozen `SearchRequest`; response is the frozen `SearchResponse`.

```bash
curl -s localhost:8080/api/v1/search/excel -H 'content-type: application/json' -d '{
  "query": "How is projected revenue calculated?",
  "filters": {"workbook_ids": ["wb42"], "acl_scopes": ["finance-team"]},
  "top_k": 10,
  "include_structure": true,
  "expand_references": true,
  "reference_depth": 1,
  "max_related_nodes": 20
}' | jq
```

Response fields:

| field | meaning |
|---|---|
| `hits[]` | scored chunks. `source` carries `workbook_id`/`version`/`sheet`/`a1_range`; `node_id` links to the structural node; `related_node_ids` are the expansions that came back from this hit. |
| `nodes` | the exact structural payloads, keyed by node id: cached value, display value, formula, table/column/named-range, references, and the `depth` at which the expansion found them. |
| `unresolved_references[]` | references that could not be resolved to a node (with the reason). Never dropped. |
| `truncation` | `{truncated, reason, dropped_nodes, depth_limit, bytes_returned}` — what the budgets dropped and why. |
| `took_ms`, `es_requests` | wall-clock and the number of Elasticsearch calls the request cost. |

Blank query → `400`. Unknown workbook in `filters.workbook_ids` → `404`. Oversized body → `413`.
Every non-2xx carries `{"error": {"type", "message", "details"}}`.

### `GET /api/v1/excel/{workbook_id}/structure`

Structural nodes for one workbook at its active version, no semantic query.

```bash
curl -s "localhost:8080/api/v1/excel/wb42/structure?sheet_names=Forecast&node_types=cell&limit=200" | jq
```

Unknown workbook → `404`.

### `POST /api/v1/excel/range`

Nodes whose `row_span`/`column_span` intersect an A1 rectangle. Answered by `integer_range`
intersection on both axes — **never by walking cells**, which is what keeps a `SUM(Actuals!D2:D500)`
reference from becoming 499 references.

```bash
curl -s localhost:8080/api/v1/excel/range -H 'content-type: application/json' -d '{
  "workbook_id": "wb42", "sheet_name": "Forecast", "a1": "A12:C20",
  "node_types": ["cell", "region"], "limit": 50
}' | jq
```

A malformed A1 rectangle → `400` (`invalid_range`). Unknown workbook → `404`.

### `GET /health`

```bash
curl -s localhost:8080/health | jq
# {"status":"ok","version":"0.1.0","use_live_elasticsearch":false,
#  "indices":{"excel_chunks":2160,"excel_structure":32244,"excel_versions":3}}
```

## Where ACL and version filtering are enforced

**One place builds every query**: `retrieval/repository.py`. Primary chunk search, direct structure
queries, range-intersection queries and the related-node `_mget` all go through a single `Scope`
holding the caller's ACL scopes and the active `(workbook_id, version)` pins.

- `Scope.filters()` produces the server-side clauses: a `terms` clause on `acl_scope` (when scopes
  are set) and a `bool/should` of `(term workbook_id AND term version)` per active workbook. It is
  added to the primary search, to `query_structure` and to `query_range`.
- **`_mget` enforces nothing.** A multi-get is a document fetch, not a query, so Elasticsearch hands
  back whatever id it holds. `Scope.visible()` re-checks every document `_mget` returns, and
  `Repository.get_nodes` classifies the result: `nodes` (visible), `denied` (exists but filtered by
  ACL or version) and `missing` (absent). A denied node is **never** put in `nodes`.

A request never mixes versions: the version pin comes from the `excel_versions` manifest, and a
workbook named in a filter that has no manifest is a `404`, not a silent empty result.

**One thing to be precise about:** a readable node's `references` edges are returned as part of that
node, so an edge may *name* a target the caller cannot read. The target's payload and value are
never returned — only the reference text that already appears inside the readable node's own
`formula`. The caller can see their own document's formulas; they cannot see a node they lack scope
for.

When `filters.acl_scopes` is empty the deployment applies `Settings.default_acl_scope`; if that is
also empty the deployment is unscoped and no ACL clause is added. This is the "no ACLs configured"
case and the route layer is where that decision lives.

## Hybrid fusion — and its honest limits

`retrieval/fusion.py`. The lexical query returns a bounded **candidate window** (5×`top_k`, floored
at 50, capped at 500). When the caller supplies a query vector and its embedding model, each chunk
in that window is scored by cosine similarity against its stored `embedding`, and the two rankings
are combined by **reciprocal-rank fusion** (RRF, k=60).

- **Why RRF, not a weighted sum.** The lexical score this service receives is *ordinal*: the
  in-memory double returns counts of matched leaf clauses, not BM25 — the live client returns BM25.
  A fixed weight cannot reconcile two scales that differ by client. RRF needs only each retriever's
  *order*, so it behaves the same whichever client is underneath.
- **A vector from a different model is a miss.** Each chunk carries `embedding_model`. A chunk whose
  stored vector was produced by a different model (or of a different dimensionality) is excluded from
  the vector ranking entirely: it keeps its lexical rank and earns no vector contribution. It is not
  scored against an incomparable vector, and no error is raised.
- **This is a candidate-window re-rank, not an exhaustive vector search.** The in-memory client does
  not implement `knn` — it refuses it by design. Vectors are scored in Python over the window the
  lexical query already returned, so a chunk that the lexical query did not surface cannot be
  promoted. The live adapter builds a real `knn` query
  (`live.build_knn_query`, carrying the same ACL/version `filter` clauses); there the candidate
  window is whatever Elasticsearch returns.
- **The HTTP endpoint runs lexical-only.** The frozen `SearchRequest` has no vector field, so
  `POST /api/v1/search/excel` has no way to carry one; it scores with the index's own order. The
  hybrid path is reached by calling `RetrievalService.search(request, query_vector=..., query_embedding_model=...)`
  directly, which is what the fusion tests and the service tests do. This is a deliberate consequence
  of not editing the frozen model.

## Reference expansion and truncation

`retrieval/expansion.py`. Breadth-first over the `references` edges of `excel_structure`:

- **Seeds** are the hit `node_id`s (depth 0). With `expand_references: true` the walk follows
  references up to `reference_depth` hops; the effective depth is capped at
  `BudgetSettings.reference_depth`, and the node count at `BudgetSettings.max_related_nodes`.
- **Cycle detection.** A formula graph has cycles (`A1 → B1 → A1`). A visited set makes every node
  appear at most once and the traversal terminate. The response schema has no cycle field, so this is
  a tested behaviour (`tests/retrieval/test_expansion.py::TestCycles`) rather than a payload field.
- **Silent omission is a bug.** Budget stops are reported in `TruncationInfo`:
  - `max_related_nodes` — the node budget was hit; `dropped_nodes` counts the refused nodes.
  - `max_payload_bytes` — the serialized payload budget (`BudgetSettings.max_payload_bytes`) was hit.
  - `depth_limit` — the depth bound stopped expansion while references remained; `depth_limit` is set.
  - `timeout` — the wall-clock budget (`BudgetSettings.timeout_seconds`) elapsed.
- **Unresolvable references are surfaced.** A reference whose target is absent from the index becomes
  a `MALFORMED` entry in `unresolved_references`, and each node's own stored `unresolved_references`
  (INDIRECT, volatile OFFSET, external links, …) are carried through.
- **ACL applies at every level.** A node the caller may not read is excluded and only counted
  (`ExpansionResult.denied`), never named — surfacing it would leak its existence.

## Benchmarks — measured, on the in-memory client

`python -m excel_rag.bench` builds a synthetic corpus (3 workbooks × 4 sheets, 64-dim embeddings),
indexes it into the in-memory double (2160 chunk documents, 32244 structure nodes), and measures
each endpoint. The numbers below are **as measured on this machine**, and they are **not**
Elasticsearch latencies and **not** production latencies: the "index" is a Python dict and the call
goes through an in-process `TestClient`. The one figure that would carry over to a cluster is the
Elasticsearch call count.

```
excel-rag benchmark  (version 0.1.0)
corpus: 2160 chunk docs, 32244 structure nodes, 64-dim embeddings

per-endpoint latency (ms) and Elasticsearch calls per request
endpoint                                                     p50     p95     p99    mean  es_req
------------------------------------------------------------------------------------------------
POST /api/v1/search/excel (lexical)                       121.80  223.86  237.28  129.12       2
POST /api/v1/search/excel (structure + expansion depth 1)  122.68  220.09  222.49  129.60       4
GET /api/v1/excel/{workbook}/structure                    501.12  515.87  520.11  460.22       3
POST /api/v1/excel/range                                  168.04  171.37  171.67  168.32       3

reference expansion (depth 0 = no related nodes; 1 and 2 follow references)
 depth   nodes  mean_nodes  truncated   p50_ms   p95_ms  es_req
------------------------------------------------------------------------------------------------
     0      10        10.0      False   123.00   126.06       3
     1      20        20.0       True   122.73   220.48       4
     2      21        21.0       True   121.98   221.92       5
```

Reading it honestly:

- The structure endpoint is the slow one (~0.5 s p50) because the in-memory double scans and deep-
  copies every matching document — an O(index) full scan with no inverted index. A real cluster
  answers this with a term filter; that is exactly the cost this double does *not* model.
- Expansion at depth 1 and 2 is reported `truncated` because the bench budget stops the walk while
  more references remain (the depth bound) — the flag is doing its job, not failing.
- The latency here is dominated by TestClient/pydantic overhead and Python-object scanning; it says
  nothing about a real deployment's p50.

## What is not measured

- **The live path is UNMEASURED.** No Elasticsearch cluster is reachable from the machine this was
  built on. `live.py` is written and unit-tested against a stand-in client (`tests/retrieval/test_live.py`),
  but no query in this repository has run against a real cluster, and no latency here comes from one.
- **Recall is not measured.** The in-memory double returns ordinal match counts and does not
  implement dense-vector search, so no precision/recall number is claimed for the fusion.
- Real `knn` behaviour, index-time analysis, and production BM25 scoring are all cluster-side and
  unmeasured here.
