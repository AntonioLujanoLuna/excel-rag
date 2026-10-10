# Retrieval

The retrieval half of excel-rag: the Elasticsearch query layer, hybrid fusion, bounded reference
expansion, and the stateless FastAPI service that exposes them. The response is **retrieval
evidence** — coordinates, values, formulas and resolved references — never a generated answer.

Everything below applies to the two real code paths: the in-memory double the tests and a laptop
demo run against, and the live adapter in `src/excel_rag/live.py`. Where they differ — latency, the
exactness of `knn` — the difference is stated.

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
| `hits[]` | scored chunks. `source` carries `workbook_id`/`version`/`sheet`/`a1_range`; `node_id` links to the structural node; `related_node_ids` are the expansions that came back from this hit. `score_kind` says what `score` measures — `lexical` (BM25, no vector ran), `rrf` (fused rank) or `rerank` (cross-encoder) — and scores of different kinds are not comparable. |
| `nodes` | the exact structural payloads, keyed by node id: cached value, display value, formula, table/column/named-range, references, and the `depth` at which the expansion found them. |
| `unresolved_references[]` | references that could not be resolved to a node (with the reason). Never dropped. |
| `truncation` | `{truncated, reason, dropped_nodes, depth_limit, bytes_returned}` — what the budgets dropped and why. |
| `took_ms`, `es_requests` | wall-clock and the number of Elasticsearch calls the request cost. |

Blank query → `400`. Unknown workbook in `filters.workbook_ids` → `404`. Oversized body → `413`,
whether it declares a `Content-Length` or arrives chunked (the bytes are counted as they are read).
An ACL scope the caller does not hold → `403`. An unfiltered search when more workbooks are active
than one manifest read returns (10,000) → `400` `workbook_filter_required`.
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
workbook named in a filter that has no manifest is a `404`, not a silent empty result. The pin is
one `terms` clause on `version_key` (`workbook_id:vN`, carried by every document), so its size does
not grow a boolean clause per workbook. With no manifest at all nothing is active, and a search
returns nothing rather than reading a version that was never activated.

**One thing to be precise about:** a readable node's `references` edges are returned as part of that
node, so an edge may *name* a target the caller cannot read. The target's payload and value are
never returned — only the reference text that already appears inside the readable node's own
`formula`. The caller can see their own document's formulas; they cannot see a node they lack scope
for.

When `filters.acl_scopes` is empty the deployment applies `Settings.default_acl_scope`; if that is
also empty the deployment is unscoped and no ACL clause is added. This is the "no ACLs configured"
case and the route layer is where that decision lives.

## Hybrid search — and its honest limits

`retrieval/service.py`, `retrieval/repository.py`, `retrieval/fusion.py`. The query text is embedded
with the configured model (`lightonai/mDenseOn` by default, with its `query: ` prompt; chunks were
embedded at ingestion with `document: `), and two searches run over `excel_chunks`:

1. a BM25 `bool` query over title/content/headers, returning a window of 5×`top_k` hits (floored at
   50, capped at 500);
2. a `knn` query over `embedding`, returning the same window size, with `num_candidates` 4× that
   (floored at 100, capped at 10,000).

Both carry the same ACL, version and facet filters — inside the `knn` clause, so they restrict the
candidates *before* the nearest neighbours are chosen rather than thinning the result afterwards.
Neither returns the vectors in `_source`. The two ranked lists are merged by **reciprocal-rank
fusion** (RRF, k=60); a hit the vector search found lists `embedding` in its `matched_fields`.

- **Why RRF, not a weighted sum.** BM25 is unbounded and corpus-dependent, a cosine `knn` score is
  `(1 + cos) / 2`, and the in-memory double's lexical score is an ordinal match count. RRF needs
  only each list's *order*, so it behaves the same whichever client is underneath.
- **A vector from a different model is a miss.** The `knn` filter pins `embedding_model` to the
  query's model, so a chunk embedded by another model never enters the vector list: it keeps its
  lexical rank and earns no vector contribution. Changing the model means re-indexing.
- **What the vectors add.** Against a real cluster with the real model (`tests/live/test_mdenseon.py`),
  questions that share no word with the chunk they should find — “¿Qué moneda se usa para los
  ingresos?”, “Quel est le chiffre d'affaires par région ?”, “expenses per quarter” over English
  workbooks — return nothing lexically and the right table, column or formula in the hybrid top 3.
- **No embedder, no vectors.** With `EXCEL_RAG_EMBEDDING__PROVIDER=none` the endpoint is lexical and
  reports the index's own scores. A caller may also pass its own vector through
  `RetrievalService.search(request, query_vector=..., query_embedding_model=...)`; a vector without
  its model is not used.
- **The in-memory double's `knn` is exact.** It compares every filtered vector, which is the ceiling
  of what an approximate HNSW search returns; no recall figure from it is a cluster's.

## Reference expansion and truncation

`retrieval/expansion.py`. Breadth-first over the `references` edges of `excel_structure`:

- **Seeds** are the hit `node_id`s (depth 0). With `expand_references: true` the walk follows
  references up to `reference_depth` hops; the effective depth is capped at
  `BudgetSettings.reference_depth`, and the count of *related* nodes (past the seeds) at
  `BudgetSettings.max_related_nodes`. The seeds are bounded by `top_k`, so a large `top_k` never
  has its own hits' nodes reported as dropped by the related-node budget.
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

- **The live path is tested for behaviour, not measured for speed.** `tests/live` runs ingestion,
  search, expansion, range, dependents, ACL filtering and version garbage collection against a real
  Elasticsearch 8.15 (the `live-elasticsearch` CI job; locally, set `EXCEL_RAG_TEST_ES_URL`). Its
  first run found what the in-memory double could not: `delete_by_query` without a refresh left a
  replaced version visible to search and count. No latency number here comes from a cluster.
- **Recall is not measured.** `tests/live/test_mdenseon.py` checks that specific cross-lingual and
  paraphrased questions land in the top 3, which is a regression guard, not a recall figure. No
  labelled question set exists yet, so no precision/recall number is claimed for the fusion.
- One data point, not a benchmark: on a CPU-only sandbox, a hybrid query against a single-node
  Elasticsearch 8.15 took ~100 ms including the query embedding, and embedding + indexing six small
  workbooks (60 chunks) took ~17 s with the model already downloaded.

## Request validation: a misplaced field is a 400, not a no-op

The request models (`SearchRequest`, `SearchFilters`, `StructureQuery`, `RangeQuery`) are declared
`extra="forbid"`. Before that, a caller that put `expand_references` or `reference_depth` inside
`filters` instead of at the top level received **200 with no expansion and no warning** — which reads
exactly like a broken traversal, and cost a debugging round here: the first version of the seam test
reported "no reference was followed" when the request was what was wrong. A request field in the
wrong place fails loudly now: the app turns the validation error into a 400 with the standard
error envelope (`validation_error`), not a bare 422.

## The seam is tested

`tests/test_integration.py` covers the join neither half could test alone: a document ingestion writes
is one retrieval finds; every hit names a node that exists in the structure index; the formula edge
of the `cross_sheet_formula` fixture is followed end to end through the API to both of its targets;
a scope the caller lacks returns nothing; a related node the caller may not read is never returned
through the join; and a search never mixes versions.
