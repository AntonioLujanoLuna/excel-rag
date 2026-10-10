# excel-rag

**Structure-aware RAG over Excel workbooks.** An `.xlsx`/`.xlsm` (or `.xlsb`, or a CSV) is treated
like a codebase: worksheets are source files, tables and named ranges are symbols, formulas are the
dependency graph. Both the semantics and that structure are indexed in Elasticsearch, and one stateless
endpoint answers a semantic query with **exact coordinates, values and statically resolved
references**.

Design notes: [docs/ingestion.md](docs/ingestion.md), [docs/retrieval.md](docs/retrieval.md) and
[docs/context.md](docs/context.md).

## The constraint, and what follows from it

**Elasticsearch is the only persistent search/vector/metadata store.** No graph database, no SQL
engine, no second datastore. That forces three decisions:

- The canonical workbook representation exists **only during ingestion**; the durable copy of the
  structure is the `excel_structure` index.
- References are resolved **at ingestion** into typed edges between stable node ids. A retrieval
  request never parses a formula.
- Range overlap is answered by **`integer_range` intersection**, not by walking cells, which is what
  keeps a `SUM(Actuals!D2:D500)` edge from becoming 499 edges.

## Architecture

```text
Excel (.xlsx / .xlsm)
        |
  ingestion (openpyxl + region detection + formula parser)
        |
  canonical representation (in memory: hierarchy, ranges, tables, values, dependencies)
        |
  embedding (lightonai/mDenseOn) + bulk index
        |
        +--------------------------------+
        | Elasticsearch                  |
        |   excel_chunks    <- BM25 + dense vectors (knn), fused by RRF
        |   excel_structure <- nodes, values, typed edges
        |   excel_versions  <- active version per workbook
        +--------------------------------+
                       ^
         stateless FastAPI: POST /api/v1/search/excel
                       |
        hits + A1 coordinates + optional related nodes
        + unresolved references + truncation metadata
```

## Two indices

`excel_chunks` holds what is *searchable* — workbook and sheet summaries, region and table
descriptions, column schemas, contextualized row groups, formula summaries — with the A1 range,
headers, ACL scope and the embedding model that produced the vector.

`excel_structure` holds what is *exact* — cells, ranges, tables, columns, named ranges, formulas —
with `row_span`/`column_span` as `integer_range` fields, the cached value, the formula text, and
`references` as nested typed edges. Charts, pivot tables and data validations carry the same
edges, so "what reads this cell?" finds the chart that plots it as well as the formula that sums
it. Unresolvable references (`INDIRECT`, volatile `OFFSET`, external links, unsupported dynamic
arrays) are stored with their reason, never invented.

Both mappings are data in `src/excel_rag/es.py`, and a test asserts the code only touches fields
the mappings declare — Elasticsearch happily accepts a document whose unknown field can then never
be matched on, which is the one failure this arrangement prevents.

The chunk mapping is built by `index_mappings(dims)`, and the dimensionality comes from
`EmbeddingSettings.dims` (`EXCEL_RAG_EMBEDDING__DIMS`). That is not cosmetic: an indexed
`dense_vector` needs its `dims` at index-creation time and a real cluster **refuses the create**
without it, while the in-memory double accepts any mapping at all. Since no test running against a
Python dict could notice, the dimensionality is asserted directly (including a test that documents
the double's blindness), and changing it is an index rebuild.

## Retrieval contract

```bash
curl -s localhost:8080/api/v1/search/excel -H 'content-type: application/json' -d '{
  "query": "How is projected revenue calculated?",
  "filters": {"workbook_ids": ["wb42"]},
  "top_k": 10,
  "include_structure": true,
  "expand_references": true,
  "reference_depth": 1,
  "max_related_nodes": 20
}' | jq
```

Search is **hybrid**: the query is embedded with
[`lightonai/mDenseOn`](https://huggingface.co/lightonai/mDenseOn) — a 307M-parameter multilingual
dense retriever (768 dimensions, `query:`/`document:` prompts) — and a `knn` search over the chunk
embeddings runs next to BM25, under the same ACL and version filters; the two ranked lists are
merged by reciprocal-rank fusion. A question in Spanish, French or German finds the English
workbook's table that answers it, which no keyword query does.

The response carries `hits` (scored chunks with `source.workbook_id`/`version`/`sheet`/`a1_range`),
`nodes` (the exact structural payloads, keyed by node id), `unresolved_references` and
`truncation` — what a bounded expansion dropped and why. It is **retrieval evidence, not a
generated answer**.

Optional direct-inspection endpoints: `GET /api/v1/excel/{workbook_id}/structure`,
`POST /api/v1/excel/range` (which nodes overlap a rectangle) and `POST /api/v1/excel/dependents`
(which formulas read a rectangle -- "if I change `Assumptions!C7`, what moves?") for callers that need
deterministic structure without a semantic query. Dependents are answered by intersecting the
`integer_range` spans every nested reference edge carries, so a formula reading `Actuals!D2:D500` is
a dependent of `Actuals!D100` without one edge per cell.

## Attaching a workbook to a conversation instead

Not every workbook needs an index. `excel_rag.context` renders an attached `.xlsx` (a path or an
upload's bytes) as markdown for a context window within a token budget — grids with row numbers and
column letters, regions with headers, units and notes, formulas with Excel's saved values and what
they read, every omission marked — and gives the model four tools (`read_range`, `find`,
`precedents`, `dependents`) for what the budget left out. Same reader, same guarantees, no
Elasticsearch. See [docs/context.md](docs/context.md), `excel-rag render`, and
`examples/ask_workbook.py`. `excel-rag diff before.xlsx after.xlsx` says what changed between two
versions — inputs, formulas, saved results — and which formulas read each changed input.

## Using it from Claude or another MCP client

`excel-rag mcp` serves the workbook tools over stdio to any MCP client: `list_workbooks`,
`render_workbook`, `read_range`, `find`, `precedents`, `dependents` and `diff_workbooks`, on the
`.xlsx`/`.xlsm` files under the directories given with `--root` (default: the working directory).
Every tool is read-only, and a path outside the roots is refused. With `--search` it also serves
`search_index`, the retrieval endpoint's search over the configured Elasticsearch.

```bash
uv sync --extra mcp
claude mcp add excel-rag -- uv run --directory "$PWD" excel-rag mcp --root ~/Spreadsheets
```

For Claude Desktop, the same command goes in `claude_desktop_config.json`:

```json
{"mcpServers": {"excel-rag": {"command": "excel-rag", "args": ["mcp", "--root", "/path/to/books"]}}}
```

## Invariants

- **No cross-version reads.** Ids are deterministic within `(workbook_id, version)`; a replacement
  version is indexed before activation and older versions are garbage-collected.
- **ACL filters apply to primary hits *and* to every related-node lookup.** `_mget` does not enforce
  document-level ACLs, so the service does.
- **ACL scopes come from the caller's token, not the request body.** A configured principal is
  filtered by the scopes it holds; `filters.acl_scopes` may narrow them, and naming one it does not
  hold is a 403. The same applies to the direct-inspection routes.
- **Bounded expansion.** Depth, node count, payload bytes and wall-clock are all capped, and a
  truncated expansion says so.
- **No cell-per-document explosion.** Large grids are bounded range and row-group nodes; a workbook
  that would exceed `max_documents_per_workbook` is refused, not silently clipped.
- **Macros are never executed**, external links are never refreshed, formulas are never evaluated.
- **A cached value is not a computed value** and is labelled as such wherever it is returned.

## Quickstart

```bash
uv sync --extra dev
uv run pytest                                  # in-memory Elasticsearch double, no model weights
uv run excel-rag inspect path/to/book.xlsx     # print the regions, tables and counts ingestion detects
```

No `.xlsx` is committed: the test workbooks are generated at test time by
`tests/fixtures/make_fixtures.py`.

`index` and `serve` are separate processes, and the in-memory double lives only as long as one of
them, so indexing and then serving needs a live cluster:

```bash
uv sync --extra dev --extra es --extra embed   # embed: sentence-transformers + torch for mDenseOn
export EXCEL_RAG_USE_LIVE_ELASTICSEARCH=true
export EXCEL_RAG_ELASTICSEARCH__URLS='["http://127.0.0.1:9200"]'
uv run excel-rag index path/to/book.xlsx --workbook-id wb42 --version 1
uv run excel-rag serve
```

The model downloads (~1.2 GB) on first use and loads before the server accepts requests. Embedding
settings: `EXCEL_RAG_EMBEDDING__MODEL` (default `lightonai/mDenseOn`), `__DIMS` (768),
`__DEVICE` (`cpu`/`cuda`/`mps`), `__BATCH_SIZE`, `__MAX_SEQ_LENGTH` (1024 tokens per chunk), and
`__PROVIDER=none` to index and search without vectors. An optional cross-encoder reranks the head
of the fused list: `EXCEL_RAG_RERANK__PROVIDER=cross-encoder` (default model
`cross-encoder/mmarco-mMiniLMv2-L12-H384-v1`, `__WINDOW` 50 candidates). On a CPU-only host, install torch from
`https://download.pytorch.org/whl/cpu` first to skip the CUDA wheels.

To measure retrieval quality — hit@k, recall@k and MRR for lexical, hybrid and reranked search over
question → expected-range pairs — run `excel-rag evaluate` (a built-in sample), `--mine` questions from
your workbooks' own labels, or pass your own
`--cases` and `--workbook ID=PATH`; see [docs/retrieval.md](docs/retrieval.md#measuring-retrieval-quality).

To run the live tests, including the real model end to end:

```bash
EXCEL_RAG_TEST_ES_URL=http://127.0.0.1:9200 EXCEL_RAG_TEST_MDENSEON=1 uv run pytest tests/live
```

Callers and the scopes they hold (any configured token makes `/api/v1` require one):

```bash
export EXCEL_RAG_SERVER__PRINCIPALS='[
  {"name": "finance-app", "token": "…", "acl_scopes": ["finance-team"]},
  {"name": "admin", "token": "…", "unrestricted": true}
]'
```

`EXCEL_RAG_SERVER__SERVICE_TOKEN` remains for a single trusted caller that applies its own
authorisation: it is unrestricted and passes the scopes it names through.

## Phases

1. **MVP ingestion and search** — parse without executing macros; index sheets, regions, chunks,
   headers, A1 coordinates and cached values; stateless hybrid search with citations and filters.
2. **Cross-reference retrieval** — formula precedents, named ranges, structural nodes, range
   overlap queries, bounded expansion with cycle detection, explicit unresolved references.
3. **Production hardening** — active-version switching, deterministic reindexing, stale-version
   cleanup, per-request budgets, large-workbook benchmarks, and parsing hardened against zip bombs
   and hostile sheet dimensions.

## Non-goals

Multi-turn orchestration, agent tool selection, answer generation, feedback-based relevance
evaluation, executing spreadsheets, SQL analytics, recalculating Excel formulas, and any additional
persistent store.
