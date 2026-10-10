# Ingestion

How an `.xlsx`/`.xlsm` becomes `excel_chunks` and `excel_structure` documents. The canonical
representation lives **only in memory**; Elasticsearch is the durable copy, so everything here is
about reading once, deciding what is searchable and what is exact, and failing explicitly.

Reading and modelling live in `src/excel_rag/workbook/`, shared with the context renderer
(`src/excel_rag/context/`, which renders an attached workbook for a conversation without indexing
it); turning the model into documents lives in `src/excel_rag/ingest/`:

| module | responsibility |
|---|---|
| `workbook/reader.py` | openpyxl loading (from a path or bytes) with the macros/dimension/zip guards |
| `workbook/regions.py` | split a sheet into title / table / notes regions |
| `workbook/formulas.py` | formula text → typed edges and explicit gaps |
| `workbook/build.py` | assemble the canonical `WorkbookModel`, group formula clusters |
| `ingest/documents.py` | canonical model → `ChunkDocument` / `StructureDocument` |
| `ingest/indexer.py` | embedding, index creation, bulk write, budget refusal, version flip |
| `cli.py` | `index`, `inspect`, `render`, `serve` |

## Decisions and why

### Two loads: `data_only=False` then `data_only=True`

Openpyxl exposes a formula cell's text **or** its last-saved value, never both, under one flag. So the
reader loads twice: the formula workbook supplies cell values and formula text (and merges, tables
and defined names), the cached workbook supplies the last value Excel saved for formula cells. A
cached value is copied verbatim and labelled as cached — it is **not** a freshly computed result, and
nothing in ingestion recomputes it. (The `cached_value` fixture deliberately stores `999` for a cell
whose formula is `=A1+B1`; the documents carry `999`, never `5`.)

### Macros are noticed, never loaded

`load_workbook(..., keep_vba=False)` never loads the VBA project, and openpyxl never executes VBA.
Macro and VBA information is read from the **raw package** instead: a `vbaProject.bin` part sets
`has_vba`, and a sheet whose workbook relationship type is `…/xlMacrosheet` (or whose part content
type is `…macrosheet+xml`) is recorded as a macro sheet and its cells are not read. A formula that
points at such a sheet produces `UnresolvedReason.MACRO_SHEET`, not a fabricated edge.

### A declared dimension is reported, not trusted

A workbook may declare `<dimension ref="A1:XFD1048576">` while holding four cells. openpyxl already
derives the used range from the cells that exist, so nothing is allocated for the declared area; the
reader reads the string from the package purely to flag it (`declared_dimension_flagged`) and warn. A
malformed dimension parses to area `0` and is silently ignored. A genuinely oversized package (sum of
declared uncompressed sizes over 512 MiB) is refused before openpyxl touches it.

### Clear failures

An unreadable package (corrupt, truncated or password-protected — all land as a non-zip) raises
`IngestError`, as does a missing file and a package without `xl/workbook.xml`. No bare
`zipfile.BadZipFile` or openpyxl trace ever escapes.

## Region detection rules

A worksheet is **not** one table. Detection runs in two stages.

**1. Split into blocks.** Rows empty across the used width and columns empty across the used height
are boundaries; the cross-product of the surviving bands gives rectangular blocks. Two tables on one
sheet — separated by a blank row, a blank column, or both — therefore come out as two blocks. (The
`two_tables_one_sheet` fixture is exactly this, plus a merged title banner.)

**2. Classify each block.**

- **Title** — one or more leading rows that are a single string cell inside a merge spanning the
  block width. A single-cell, all-text block whose value is a string is also a title.
- **All-text block** (no numeric, boolean, date or formula cell anywhere): a one-row block is a
  title; otherwise it is a **notes** block if any row starts with a note word (`Note`, `Notes`,
  `Assumption(s)`, `Source(s)`, `Legend`, `Disclaimer`, `Warning`, `Footnote(s)`, `Data as of`,
  `As of`, `Prepared by`, `Methodology`, `Commentary`, `Definition(s)`), else a **generic** region.
- **Header rows** — the run of *label-like* rows at the top: a row is label-like when strictly fewer
  than half its populated cells are data cells (numeric, boolean, date or formula). A run of two rows
  under a merged banner is a **multi-level header**.
- **Units row** — a label-like row directly under at least one header row whose every populated cell
  is a units word (`USD`, `%`, `bps`, `kg`, `in thousands`, `per unit`, …). It is peeled off the
  header and recorded per column.
- **Data rows** — everything after the header/units run. Trailing note-like rows are pulled out into
  their own notes region, so a `Total` row stays with the data but a `Notes:` line does not.

**Columns** are built per block column that has a header or data cell. A merged group label is
propagated to every column the merge spans, so `Sales` over `B1` and `Q2` under it yields the name
`Sales / Q2`. Each column records name, aliases (the individual header-level texts), inferred type
(number / date / percentage / currency / text / boolean / formula / error / mixed), distinct count,
up to five examples, units, A1 range and a deterministic node id.

An Excel table object (`ws.tables`) whose ref lies inside a region names that region (`table_name`)
and switches its node id and type from `region`/`REGION` to `table`/`TABLE`.

**Known limits** (deliberate): a header whose labels are themselves numbers (e.g. years stored as
integers) is read as data, since a numeric row is not label-like; and a table of exclusively text
data is treated as prose, not a table.

## What qualifies as an "important" cell

The design forbids a cell-per-document explosion, so a structure cell node is created only for:

- **header cells** — one per column, the topmost header-row cell of that column;
- **formula cells** — a single-cell formula becomes a `FORMULA` node; a *cluster* of one repeated
  pattern (see below) becomes **one** `RANGE` node, not one node per cell;
- **named-range anchors** and the targets of every resolved reference (`RANGE`, `CELL` or table
  `COLUMN` node), so a `target_node_id` never dangles.

Ordinary data cells never get their own document. A 500×6 region produces six header cell nodes and
nothing else.

### Formula clusters

Formula cells are grouped by a normalised pattern: the row digits of *relative* references are
replaced with `#` (`=B2*$B$1` and `=B3*$B$1` share a pattern; `=B2+100` and `=B3+200` do not, because
`100`/`200` are constants, not references). A group of one is a normal formula node. A group of many
becomes one `RANGE` node covering the group, one `FORMULA_SUMMARY` chunk, and one broadened edge per
relative precedent — so a 200-row `=A2*2` column is a single edge to `A2:A201`, matching the
"`SUM(Actuals!D2:D500)` is one edge" rule.

### Array formulas and data tables

openpyxl hands an array formula (a legacy CSE formula, or a dynamic-array formula Excel 365 saved) to
the reader as an `ArrayFormula` object and a what-if data table as a `DataTableFormula`, not as text.
The reader takes the formula text from the former and writes the latter as Excel shows it,
`=TABLE(row_input, column_input)`, whose input cells become edges and which carries an
`unsupported_function` gap: Excel recomputes the table by substituting each input, which no static
edge can say. An array formula is never folded into a cluster, and its formula node records the
extent its result last covered (`array_range`).

Excel saves the rest of that extent as bare values, which openpyxl reads as constants. They are
formula *results*, so the reader keeps them as cached values pointing at their master cell
(`array_master`): region detection types them as formula columns, and the rendering marks them `ƒ`
like any other saved formula value.

### Named ranges

A workbook-wide name keys its node on the name (`wb:v1:named_range:growthrate`); a sheet-scoped one on
`Sheet!Name` (`wb:v1:named_range:jan!rate`), so the same local name on twelve monthly sheets is twelve
nodes. openpyxl 3.1 keeps sheet-scoped names on the worksheet rather than the workbook; the reader
reads both. A bare name in a formula resolves to its own sheet's local name first and to the
workbook-wide one otherwise, as Excel does; `Feb!Rate` resolves on `Feb`.

## Formula reference extraction

Pure text analysis, never evaluation. Formulas are tokenised with openpyxl's
`openpyxl.formula.tokenizer.Tokenizer`, which separates function calls, string literals and reference
operands; only `OPERAND RANGE` tokens are resolved, so a function whose name spells a cell (`LOG10`,
`ATAN2`, `DAYS360`) is never an edge. It records typed `Reference` edges for same-sheet and
cross-sheet A1 cells and ranges (including `'Bob''s Inputs'!B2` and `Sheet1!A1:Sheet1!B2`), absolute
(`$A$1`) and relative refs, whole-column refs (`A:B`, clipped to the used height), whole-row refs
(`2:3`, clipped to the used width), 3-D refs (`Jan:Mar!B2`, one edge per sheet in workbook order),
structured table refs (`Table1[Amount]`), and workbook- or sheet-scoped named ranges. Everything that
cannot be resolved statically becomes an `UnresolvedReference` with the right reason:

| trigger | reason |
|---|---|
| `INDIRECT(...)` | `indirect` |
| `OFFSET(...)` | `volatile_offset` |
| `[Budget.xlsx]…`, `'…[Book.xlsx]Sheet'!A1`, `[1]…` | `external_link` |
| `LET`, `LAMBDA`, `@`, spill `#`, dynamic-array functions | `dynamic_array` |
| `_xlfn.`/`_xludf.` forms; a range bounded by a function (`A1:INDEX(…)`) | `unsupported_function` |
| reference to a macro sheet | `macro_sheet` |
| unknown sheet, table, column or name; a 3-D ref with unknown sheet order | `out_of_range` |
| `#REF!`, a range across two sheets, a formula the tokenizer refuses | `malformed` |

Inside `LET`/`LAMBDA` a bare identifier is a local variable, so it is not reported as an unknown name
(the formula already carries its `dynamic_array` gap).

A string literal is never scanned for references (`=IF(A1="B2",…)` yields only the `A1` edge), and
duplicate edges are collapsed. Edges point at ranges, not cells.

## Chunk text templates

One region, one column, one row group and one formula summary, each including the sheet and the A1
range so a hit can be opened in Excel.

- **Workbook:** `Workbook 'wb' (file two_tables.xlsx, version 1) with 1 worksheet(s): Report. 3
  region(s)/table(s) detected.`
- **Sheet:** `Worksheet 'Report' (visibility visible) in workbook wb, used range A1:C9. 3 region(s)
  detected.`
- **Region/table (worked example, `two_tables_one_sheet` A3:C5):**
  `Table on worksheet Report covering A3:C5. Titled 'Region'. Columns: Region (text), Q1 (number),
  Q2 (number). 2 data row(s) across 1 row group(s).`
- **Column:** `Column 'Q1' (B) on worksheet Report spanning B4:B5: inferred type number, 2 distinct
  value(s). Examples: 100, 90.`
- **Row group, 500×6 fixture, `Data!A2:F64` (group size 63):** `col_0 rows 2-64 on worksheet Data
  (A2:F64). Columns: col_0, col_1, col_2, col_3, col_4, col_5. col_0=1; col_1=2; col_2=3; … |
  col_0=7; col_1=8; …` Every row carries its column headers, so a single row is meaningful
  standalone.
- **Formula summary:** `Formulas on worksheet Forecast at B2 compute: =SUM(Actuals!D2:D500)*(1+
  Assumptions!C7). Reads: Actuals!D2:D500 (range), Assumptions!C7 (cell).`

Row-group text writes at most `MAX_TEXT_COLUMNS = 12` column=value pairs per row and notes how many
columns were omitted.

## Document-count bounds

Per region, regardless of its height:

- 1 region/table node + 1 chunk;
- 1 column node + 1 chunk per column;
- at most `MAX_ROW_GROUPS_PER_REGION = 8` row-group nodes + chunks (group size is
  `max(ROW_GROUP_ROWS=50, ceil(rows/8))`, so 500 rows give 8 groups of 63);
- 1 formula node + summary per single formula or formula cluster;
- 1 cell node per header cell and per referenced target.

Plus one workbook node/chunk and one sheet node/chunk per sheet. The total therefore grows with
columns, regions and distinct formula patterns — **not** with rows.

## What is refused

A workbook that would write more than `budgets.max_documents_per_workbook` (default 200 000)
documents is refused with `IngestError` **before** anything is written, never silently clipped. An
unreadable package, a package over the uncompressed-size ceiling, and a file without
`xl/workbook.xml` are refused at read time.

## Version replacement

`Indexer.index_workbook` writes the new version's chunks and structure documents, then flips the
`excel_versions` manifest to the new version, then deletes the previous version's documents by a
`workbook_id`+`version` query. Ids are deterministic within `(workbook_id, version)`, so a reader
that consults the manifest never sees a half-built version and no reference ever crosses versions.

## Measured document counts

Gates: `pytest --cov` at 91.78% (floor 85); 149 tests pass. Counts below are from `ingest_workbook`
on the generated fixtures (chunks + structure).

| fixture | regions | chunks | structure | total |
|---|---:|---:|---:|---:|
| two_tables_one_sheet | 3 | 12 | 17 | 29 |
| merged_multilevel_header | 1 | 8 | 12 | 20 |
| units_and_notes | 2 | 7 | 9 | 16 |
| cross_sheet_formula | 4 | 16 | 22 | 38 |
| named_range | 2 | 11 | 16 | 27 |
| indirect_offset | 1 | 9 | 13 | 22 |
| table_object | 1 | 6 | 8 | 14 |
| cached_value | 1 | 8 | 11 | 19 |
| hostile_dimension | 1 | 6 | 8 | 14 |
| macro_workbook | 1 | 8 | 10 | 18 |
| **large_500x6 (3000 cells)** | 1 | 17 | 23 | **40** |
| large_5000x6 (30000 cells) | 1 | 17 | 23 | **40** |
| wide_400x40 (refused at budget 100) | 1 | 51 | 91 | 142 |

The 500×6 grid yields 40 documents, not thousands, and a ten-times-taller 5000×6 grid yields the
same 40 — the bound is independent of row count. The `wide_400x40` fixture exceeds a 100-document
budget and is the refusal case.
