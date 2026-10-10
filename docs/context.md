# Attaching a workbook to a conversation

Indexing is for a corpus that is searched again and again. A workbook attached to one conversation
needs none of it: no Elasticsearch, no embeddings, nothing persisted. `excel_rag.context` turns the
workbook into text for the context window and gives the model tools to read whatever that text left
out. It shares the reading and modelling layer (`excel_rag.workbook`) with ingestion, so an
attachment gets the same guarantees: macros are never run, external links never refreshed, a hostile
package is refused, and a formula is evaluated only when the `calculate` tool is asked to.

```python
from excel_rag.context import WorkbookSession, render_workbook

rendered = render_workbook(upload_bytes, name="budget.xlsx", token_budget=20_000, tools_hint=True)
rendered.text  # markdown to put in the prompt
rendered.complete  # False when anything was omitted (and the text says where)

session = WorkbookSession(rendered.model)
tools = session.tool_definitions()  # pass as `tools`
block = session.tool_result(tool_use.id, tool_use.name, tool_use.input)  # answer a tool_use
```

From the command line: `excel-rag render book.xlsx --budget 8000 --tools-hint`.
`examples/ask_workbook.py` is a complete question-answering loop with the Anthropic SDK.

## What the model reads

```text
# Workbook: budget.xlsx
3 sheet(s), 3 table(s), 1 formula cell(s), 0 named range(s).
Formula cells (ƒ) show the value Excel last saved; it is not recalculated here.

## Sheet "Forecast" - used range A1:B2
### Table "Metric" - Forecast!A1:B2, header row(s) 1
Columns: A Metric (text); B Value (formula)
| | A | B |
|---|---|---|
| 1 | Metric | Value |
| 2 | Revenue | 157.5 ƒ |
Formulas:
- B2: =SUM(Actuals!D2:D500)*(1+Assumptions!C7) → 157.5; reads Actuals!D2:D500, Assumptions!C7
```

- **Coordinates everywhere.** Every grid has row numbers and column letters, every region its A1
  range, so an answer can cite `Forecast!B2`.
- **Structure, not just cells.** Region detection supplies titles, header rows, units rows, notes,
  column types and Excel table names; formula clusters (a 50,000-row `=E2*F2` column) are one line.
- **Saved values, labelled.** A formula cell shows the value Excel last saved, marked `ƒ`. A file
  written by a library that saves no values (pandas, openpyxl) is called out at the top, and its
  formula cells show their formula rather than a blank that reads like zero.
- **Gaps are named.** `INDIRECT`, `OFFSET`, external links, undefined names and `#REF!` appear as
  unresolved references with their reason.

## Fitting the budget

The workbook is rendered at decreasing detail and the first rendering that fits is returned:

| detail | per table |
|---|---|
| `full` | every row and column |
| `rows-2000` … `rows-3` | header and units rows, the first N data rows, the last few, capped columns |
| `schemas` | header line and column schemas only |
| `overview` | sheets and their regions |
| `clipped` | the overview, cut at a line |

Every omission is written into the text (`| … | rows 52-491 omitted (440 rows with data) |`,
`(Columns M-GR omitted: 188 columns.)`, `(12 more formula(s) omitted.)`), and `tools_hint=True`
tells the model it can read them with the tools.

Tokens are **estimated** at three characters per token, which is conservative for grids of numbers
(prose runs nearer four). For an exact budget pass `count_tokens`, for example one backed by the
Messages API's `count_tokens` endpoint for the model you will call.

## The tools

| tool | answers |
|---|---|
| `read_range(sheet, range)` | a grid of the rectangle, with the formulas inside it; at most 2,000 cells, cut by rows with the next range named |
| `find(query)` | cells (value or formula) and named ranges containing the text; at most 50 matches, with the total |
| `precedents(sheet, range)` | the formulas in the range, their saved values and what they read, with single-cell precedents' values; a chart, pivot table or validation in the range, with what it reads |
| `calculate(sheet, range, changes)` | the range's values, after setting `changes` (cell → value as typed); with the `calc` extra — see below |
| `dependents(sheet, range)` | the formulas, charts, pivot tables and data validations anywhere that read any cell of the range — `D100` finds a formula reading `D2:D500` |

The definitions are plain dicts in the Messages API shape with `strict: true` and every field
required. A bad call (unknown sheet, malformed range, extra argument) is returned as a
`tool_result` with `is_error: true` and a message the model can correct from, never an exception.

## Calculating (what-if)

`excel_rag.workbook.calc` (the `calc` extra; `excel-rag calc`, the `calculate` tool) computes cells,
optionally after changing inputs. It never writes the file, and each call starts from the workbook
as saved.

- **Excel's values wherever they still hold.** Which formulas a change reaches is decided per cell
  from what each formula reads; every other formula keeps the value Excel saved (`saved`). Reached
  ones are recomputed (`recalculated`) and shown beside Excel's saved value. A workbook saved
  without values (written by a script) is computed throughout. `--check` / `check_saved_values`
  recomputes everything and lists where excel-rag and Excel disagree.
- **How a formula is computed.** Its references are resolved by this package — structured table
  references, defined names, `LET` variables, spill references to their saved extent — and the
  formula is compiled by [`formulas`](https://pypi.org/project/formulas/), which builds a function
  graph from the text (no generated Python is executed), with each distinct reference as a
  placeholder so a filled-down column compiles once. Plain arithmetic over cells (`+ - * / ^ %`)
  takes a faster path of its own with Excel's precedence and errors (`-2^2` is 4, `0^0` is `#NUM!`).
- **Unknown, with the reason.** `INDIRECT`, `OFFSET`, a user-defined function, a what-if data
  table, a 3-D reference, a function `formulas` lacks: these cannot be recomputed, so when a change
  might reach one (for `INDIRECT`/`OFFSET`, any change might) its value is `unknown`, and so is
  everything computed from it. Circular references are reported, not iterated. A volatile cell
  (`NOW`, `RAND`) is computed as of now and named.
- **Bounded.** 100,000 formula evaluations, 1,000,000 cells per range input, 30 s; a stop is
  reported and the cells it did not reach are unknown. A 50,000-row running balance recomputes in
  about 9 s; a change that reaches a handful of cells answers in well under a second.

```text
$ excel-rag calc model.xlsx "'Model Sheet'!A1:A5" --set Inputs!B2=20
With Inputs!B2 = 20:
- 'Model Sheet'!A1: 80 (recalculated; Excel saved 40)
- 'Model Sheet'!A2: 88 (recalculated; Excel saved 44)
- 'Model Sheet'!A3: 2026 (saved)
- 'Model Sheet'!A4: 160 (recalculated; Excel saved 80)
- 'Model Sheet'!A5: unknown (Excel last saved 11) -- it reads Inputs!B5, which is unknown (it uses
  INDIRECT(...): INDIRECT target depends on a runtime string; not statically resolvable)
3 formula evaluation(s) by excel-rag; 'saved' values are Excel's own, which no change reaches.
```

## Comparing two versions

`excel_rag.workbook.diff` (`excel-rag diff before.xlsx after.xlsx`, `--json` for the structure)
compares two files with the same reader: sheets added and removed, and every cell that changed,
classified as an input `value`, a `formula`, a `result` (same formula, different saved value — an
upstream input moved), a `kind` change (value ↔ formula), or `added`/`removed`. A changed formula
lists what it now reads and no longer reads; a changed input lists the formulas in the new version
that read it, so "the growth rate moved, and these cells depend on it" is one answer. Nothing is
recalculated. Named ranges whose target moved are listed too. The listing is capped
(`--max-changes`, 500), the per-sheet counts never are. Sheets are matched by name, so a renamed
sheet reads as removed plus added.

## Limits

- **Reading is the slow part.** openpyxl parses every cell's XML in pure Python. The saved values
  of formula cells now come from one streaming pass over the sheet XML instead of a second openpyxl
  load, which took a 50,000-row sheet with a formula column from 9.9 s to 5.7 s; rendering and
  every tool call stay well under that.
- **Region detection is heuristic.** A cell no detected region covers is still listed (“Other
  cells”), so nothing disappears, but an unusual layout may be split differently than a person
  would.
- **The whole workbook is held in memory** for the conversation, as the model the tools read from.
