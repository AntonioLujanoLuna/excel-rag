"""Render a workbook as text for a conversation's context window, within a token budget.

The rendering is built for a model to read and cite, not for search:

* every table is a grid with **row numbers and column letters**, so an answer can name ``B7``;
* every region says what it is -- a titled table, a notes block, a headerless grid -- with its A1
  range, header row, units and notes, as region detection found them;
* formula cells show the value **Excel last saved** (marked ``ƒ``), and a per-sheet list gives each
  formula, that saved value, and the ranges it reads. Nothing is recalculated; a workbook saved by a
  tool that writes no cached values says so instead of showing blanks as zeros;
* anything left out is **stated** -- ``rows 23-4,980 omitted`` -- never dropped silently.

The budget is met by rendering at decreasing detail (:data:`DETAIL_LADDER`): everything, then the
first and last rows of each table, then schemas only, then an overview. The first level whose text
fits is returned. Tokens are estimated by default (characters / 3, conservative for grids of
numbers); pass ``count_tokens`` to use an exact counter instead.
"""

from __future__ import annotations

import math
from collections.abc import Callable
from dataclasses import dataclass
from datetime import date, datetime, time
from pathlib import Path
from typing import Any

from ..models import A1Range
from ..workbook import RegionConfig, WorkbookModel, load_workbook
from ..workbook.canonical import (
    CellValue,
    FormulaEntry,
    Region,
    RegionKind,
    SheetModel,
    column_letter,
)

#: Marks a formula cell in a grid; its formula is listed under the sheet's "Formulas".
FORMULA_MARK = "ƒ"

TokenCounter = Callable[[str], int]


def approx_tokens(text: str) -> int:
    """A conservative token estimate: one token per three characters.

    Prose runs nearer four characters per token; grids of numbers, coordinates and separators run
    nearer three, so three keeps a rendering inside its budget. For an exact count, pass a counter
    backed by the model's own tokenizer (for Claude, ``messages.count_tokens``).
    """
    return math.ceil(len(text) / 3)


@dataclass(frozen=True)
class Detail:
    """One rung of the ladder: how much of each table, formula list and sheet to render."""

    name: str
    head_rows: int | None  # data rows kept from the top of a table; ``None``: all
    tail_rows: int  # data rows kept from the bottom when the middle is omitted
    max_columns: int | None
    max_formulas: int | None  # per sheet
    cell_chars: int | None
    grids: bool = True
    schemas: bool = True
    other_cells: int | None = None  # cells outside any region, per sheet; ``None``: all


DETAIL_LADDER: tuple[Detail, ...] = (
    Detail("full", None, 0, None, None, None),
    Detail("rows-2000", 2000, 100, 100, 1000, 200, other_cells=1000),
    Detail("rows-1000", 1000, 50, 80, 600, 160, other_cells=500),
    Detail("rows-500", 500, 25, 60, 400, 120, other_cells=300),
    Detail("rows-200", 200, 20, 60, 300, 120, other_cells=200),
    Detail("rows-50", 50, 10, 40, 100, 80, other_cells=50),
    Detail("rows-20", 20, 5, 25, 40, 60, other_cells=20),
    Detail("rows-8", 8, 3, 15, 15, 40, other_cells=8),
    Detail("rows-3", 3, 1, 10, 6, 30, other_cells=3),
    Detail("schemas", 0, 0, 10, 3, 30, grids=False, other_cells=0),
    Detail("overview", 0, 0, 0, 0, 30, grids=False, schemas=False, other_cells=0),
)


@dataclass(frozen=True)
class RenderedWorkbook:
    """The text to put in context, what it cost, and whether anything was left out."""

    text: str
    tokens: int
    detail: str
    complete: bool
    model: WorkbookModel


def render_workbook(
    source: str | Path | bytes | WorkbookModel,
    *,
    token_budget: int = 8_000,
    name: str | None = None,
    count_tokens: TokenCounter = approx_tokens,
    tools_hint: bool = False,
    config: RegionConfig | None = None,
) -> RenderedWorkbook:
    """Render ``source`` (a path, an upload's bytes, or an already-loaded model) for a prompt.

    Returns the most detailed rendering that fits ``token_budget``. ``tools_hint`` adds a line
    telling the model that omitted ranges can be read with the workbook tools
    (:mod:`excel_rag.context.tools`) -- set it when those tools are offered in the same request.
    """
    if token_budget < 50:
        raise ValueError("a token budget under 50 cannot hold even the workbook overview")
    model = (
        source
        if isinstance(source, WorkbookModel)
        else load_workbook(source, name=name, config=config)
    )
    cells = sum(len(sheet.cells) for sheet in model.sheets)
    for detail in DETAIL_LADDER:
        # A full grid costs at least a couple of tokens per cell; skip rendering what cannot fit.
        if detail.head_rows is None and cells * 2 > token_budget:
            continue
        text, complete = _render(model, detail, tools_hint=tools_hint)
        tokens = count_tokens(text)
        if tokens <= token_budget:
            return RenderedWorkbook(text, tokens, detail.name, complete, model)
    text = _clip(text, token_budget, count_tokens)
    return RenderedWorkbook(text, count_tokens(text), "clipped", False, model)


# -------------------------------------------------------------------------------------------------
# Rendering
# -------------------------------------------------------------------------------------------------
class _Writer:
    def __init__(self) -> None:
        self.lines: list[str] = []
        self.omitted = False

    def add(self, line: str = "") -> None:
        self.lines.append(line)

    def omit(self, line: str) -> None:
        self.omitted = True
        self.lines.append(line)


def _render(model: WorkbookModel, detail: Detail, *, tools_hint: bool) -> tuple[str, bool]:
    out = _Writer()
    _header(out, model)
    body = _Writer()
    for sheet in model.sheets:
        _sheet(body, sheet, detail)
    if body.omitted:
        hint = (
            " Use the workbook tools (read_range, find, precedents, dependents) to read them."
            if tools_hint
            else ""
        )
        out.add(f"This rendering is partial (detail: {detail.name}); omissions are marked.{hint}")
    out.add()
    out.lines.extend(body.lines)
    return "\n".join(out.lines).rstrip() + "\n", not body.omitted


def _header(out: _Writer, model: WorkbookModel) -> None:
    sheets = model.sheets
    tables = sum(
        1 for sheet in sheets for region in sheet.regions if region.kind is RegionKind.TABLE
    )
    formulas = sum(len(entry.member_coordinates) for sheet in sheets for entry in sheet.formulas)
    out.add(f"# Workbook: {model.source_file}")
    out.add(
        f"{len(sheets)} sheet(s), {tables} table(s), {formulas} formula cell(s), "
        f"{len(model.named_ranges)} named range(s)."
    )
    if formulas:
        saved = [
            entry.cached_value is not None
            for sheet in sheets
            for entry in sheet.formulas
            if not entry.is_cluster
        ]
        if saved and not any(saved):
            out.add(
                "Formula cells carry no saved values (the file was not last saved by Excel), "
                "so they show their formula instead of a result. Nothing is recalculated."
            )
        else:
            out.add(
                f"Formula cells ({FORMULA_MARK}) show the value Excel last saved; it is not "
                "recalculated here."
            )
    if model.has_vba or model.macro_sheet_names:
        out.add("Contains macros; they were not loaded or run.")
    for warning in model.warnings:
        out.add(f"Note: {warning}")
    if model.named_ranges:
        out.add("Named ranges: " + "; ".join(_named(name) for name in model.named_ranges))


def _named(named: Any) -> str:
    if named.resolved and named.sheet_name and named.a1:
        return f"{named.name} = {_qualified(named.sheet_name, named.a1)}"
    return f"{named.name} (not a static range: {named.detail or 'unresolved'})"


def _sheet(out: _Writer, sheet: SheetModel, detail: Detail) -> None:
    if sheet.is_macro_sheet:
        out.add(f"## Sheet {_quote(sheet.name)} - macro sheet, not loaded")
        out.add()
        return
    used = sheet.a1_range.a1 if sheet.a1_range else "empty"
    hidden = "" if sheet.visibility == "visible" else f", {sheet.visibility}"
    out.add(f"## Sheet {_quote(sheet.name)} - used range {used}{hidden}")
    if not sheet.cells:
        out.add()
        return
    if not detail.schemas:
        summary = ", ".join(_region_label(region) for region in sheet.regions) or "no regions"
        out.omit(f"Regions: {summary}. (Contents omitted.)")
        out.add()
        return
    for region in sheet.regions:
        _region(out, sheet, region, detail)
    _other_cells(out, sheet, detail)
    _formulas(out, sheet, detail)
    out.add()


def _region_label(region: Region) -> str:
    title = f" {_quote(region.title)}" if region.title else ""
    return f"{region.kind.value}{title} {region.a1_range.a1}"


def _region(out: _Writer, sheet: SheetModel, region: Region, detail: Detail) -> None:
    where = _qualified(sheet.name, region.a1_range.a1)
    if region.kind is RegionKind.TITLE:
        out.add(f"### Title {_quote(region.title or '')} - {where}")
        return
    if region.kind is RegionKind.NOTES:
        out.add(f"### Notes - {where}")
        for note in region.notes:
            out.add(f"> {_cell_text(note, detail.cell_chars)}")
        return

    label = "Table" if region.kind is RegionKind.TABLE else "Grid (no header detected)"
    named = f" {_quote(region.title)}" if region.title else ""
    excel_table = f", Excel table {_quote(region.table_name)}" if region.table_name else ""
    headers = (
        f", header row(s) {', '.join(str(row) for row in region.header_rows)}"
        if region.header_rows
        else ""
    )
    out.add(f"### {label}{named} - {where}{excel_table}{headers}")
    if region.columns:
        out.add("Columns: " + "; ".join(_column(column) for column in region.columns))
    for note in region.notes:
        out.add(f"Note: {_cell_text(note, detail.cell_chars)}")
    if detail.grids:
        _grid(out, sheet, region, detail)
    else:
        out.omit(f"(Rows omitted: {region.a1_range.a1}.)")


def _column(column: Any) -> str:
    unit = f", {column.unit}" if column.unit else ""
    return f"{column.letter} {column.name} ({column.inferred_type}{unit})"


def _grid(out: _Writer, sheet: SheetModel, region: Region, detail: Detail) -> None:
    bounds = region.a1_range
    pinned = set(region.header_rows) | ({region.units_row} if region.units_row else set())
    rows = [
        row
        for row in range(bounds.min_row, bounds.max_row + 1)
        if any(
            _coordinate(column, row) in sheet.cells
            for column in range(bounds.min_col, bounds.max_col + 1)
        )
    ]
    data = [row for row in rows if row not in pinned]
    kept_data, gap = _window(data, detail)
    columns = list(range(bounds.min_col, bounds.max_col + 1))
    dropped_columns: list[int] = []
    if detail.max_columns is not None and len(columns) > detail.max_columns:
        columns, dropped_columns = columns[: detail.max_columns], columns[detail.max_columns :]

    out.add("| | " + " | ".join(column_letter(column) for column in columns) + " |")
    out.add("|---|" + "---|" * len(columns))
    shown = sorted(set(kept_data) | (pinned & set(rows)))
    marked = False
    for row in shown:
        if gap is not None and not marked and row > gap[1]:
            _gap_row(out, gap)
            marked = True
        cells = [_display(sheet.cells.get(_coordinate(column, row)), detail) for column in columns]
        out.add(f"| {row} | " + " | ".join(cells) + " |")
    if gap is not None and not marked:
        _gap_row(out, gap)
    if dropped_columns:
        first, last = column_letter(dropped_columns[0]), column_letter(dropped_columns[-1])
        out.omit(f"(Columns {first}-{last} omitted: {len(dropped_columns)} columns.)")


def _gap_row(out: _Writer, gap: tuple[int, int, int]) -> None:
    first, last, count = gap
    out.omit(f"| … | rows {first}-{last} omitted ({count} rows with data) |")


def _window(data: list[int], detail: Detail) -> tuple[list[int], tuple[int, int, int] | None]:
    """Keep the first ``head_rows`` and last ``tail_rows`` data rows; describe the gap."""
    head = detail.head_rows
    if head is None or len(data) <= head + detail.tail_rows:
        return data, None
    kept = data[:head] + (data[-detail.tail_rows :] if detail.tail_rows else [])
    middle = data[head : len(data) - detail.tail_rows]
    return kept, (middle[0], middle[-1], len(middle))


def _other_cells(out: _Writer, sheet: SheetModel, detail: Detail) -> None:
    """Populated cells no region covers: listed, so region detection never hides a value."""
    loose = [
        cell
        for cell in sorted(sheet.cells.values(), key=lambda item: (item.row, item.column))
        if not any(_inside(cell, region.a1_range) for region in sheet.regions)
        and _display(cell, detail)
    ]
    if not loose:
        return
    limit = detail.other_cells
    shown = loose if limit is None else loose[:limit]
    if shown:
        out.add(
            "Other cells: "
            + "; ".join(f"{cell.coordinate} = {_display(cell, detail)}" for cell in shown)
        )
    if len(shown) < len(loose):
        out.omit(f"({len(loose) - len(shown)} other cell(s) omitted.)")


def _formulas(out: _Writer, sheet: SheetModel, detail: Detail) -> None:
    if not sheet.formulas:
        return
    limit = detail.max_formulas
    entries = list(sheet.formulas) if limit is None else list(sheet.formulas)[:limit]
    if entries:
        out.add("Formulas:")
        for entry in entries:
            limit = None if detail.cell_chars is None else detail.cell_chars * 3
            out.add(f"- {formula_line(entry, max_chars=limit)}")
    if len(entries) < len(sheet.formulas):
        out.omit(f"({len(sheet.formulas) - len(entries)} more formula(s) omitted.)")


def formula_line(
    entry: FormulaEntry, *, qualified: bool = False, max_chars: int | None = None
) -> str:
    """One formula (or a cluster of one repeated pattern): where, text, saved value, reads.

    ``qualified`` prefixes the sheet, for lists that span sheets; ``max_chars`` shortens a long
    formula's text.
    """
    text = entry.formula if entry.formula.startswith("=") else f"={entry.formula}"
    text = _cell_text(text, max_chars)
    where = _qualified(entry.sheet_name, entry.a1_range.a1) if qualified else entry.a1_range.a1
    if entry.is_cluster:
        where += f" ({len(entry.member_coordinates)} cells, one pattern)"
        text = f"{text} (as in {entry.member_coordinates[0]})"
        result = "values in the grid" if entry.cached_value is not None else "no saved values"
    else:
        result = (
            _value(entry.cached_value, None) if entry.cached_value is not None else "no saved value"
        )
    line = f"{where}: {text} → {result}"
    reads = ", ".join(
        dict.fromkeys(
            _qualified(reference.sheet_name, reference.a1_range) for reference in entry.references
        )
    )
    if reads:
        line += f"; reads {reads}"
    if entry.unresolved_references:
        gaps = ", ".join(
            dict.fromkeys(
                f"{gap.reference_text} ({gap.reason.value})" for gap in entry.unresolved_references
            )
        )
        line += f"; unresolved: {gaps}"
    return line


# -------------------------------------------------------------------------------------------------
# Values
# -------------------------------------------------------------------------------------------------
def _display(cell: CellValue | None, detail: Detail) -> str:
    """A cell as it should read in a grid: the saved value for a formula, marked ``ƒ``."""
    if cell is None:
        return ""
    if cell.formula:
        if cell.cached_value is None:
            return _cell_text(cell.formula, detail.cell_chars)
        return f"{_value(cell.cached_value, cell, detail.cell_chars)} {FORMULA_MARK}"
    return _value(cell.value, cell, detail.cell_chars)


def _value(value: Any, cell: CellValue | None, limit: int | None = None) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, datetime):
        text = value.date().isoformat() if value.time() == time(0) else value.isoformat(" ")
        return text
    if isinstance(value, date | time):
        return value.isoformat()
    if isinstance(value, int | float):
        if cell is not None and cell.is_percentage:
            return f"{_number(value * 100)}%"
        return _number(value)
    return _cell_text(str(value), limit)


def _number(value: float) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return format(value, ".12g") if isinstance(value, float) else str(value)


def _cell_text(text: str, limit: int | None) -> str:
    flat = " ".join(text.split()).replace("|", "\\|")
    if limit is not None and len(flat) > limit:
        return flat[: limit - 1] + "…"
    return flat


# -------------------------------------------------------------------------------------------------
# Helpers
# -------------------------------------------------------------------------------------------------
def _coordinate(column: int, row: int) -> str:
    return f"{column_letter(column)}{row}"


def _inside(cell: CellValue, bounds: A1Range) -> bool:
    return bounds.contains_row(cell.row) and bounds.contains_column(cell.column)


def _quote(name: str) -> str:
    return f'"{name}"'


def _qualified(sheet: str, a1: str) -> str:
    plain = sheet.replace("_", "").isalnum()
    return f"{sheet}!{a1}" if plain else f"'{sheet.replace(chr(39), chr(39) * 2)}'!{a1}"


def _clip(text: str, budget: int, count_tokens: TokenCounter) -> str:
    """Last resort for a workbook whose overview alone exceeds the budget: cut at a line."""
    marker = "\n(Rendering clipped to the token budget.)\n"
    lines = text.splitlines()
    kept: list[str] = []
    for line in lines:
        candidate = "\n".join([*kept, line]) + marker
        if count_tokens(candidate) > budget:
            break
        kept.append(line)
    return "\n".join(kept) + marker


__all__ = [
    "DETAIL_LADDER",
    "FORMULA_MARK",
    "Detail",
    "RenderedWorkbook",
    "TokenCounter",
    "approx_tokens",
    "formula_line",
    "render_workbook",
]
