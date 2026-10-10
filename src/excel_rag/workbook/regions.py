"""Split a worksheet into logical regions.

A worksheet is not one table. The detector works in two stages:

1. **Separate blocks by blank lines.** Rows that are empty across the used width, and columns empty
   across the used height, are boundaries. Their cross-product yields rectangular blocks, so two
   tables on one sheet -- separated by a blank row, a blank column, or both -- come out as two
   blocks. This is the acceptance case "two independent tables on one sheet".
2. **Classify each block.** Leading full-width banner rows are a *title*; the run of label-like rows
   under it is the *header* (a run of two is a merged multi-level header); a row of unit words after
   the header is a *units* row and is peeled off; the rest are *data* rows and trailing note-like
   rows become a *notes* region. The columns, their inferred types, distinct counts, examples and
   units are recorded per column, and the data rows are cut into a bounded number of row groups.

Nothing here evaluates a formula or trusts a declared dimension: it only reads the populated cells
the reader already extracted.
"""

from __future__ import annotations

import math
import re
from collections.abc import Mapping, Sequence
from dataclasses import dataclass

from ..models import A1Range
from .canonical import (
    CellValue,
    ColumnSchema,
    Region,
    RegionKind,
    RowGroup,
    column_letter,
    column_node_id,
    format_value,
    region_node_id,
    row_group_node_id,
    table_node_id,
)
from .reader import RawSheet

#: Unit words that mark a "units" row under a header (all lowercase, punctuation stripped).
_UNITS = frozenset(
    {
        "usd",
        "eur",
        "gbp",
        "jpy",
        "cad",
        "aud",
        "chf",
        "rmb",
        "cny",
        "mxn",
        "dollars",
        "euros",
        "pounds",
        "us dollars",
        "usd 000s",
        "euros 000s",
        "$",
        "€",
        "£",
        "¥",
        "₹",
        "%",
        "percent",
        "percentage",
        "bps",
        "pp",
        "kg",
        "g",
        "mg",
        "lb",
        "lbs",
        "oz",
        "t",
        "tonne",
        "tonnes",
        "mm",
        "cm",
        "m",
        "km",
        "in",
        "ft",
        "yd",
        "mi",
        "ea",
        "each",
        "pcs",
        "pieces",
        "units",
        "unit",
        "u",
        "count",
        "qty",
        "quantity",
        "hours",
        "hrs",
        "hr",
        "day",
        "days",
        "week",
        "weeks",
        "month",
        "months",
        "year",
        "years",
        "k",
        "bn",
        "thousand",
        "thousands",
        "million",
        "millions",
        "billion",
        "in thousands",
        "in millions",
        "in billions",
        "000s",
        "000",
        "per unit",
        "pct",
    }
)

_NOTE_RE = re.compile(
    r"(?i)^\s*(note|notes|assumption|assumptions|source|sources|legend|disclaimer|warning|"
    r"footnote|footnotes|data as of|as of|prepared by|methodology|commentary|definition|"
    r"definitions)\b\s*:?"
)


@dataclass(frozen=True, slots=True)
class RegionConfig:
    """Bounds the document count a region can produce, independent of how many rows it has."""

    row_group_rows: int = 50
    max_row_groups_per_region: int = 8
    max_header_rows: int = 3
    max_examples: int = 5


def _present(cells: list[CellValue]) -> list[CellValue]:
    return [cell for cell in cells if cell.value is not None or cell.computed]


def _bands(present: set[int], low: int, high: int) -> list[tuple[int, int]]:
    bands: list[tuple[int, int]] = []
    start: int | None = None
    for value in range(low, high + 1):
        if value in present:
            if start is None:
                start = value
        elif start is not None:
            bands.append((start, value - 1))
            start = None
    if start is not None:
        bands.append((start, high))
    return bands


def _normalise_unit(text: str) -> str:
    return text.strip().strip(".:").lower()


def _merged_width(merged_range: str | None) -> int:
    if not merged_range:
        return 1
    try:
        parts = merged_range.replace("$", "").split(":")
        start = A1Range.parse("x", parts[0])
        end = A1Range.parse("x", parts[1]) if len(parts) == 2 else start
    except ValueError:
        return 1
    return end.max_col - start.min_col + 1


def _is_banner(row_cells: list[CellValue], width: int) -> bool:
    present = _present(row_cells)
    if len(present) != 1:
        return False
    cell = present[0]
    if not isinstance(cell.value, str) or cell.merged_range is None:
        return False
    return _merged_width(cell.merged_range) >= width


def _is_data_cell(cell: CellValue) -> bool:
    return cell.data_type in ("n", "b", "f") or cell.is_date


def _is_label_like(row_cells: list[CellValue]) -> bool:
    present = _present(row_cells)
    if not present:
        return False
    data = sum(1 for cell in present if _is_data_cell(cell))
    return data * 2 < len(present)


def _is_units_row(row_cells: list[CellValue]) -> bool:
    present = _present(row_cells)
    if not present:
        return False
    texts = []
    for cell in present:
        if not isinstance(cell.value, str):
            return False
        texts.append(_normalise_unit(cell.value))
    return all(text in _UNITS for text in texts if text)


def _is_note_row(row_cells: list[CellValue]) -> bool:
    present = _present(row_cells)
    if not present:
        return False
    first = min(present, key=lambda cell: cell.column)
    if not isinstance(first.value, str):
        return False
    return bool(_NOTE_RE.match(first.value))


def _infer_type(cells: list[CellValue]) -> str:
    present = _present(cells)
    if not present:
        return "empty"
    kinds: set[str] = set()
    for cell in present:
        if cell.computed:
            kinds.add("formula")
        elif cell.is_date:
            kinds.add("date")
        elif cell.data_type == "n":
            kinds.add("number")
        elif cell.data_type == "b":
            kinds.add("boolean")
        elif cell.data_type == "e":
            kinds.add("error")
        else:
            kinds.add("text")
    if kinds <= {"number"} and any(cell.is_percentage for cell in present):
        return "percentage"
    if kinds <= {"number"} and any(cell.is_currency for cell in present):
        return "currency"
    if kinds == {"number"} or kinds == {"number", "date"}:
        return "number"
    if kinds == {"date"}:
        return "date"
    if kinds == {"text"}:
        return "text"
    if kinds == {"boolean"}:
        return "boolean"
    if kinds == {"formula"}:
        return "formula"
    if kinds == {"error"}:
        return "error"
    return "mixed"


def _cell_display(cell: CellValue) -> str:
    if cell.computed:
        return format_value(cell.cached_value)
    return format_value(cell.value)


def _column_range(sheet: str, index: int, rows: list[int], fallback_row: int) -> A1Range:
    letter = column_letter(index)
    if rows:
        return A1Range.parse(sheet, f"{letter}{min(rows)}:{letter}{max(rows)}")
    return A1Range.parse(sheet, f"{letter}{fallback_row}")


def _chunks(rows: list[int], size: int) -> list[list[int]]:
    return [rows[start : start + size] for start in range(0, len(rows), size)]


def _merge_cover(
    merged_ranges: Sequence[str],
) -> dict[tuple[int, int], tuple[int, int]]:
    """Map every coordinate covered by a merged range to that range's top-left (master) cell."""
    cover: dict[tuple[int, int], tuple[int, int]] = {}
    for reference in merged_ranges:
        parts = reference.replace("$", "").split(":")
        try:
            start = A1Range.parse("x", parts[0])
            end = A1Range.parse("x", parts[1]) if len(parts) == 2 else start
        except ValueError:
            continue
        if (end.max_row - start.min_row + 1) * (end.max_col - start.min_col + 1) > 100_000:
            continue
        master = (start.min_row, start.min_col)
        for row in range(start.min_row, end.max_row + 1):
            for column in range(start.min_col, end.max_col + 1):
                cover[(row, column)] = master
    return cover


def _build_columns(
    by_row: Mapping[int, dict[int, CellValue]],
    header_rows: list[int],
    units_row: int | None,
    data_rows: list[int],
    col_band: tuple[int, int],
    workbook_id: str,
    version: int,
    sheet_name: str,
    config: RegionConfig,
    cover: Mapping[tuple[int, int], tuple[int, int]],
) -> list[ColumnSchema]:
    columns: list[ColumnSchema] = []
    for index in range(col_band[0], col_band[1] + 1):
        header_cells: list[CellValue] = []
        for row in header_rows:
            cell = by_row[row].get(index)
            if cell is None:
                master = cover.get((row, index))
                if master is not None:
                    cell = by_row.get(master[0], {}).get(master[1])
            if cell is not None:
                header_cells.append(cell)
        data_cells = [by_row[row][index] for row in data_rows if index in by_row[row]]
        if not header_cells and not data_cells:
            continue
        header_texts = tuple(
            text
            for cell in header_cells
            if isinstance(cell.value, str) and (text := cell.value.strip())
        )
        name = " / ".join(header_texts) if header_texts else f"Column {column_letter(index)}"
        letter = column_letter(index)
        present_rows = [cell.row for cell in data_cells if cell.value is not None or cell.computed]
        a1 = _column_range(
            sheet_name, index, present_rows, header_rows[0] if header_rows else col_band[0]
        )
        distinct: list[str] = []
        seen: set[str] = set()
        for cell in data_cells:
            display = _cell_display(cell)
            if display and display not in seen:
                seen.add(display)
                if len(distinct) < config.max_examples:
                    distinct.append(display)
        unit = None
        if units_row is not None and index in by_row.get(units_row, {}):
            unit_value = by_row[units_row][index].value
            unit = unit_value if isinstance(unit_value, str) else None
        header_coordinate = (
            header_cells[-1].coordinate
            if header_cells
            else f"{letter}{header_rows[0] if header_rows else col_band[0]}"
        )
        columns.append(
            ColumnSchema(
                index=index,
                letter=letter,
                name=name,
                inferred_type=_infer_type(data_cells),
                aliases=header_texts,
                distinct_count=len(seen),
                examples=tuple(distinct),
                unit=unit,
                a1_range=a1,
                node_id=column_node_id(workbook_id, version, sheet_name, a1.a1),
                header_coordinate=header_coordinate,
            )
        )
    return columns


def _build_row_groups(
    by_row: Mapping[int, dict[int, CellValue]],
    data_rows: list[int],
    columns: list[ColumnSchema],
    col_band: tuple[int, int],
    workbook_id: str,
    version: int,
    sheet_name: str,
    config: RegionConfig,
) -> list[RowGroup]:
    if not data_rows or not columns:
        return []
    size = max(config.row_group_rows, math.ceil(len(data_rows) / config.max_row_groups_per_region))
    groups: list[RowGroup] = []
    headers = tuple(column.name for column in columns)
    for group_index, chunk in enumerate(_chunks(data_rows, size)):
        a1 = A1Range.parse(
            sheet_name,
            f"{column_letter(col_band[0])}{chunk[0]}:{column_letter(col_band[1])}{chunk[-1]}",
        )
        rows: list[tuple[str, ...]] = []
        for row in chunk:
            row_cells = by_row.get(row, {})
            rows.append(
                tuple(
                    _cell_display(row_cells[index]) if index in row_cells else ""
                    for index in (column.index for column in columns)
                )
            )
        groups.append(
            RowGroup(
                index=group_index,
                a1_range=a1,
                node_id=row_group_node_id(workbook_id, version, sheet_name, a1.a1),
                headers=headers,
                rows=tuple(rows),
            )
        )
    return groups


def _match_table(sheet: RawSheet, region: A1Range) -> tuple[str, tuple[str, ...]] | None:
    for table in sheet.tables:
        if not table.ref:
            continue
        try:
            table_range = A1Range.parse(sheet.name, table.ref)
        except ValueError:
            continue
        if (
            table_range.min_row >= region.min_row
            and table_range.max_row <= region.max_row
            and table_range.min_col >= region.min_col
            and table_range.max_col <= region.max_col
        ):
            return table.name, table.columns
    return None


def _table_region(
    sheet: RawSheet,
    by_row: Mapping[int, dict[int, CellValue]],
    header_rows: list[int],
    units_row: int | None,
    data_rows: list[int],
    notes: list[str],
    col_band: tuple[int, int],
    workbook_id: str,
    version: int,
    config: RegionConfig,
) -> Region:
    rows = [*header_rows, *([units_row] if units_row is not None else []), *data_rows]
    min_row = min(rows) if rows else col_band[0]
    max_row = max(rows) if rows else min_row
    a1 = A1Range.parse(
        sheet.name,
        f"{column_letter(col_band[0])}{min_row}:{column_letter(col_band[1])}{max_row}",
    )
    columns = _build_columns(
        by_row,
        header_rows,
        units_row,
        data_rows,
        col_band,
        workbook_id,
        version,
        sheet.name,
        config,
        _merge_cover(sheet.merged_ranges),
    )
    row_groups = _build_row_groups(
        by_row, data_rows, columns, col_band, workbook_id, version, sheet.name, config
    )
    matched = _match_table(sheet, a1) if sheet.tables else None
    table_name: str | None = None
    node_type_kind = "region"
    node = region_node_id(workbook_id, version, sheet.name, a1.a1)
    if matched is not None:
        table_name = matched[0]
        node_type_kind = "table"
        node = table_node_id(workbook_id, version, table_name)
    title = None
    if header_rows:
        first_header = by_row[header_rows[0]]
        if first_header:
            leftmost = min(first_header.values(), key=lambda cell: cell.column)
            if isinstance(leftmost.value, str):
                title = leftmost.value
    return Region(
        kind=RegionKind.TABLE,
        sheet_name=sheet.name,
        a1_range=a1,
        node_id=node,
        node_type_kind=node_type_kind,
        title=title,
        table_name=table_name,
        header_rows=tuple(header_rows),
        units_row=units_row,
        columns=tuple(columns),
        row_groups=tuple(row_groups),
        notes=tuple(notes),
    )


def _simple_region(
    sheet: RawSheet,
    rows: list[int],
    col_band: tuple[int, int],
    kind: RegionKind,
    title: str | None,
    notes: tuple[str, ...],
    workbook_id: str,
    version: int,
) -> Region:
    a1 = A1Range.parse(
        sheet.name,
        f"{column_letter(col_band[0])}{min(rows)}:{column_letter(col_band[1])}{max(rows)}",
    )
    return Region(
        kind=kind,
        sheet_name=sheet.name,
        a1_range=a1,
        node_id=region_node_id(workbook_id, version, sheet.name, a1.a1),
        node_type_kind="region",
        title=title,
        table_name=None,
        header_rows=(),
        units_row=None,
        columns=(),
        row_groups=(),
        notes=notes,
    )


def _analyze_block(
    sheet: RawSheet,
    block_cells: list[CellValue],
    row_band: tuple[int, int],
    col_band: tuple[int, int],
    workbook_id: str,
    version: int,
    config: RegionConfig,
) -> list[Region]:
    by_row: dict[int, dict[int, CellValue]] = {}
    for cell in block_cells:
        by_row.setdefault(cell.row, {})[cell.column] = cell
    rows = sorted(by_row)
    width = col_band[1] - col_band[0] + 1
    regions: list[Region] = []

    index = 0
    title_rows: list[int] = []
    while index < len(rows) and _is_banner(list(by_row[rows[index]].values()), width):
        title_rows.append(rows[index])
        index += 1
    if title_rows:
        first_cell = next(iter(by_row[title_rows[0]].values()))
        title_text = first_cell.value if isinstance(first_cell.value, str) else None
        regions.append(
            _simple_region(
                sheet, title_rows, col_band, RegionKind.TITLE, title_text, (), workbook_id, version
            )
        )

    remaining = rows[index:]
    has_numeric = any(_is_data_cell(cell) for cell in block_cells)
    if remaining and not has_numeric:
        # No numeric/date values at all: this is prose -- a title, a notes block, or a free region.
        first_row_cells = by_row[remaining[0]]
        first_cell = min(first_row_cells.values(), key=lambda cell: cell.column)
        title_text = first_cell.value if isinstance(first_cell.value, str) else None
        if len(remaining) == 1 and title_text:
            kind = RegionKind.TITLE
        elif any(_is_note_row(list(by_row[row].values())) for row in remaining):
            kind = RegionKind.NOTES
        else:
            kind = RegionKind.GENERIC
        note_texts = (
            tuple(
                str(min(by_row[row].values(), key=lambda cell: cell.column).value)
                for row in remaining
                if by_row[row]
            )
            if kind is RegionKind.NOTES
            else ()
        )
        regions.append(
            _simple_region(
                sheet,
                remaining,
                col_band,
                kind,
                title_text if kind is RegionKind.TITLE else None,
                note_texts,
                workbook_id,
                version,
            )
        )
        return regions

    header_rows: list[int] = []
    while (
        index < len(rows)
        and len(header_rows) < config.max_header_rows
        and _is_label_like(list(by_row[rows[index]].values()))
    ):
        header_rows.append(rows[index])
        index += 1
    units_row: int | None = None
    if len(header_rows) >= 2 and _is_units_row(list(by_row[header_rows[-1]].values())):
        units_row = header_rows.pop()

    data_rows = list(rows[index:])
    note_rows: list[int] = []
    while data_rows and _is_note_row(list(by_row[data_rows[-1]].values())):
        note_rows.insert(0, data_rows.pop())

    if not data_rows:
        # A header with no data below it is not a table; keep it as a bare region/title.
        if header_rows:
            first_cell = min(by_row[header_rows[0]].values(), key=lambda cell: cell.column)
            title_text = first_cell.value if isinstance(first_cell.value, str) else None
            regions.append(
                _simple_region(
                    sheet,
                    header_rows,
                    col_band,
                    RegionKind.TITLE if title_text else RegionKind.GENERIC,
                    title_text,
                    (),
                    workbook_id,
                    version,
                )
            )
        return regions

    note_texts = tuple(
        str(next(iter(by_row[row].values())).value) for row in note_rows if by_row[row]
    )
    regions.append(
        _table_region(
            sheet,
            by_row,
            header_rows,
            units_row,
            data_rows,
            list(note_texts),
            col_band,
            workbook_id,
            version,
            config,
        )
    )
    if note_rows:
        regions.append(
            _simple_region(
                sheet, note_rows, col_band, RegionKind.NOTES, None, note_texts, workbook_id, version
            )
        )
    return regions


def detect_regions(
    sheet: RawSheet,
    *,
    workbook_id: str,
    version: int,
    config: RegionConfig | None = None,
) -> tuple[Region, ...]:
    """Detect every logical region on ``sheet``. Macro sheets have no regions by definition."""
    if sheet.is_macro_sheet or not sheet.cells:
        return ()
    settings = config or RegionConfig()
    cells = sheet.cells
    min_row = min(cell.row for cell in cells.values())
    max_row = max(cell.row for cell in cells.values())
    min_col = min(cell.column for cell in cells.values())
    max_col = max(cell.column for cell in cells.values())
    row_bands = _bands({cell.row for cell in cells.values()}, min_row, max_row)
    col_bands = _bands({cell.column for cell in cells.values()}, min_col, max_col)
    regions: list[Region] = []
    for row_band in row_bands:
        for col_band in col_bands:
            block = [
                cell
                for cell in cells.values()
                if row_band[0] <= cell.row <= row_band[1]
                and col_band[0] <= cell.column <= col_band[1]
            ]
            if not block:
                continue
            regions.extend(
                _analyze_block(sheet, block, row_band, col_band, workbook_id, version, settings)
            )
    return tuple(regions)
