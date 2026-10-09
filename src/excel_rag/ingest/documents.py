"""Turn the canonical model into :class:`ChunkDocument` and :class:`StructureDocument` records.

The split the design insists on: ``excel_chunks`` carries what is *searchable* (prose that includes
the sheet and the A1 range), ``excel_structure`` carries what is *exact* (coordinates, spans, the
formula, the cached value, typed edges). Nothing here recomputes a formula: a formula's
``cached_value`` is passed through verbatim and labelled as the last-saved value.

**Bounded by construction.** A region contributes at most a fixed number of documents regardless of
how many rows it has: one region node, one node per column, and at most
``MAX_ROW_GROUPS_PER_REGION`` row-group nodes. Only *important* cells -- headers, formula cells
(and, for a repeated formula, a single cluster node), named-range anchors and referenced targets --
get their own structure node, so a 500x6 grid produces dozens of documents, not thousands.
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from datetime import UTC, date, datetime, time
from typing import Any

from ..models import (
    A1Range,
    ChunkDocument,
    ChunkType,
    NodeType,
    Reference,
    ReferenceKind,
    StructureDocument,
)
from .canonical import (
    ColumnSchema,
    FormulaEntry,
    Region,
    RegionKind,
    RowGroup,
    SheetModel,
    WorkbookModel,
    cell_node_id,
    chunk_id,
    format_value,
    workbook_node_id,
)
from .reader import column_index

#: Ideal rows per row-group chunk. A larger region gets coarser groups so the group count is capped.
ROW_GROUP_ROWS = 50
#: Hard ceiling on row groups for one region, whatever its height.
MAX_ROW_GROUPS_PER_REGION = 8
#: At most this many column=value pairs are written into one row-group's text.
MAX_TEXT_COLUMNS = 12

_ONE_SPAN = {"gte": 1, "lte": 1}


def _json_value(value: Any) -> Any:
    """Reduce a cell value to a JSON-safe primitive for a document field."""
    if value is None or isinstance(value, (str, bool, int)):
        return value
    if isinstance(value, float):
        return value if math.isfinite(value) else None
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    return str(value)


@dataclass(frozen=True, slots=True)
class IngestedWorkbook:
    """The model plus every document it produced, ready for the indexer."""

    model: WorkbookModel
    chunks: tuple[ChunkDocument, ...]
    structure: tuple[StructureDocument, ...]


def _chunk_type_for_region(region: Region) -> ChunkType:
    if region.node_type_kind == "table" or region.kind is RegionKind.TABLE:
        return ChunkType.TABLE
    return ChunkType.REGION


def _bounds(sheet: str, a1: str, sheet_max_row: int) -> tuple[int, int, int, int]:
    """Return ``(min_row, max_row, min_col, max_col)`` for a cell, rectangle or whole column."""
    parts = a1.split(":")
    if len(parts) == 2 and all(part.isalpha() and part for part in parts):
        return 1, max(sheet_max_row, 1), column_index(parts[0]), column_index(parts[1])
    parsed = A1Range.parse(sheet, a1)
    return parsed.min_row, parsed.max_row, parsed.min_col, parsed.max_col


def _region_word(region: Region) -> str:
    return {
        RegionKind.TABLE: "Table",
        RegionKind.NOTES: "Notes block",
        RegionKind.TITLE: "Title",
        RegionKind.GENERIC: "Region",
    }[region.kind]


def _columns_summary(columns: tuple[ColumnSchema, ...]) -> str:
    if not columns:
        return "no columns"
    return ", ".join(f"{column.name} ({column.inferred_type})" for column in columns)


class _Builder:
    """Accumulates chunks and structure nodes, then wires child ids in one final pass."""

    def __init__(self, model: WorkbookModel, ingested_at: datetime) -> None:
        self.model = model
        self.ingested_at = ingested_at
        self.chunks: list[ChunkDocument] = []
        self.structure: dict[str, StructureDocument] = {}
        self.children: dict[str, list[str]] = {}
        self.a1ranges: dict[str, A1Range] = {}

    # -- primitives -------------------------------------------------------------------------
    def add_chunk(
        self,
        *,
        key: str,
        node_id: str,
        sheet_id: str,
        sheet_name: str,
        a1_range: str,
        chunk_type: ChunkType,
        title: str,
        content: str,
        headers: tuple[str, ...] = (),
    ) -> None:
        self.chunks.append(
            ChunkDocument(
                id=chunk_id(self.model.workbook_id, self.model.version, key),
                workbook_id=self.model.workbook_id,
                version=self.model.version,
                node_id=node_id,
                sheet_id=sheet_id,
                sheet_name=sheet_name,
                a1_range=a1_range,
                chunk_type=chunk_type,
                title=title,
                content=content,
                headers=headers,
                acl_scope=self.model.acl_scope,
                source_file=self.model.source_file,
                source_sha256=self.model.source_sha256,
                ingested_at=self.ingested_at,
            )
        )

    def add_node(self, node: StructureDocument) -> None:
        self.structure[node.node_id] = node
        if node.parent_id:
            self.children.setdefault(node.parent_id, []).append(node.node_id)

    def add_sheet_node(
        self,
        *,
        node_id: str,
        node_type: NodeType,
        sheet_id: str,
        sheet_name: str,
        a1: A1Range,
        parent_id: str | None,
        **extra: Any,
    ) -> None:
        self.add_node(
            StructureDocument(
                node_id=node_id,
                workbook_id=self.model.workbook_id,
                version=self.model.version,
                node_type=node_type,
                sheet_id=sheet_id,
                sheet_name=sheet_name,
                a1_range=a1.a1,
                row_span=a1.row_span,
                column_span=a1.column_span,
                parent_id=parent_id,
                acl_scope=self.model.acl_scope,
                source_file=self.model.source_file,
                **extra,
            )
        )

    def finish(self) -> tuple[tuple[ChunkDocument, ...], tuple[StructureDocument, ...]]:
        ordered: list[StructureDocument] = []
        for node in self.structure.values():
            node.child_ids = tuple(dict.fromkeys(self.children.get(node.node_id, ())))
            ordered.append(node)
        return tuple(self.chunks), tuple(ordered)


def _row_group_text(region: Region, group: RowGroup) -> str:
    headers = list(group.headers)
    shown = headers[:MAX_TEXT_COLUMNS]
    lines: list[str] = []
    for row in group.rows:
        pairs = [
            f"{header}={value}" for header, value in zip(shown, row, strict=False) if value != ""
        ]
        if len(headers) > MAX_TEXT_COLUMNS:
            pairs.append(f"(+{len(headers) - MAX_TEXT_COLUMNS} more columns)")
        lines.append("; ".join(pairs) if pairs else "(empty row)")
    where = region.title or f"{region.sheet_name} {region.a1_range.a1}"
    return (
        f"{where} rows {group.a1_range.min_row}-{group.a1_range.max_row} on worksheet "
        f"{region.sheet_name} ({group.a1_range.a1}). Columns: {', '.join(headers)}. "
        + " | ".join(lines)
    )


def _region_text(region: Region, region_columns: tuple[ColumnSchema, ...]) -> str:
    pieces = [
        f"{_region_word(region)} on worksheet {region.sheet_name} covering {region.a1_range.a1}."
    ]
    if region.title:
        pieces.append(f"Titled {region.title!r}.")
    if region.table_name:
        pieces.append(f"Excel table {region.table_name!r}.")
    if region_columns:
        pieces.append(f"Columns: {_columns_summary(region_columns)}.")
    units = [f"{column.name}={column.unit}" for column in region_columns if column.unit]
    if units:
        pieces.append(f"Units: {', '.join(units)}.")
    row_groups = region.row_groups
    if row_groups:
        rows = sum(len(group.rows) for group in row_groups)
        pieces.append(f"{rows} data row(s) across {len(row_groups)} row group(s).")
    if region.notes:
        pieces.append(f"Notes: {' '.join(region.notes)}")
    return " ".join(pieces)


def _column_text(column: ColumnSchema, sheet_name: str) -> str:
    aliases = [alias for alias in column.aliases if alias != column.name]
    parts = [
        f"Column {column.name!r} ({column.letter}) on worksheet {sheet_name} spanning "
        f"{column.a1_range.a1}: inferred type {column.inferred_type}, "
        f"{column.distinct_count} distinct value(s)."
    ]
    if column.unit:
        parts.append(f"Unit: {column.unit}.")
    if aliases:
        parts.append(f"Also seen as: {', '.join(aliases)}.")
    if column.examples:
        parts.append(f"Examples: {', '.join(column.examples)}.")
    return " ".join(parts)


def _formula_text(entry: FormulaEntry) -> str:
    references = ", ".join(
        f"{reference.sheet_name}!{reference.a1_range} ({reference.kind.value})"
        for reference in entry.references
    )
    parts = [
        f"Formulas on worksheet {entry.sheet_name} at {entry.a1_range.a1} compute: {entry.formula}."
    ]
    if entry.is_cluster:
        parts.append(f"This pattern repeats across {len(entry.member_coordinates)} cells.")
    if references:
        parts.append(f"Reads: {references}.")
    if entry.unresolved_references:
        gaps = ", ".join(
            f"{gap.reference_text} ({gap.reason.value})" for gap in entry.unresolved_references
        )
        parts.append(f"Unresolved: {gaps}.")
    return " ".join(parts)


def build_documents(
    model: WorkbookModel, *, ingested_at: datetime | None = None
) -> IngestedWorkbook:
    builder = _Builder(model, ingested_at or datetime.now(UTC))
    workbook_node = workbook_node_id(model.workbook_id, model.version)
    builder.add_node(
        StructureDocument(
            node_id=workbook_node,
            workbook_id=model.workbook_id,
            version=model.version,
            node_type=NodeType.WORKBOOK,
            sheet_id=workbook_node,
            sheet_name="",
            a1_range="",
            row_span=dict(_ONE_SPAN),
            column_span=dict(_ONE_SPAN),
            acl_scope=model.acl_scope,
            source_file=model.source_file,
        )
    )
    sheet_names = [sheet.name for sheet in model.sheets]
    region_total = sum(len(sheet.regions) for sheet in model.sheets)
    builder.add_chunk(
        key="workbook",
        node_id=workbook_node,
        sheet_id=workbook_node,
        sheet_name="",
        a1_range="",
        chunk_type=ChunkType.WORKBOOK,
        title=f"Workbook {model.workbook_id}",
        content=(
            f"Workbook {model.workbook_id!r} (file {model.source_file}, version {model.version}) "
            f"with {len(model.sheets)} worksheet(s): {', '.join(sheet_names)}. "
            f"{region_total} region(s)/table(s) detected. "
            + (
                f"Named ranges: {', '.join(name.name for name in model.named_ranges)}. "
                if model.named_ranges
                else ""
            )
            + (
                f"Macro sheets present (not loaded): {', '.join(model.macro_sheet_names)}. "
                if model.macro_sheet_names
                else ""
            )
            + ("Macros are never executed; formulas are never evaluated." if model.has_vba else "")
        ).strip(),
    )

    for sheet in model.sheets:
        _build_sheet(builder, model, sheet)
    _build_named_ranges(builder, model)
    _materialize_reference_targets(builder, model)

    chunks, structure = builder.finish()
    return IngestedWorkbook(model=model, chunks=chunks, structure=structure)


def _build_sheet(builder: _Builder, model: WorkbookModel, sheet: SheetModel) -> None:
    if sheet.is_macro_sheet:
        _build_macro_sheet(builder, sheet)
        return
    a1 = sheet.a1_range or A1Range.parse(sheet.name, "A1")
    builder.add_sheet_node(
        node_id=sheet.node_id,
        node_type=NodeType.SHEET,
        sheet_id=sheet.node_id,
        sheet_name=sheet.name,
        a1=a1,
        parent_id=workbook_node_id(model.workbook_id, model.version),
    )
    builder.add_chunk(
        key=f"sheet:{sheet.name}",
        node_id=sheet.node_id,
        sheet_id=sheet.node_id,
        sheet_name=sheet.name,
        a1_range=sheet.a1_range.a1 if sheet.a1_range else "",
        chunk_type=ChunkType.SHEET,
        title=f"Worksheet {sheet.name}",
        content=(
            f"Worksheet {sheet.name!r} (visibility {sheet.visibility}) in workbook "
            f"{model.workbook_id}, used range {sheet.a1_range.a1 if sheet.a1_range else 'empty'}. "
            f"{len(sheet.regions)} region(s) detected. "
            + (f"Excel tables: {', '.join(sheet.table_names)}. " if sheet.table_names else "")
            + (
                f"Declared dimension {sheet.declared_dimension!r} was ignored."
                if sheet.declared_dimension_flagged
                else ""
            )
        ).strip(),
    )
    for region in sheet.regions:
        _build_region(builder, sheet, region)
    for entry in sheet.formulas:
        _build_formula(builder, sheet, entry)
    _build_important_cells(builder, model, sheet)


def _build_macro_sheet(builder: _Builder, sheet: SheetModel) -> None:
    node = sheet.node_id
    builder.add_node(
        StructureDocument(
            node_id=node,
            workbook_id=builder.model.workbook_id,
            version=builder.model.version,
            node_type=NodeType.SHEET,
            sheet_id=node,
            sheet_name=sheet.name,
            a1_range="",
            row_span=dict(_ONE_SPAN),
            column_span=dict(_ONE_SPAN),
            acl_scope=builder.model.acl_scope,
            source_file=builder.model.source_file,
        )
    )
    builder.add_chunk(
        key=f"sheet:{sheet.name}",
        node_id=node,
        sheet_id=node,
        sheet_name=sheet.name,
        a1_range="",
        chunk_type=ChunkType.SHEET,
        title=f"Macro sheet {sheet.name}",
        content=(
            f"Worksheet {sheet.name!r} is a macro sheet; its cells are not loaded or executed, and "
            "references to it are recorded as unresolved (macro_sheet)."
        ),
    )


def _region_parent(sheet: SheetModel, a1: A1Range) -> str:
    for region in sheet.regions:
        if region.a1_range.intersects(a1):
            return region.node_id
    return sheet.node_id


def _build_region(builder: _Builder, sheet: SheetModel, region: Region) -> None:
    chunk_type = _chunk_type_for_region(region)
    node_type = NodeType.TABLE if region.kind is RegionKind.TABLE else NodeType.REGION
    builder.add_sheet_node(
        node_id=region.node_id,
        node_type=node_type,
        sheet_id=sheet.node_id,
        sheet_name=sheet.name,
        a1=region.a1_range,
        parent_id=sheet.node_id,
        table_name=region.table_name,
    )
    builder.add_chunk(
        key=f"region:{sheet.name}!{region.a1_range.a1}",
        node_id=region.node_id,
        sheet_id=sheet.node_id,
        sheet_name=sheet.name,
        a1_range=region.a1_range.a1,
        chunk_type=chunk_type,
        title=region.title or f"{sheet.name} {region.a1_range.a1}",
        content=_region_text(region, region.columns),
        headers=tuple(column.name for column in region.columns),
    )
    for column in region.columns:
        _build_column(builder, sheet, region, column)
    for group in region.row_groups:
        _build_row_group(builder, sheet, region, group)


def _build_column(
    builder: _Builder, sheet: SheetModel, region: Region, column: ColumnSchema
) -> None:
    builder.add_sheet_node(
        node_id=column.node_id,
        node_type=NodeType.COLUMN,
        sheet_id=sheet.node_id,
        sheet_name=sheet.name,
        a1=column.a1_range,
        parent_id=region.node_id,
        column_name=column.name,
    )
    builder.add_chunk(
        key=f"col:{sheet.name}!{column.a1_range.a1}",
        node_id=column.node_id,
        sheet_id=sheet.node_id,
        sheet_name=sheet.name,
        a1_range=column.a1_range.a1,
        chunk_type=ChunkType.COLUMN,
        title=column.name,
        content=_column_text(column, sheet.name),
        headers=(column.name,),
    )


def _build_row_group(builder: _Builder, sheet: SheetModel, region: Region, group: RowGroup) -> None:
    builder.add_sheet_node(
        node_id=group.node_id,
        node_type=NodeType.ROW_GROUP,
        sheet_id=sheet.node_id,
        sheet_name=sheet.name,
        a1=group.a1_range,
        parent_id=region.node_id,
    )
    builder.add_chunk(
        key=f"rg:{sheet.name}!{group.a1_range.a1}",
        node_id=group.node_id,
        sheet_id=sheet.node_id,
        sheet_name=sheet.name,
        a1_range=group.a1_range.a1,
        chunk_type=ChunkType.ROW_GROUP,
        title=(
            f"{region.title or sheet.name} rows {group.a1_range.min_row}-{group.a1_range.max_row}"
        ),
        content=_row_group_text(region, group),
        headers=group.headers,
    )


def _build_formula(builder: _Builder, sheet: SheetModel, entry: FormulaEntry) -> None:
    parent = _region_parent(sheet, entry.a1_range)
    node_type = NodeType.FORMULA if entry.node_type_kind == "formula" else NodeType.RANGE
    builder.add_sheet_node(
        node_id=entry.node_id,
        node_type=node_type,
        sheet_id=sheet.node_id,
        sheet_name=sheet.name,
        a1=entry.a1_range,
        parent_id=parent,
        formula=entry.formula,
        cached_value=_json_value(entry.cached_value),
        display_value=format_value(entry.cached_value) if entry.cached_value is not None else None,
        references=entry.references,
        unresolved_references=entry.unresolved_references,
    )
    builder.add_chunk(
        key=f"formula:{sheet.name}!{entry.a1_range.a1}",
        node_id=entry.node_id,
        sheet_id=sheet.node_id,
        sheet_name=sheet.name,
        a1_range=entry.a1_range.a1,
        chunk_type=ChunkType.FORMULA_SUMMARY,
        title=(
            f"Formula in {sheet.name} {entry.a1_range.a1}"
            if not entry.is_cluster
            else f"Formula pattern {sheet.name} {entry.a1_range.a1}"
        ),
        content=_formula_text(entry),
    )


def _header_coordinates(sheet: SheetModel) -> dict[str, str]:
    coordinates: dict[str, str] = {}
    for region in sheet.regions:
        for column in region.columns:
            coordinates[column.header_coordinate] = column.name
    return coordinates


def _build_important_cells(builder: _Builder, model: WorkbookModel, sheet: SheetModel) -> None:
    for coordinate, name in _header_coordinates(sheet).items():
        node = cell_node_id(model.workbook_id, model.version, sheet.name, coordinate)
        if node in builder.structure:
            continue
        try:
            bounds = A1Range.parse(sheet.name, coordinate)
        except ValueError:
            continue
        cell = sheet.cells.get(coordinate)
        builder.add_node(
            StructureDocument(
                node_id=node,
                workbook_id=model.workbook_id,
                version=model.version,
                node_type=NodeType.CELL,
                sheet_id=sheet.node_id,
                sheet_name=sheet.name,
                a1_range=bounds.a1,
                row_span=bounds.row_span,
                column_span=bounds.column_span,
                parent_id=_region_parent(sheet, bounds),
                cached_value=_json_value(cell.cached_value) if cell else None,
                display_value=(
                    format_value(cell.value if cell.formula is None else cell.cached_value)
                    if cell
                    else None
                ),
                column_name=name,
                acl_scope=model.acl_scope,
                source_file=model.source_file,
            )
        )


def _materialize_reference_targets(builder: _Builder, model: WorkbookModel) -> None:
    """Create a structure node for every referenced range/cell/table-column a formula points at.

    A ``target_node_id`` must never dangle: retrieval resolves it with a point lookup. Targets that
    region detection already produced a node for are left alone; the rest become bare range/cell or
    table-column nodes carrying their span, so the edge always lands on something real.
    """
    sheets = {sheet.name: sheet for sheet in model.sheets}
    sheet_max_row = {
        sheet.name: max((cell.row for cell in sheet.cells.values()), default=1)
        for sheet in model.sheets
    }
    wanted: dict[str, tuple[str, str, str, NodeType]] = {}
    for sheet in model.sheets:
        if sheet.is_macro_sheet:
            continue
        for entry in sheet.formulas:
            for reference in entry.references:
                wanted.setdefault(
                    reference.target_node_id,
                    (
                        reference.sheet_name,
                        reference.a1_range,
                        reference.kind.value,
                        _node_type_for(reference),
                    ),
                )
    for name in model.named_ranges:
        if name.resolved and name.sheet_name and name.a1 and name.target_node_id:
            single = ":" not in name.a1
            wanted.setdefault(
                name.target_node_id,
                (
                    name.sheet_name,
                    name.a1,
                    "cell" if single else "range",
                    NodeType.CELL if single else NodeType.RANGE,
                ),
            )

    for node_id, (sheet_name, a1, kind, node_type) in sorted(wanted.items()):
        if node_id in builder.structure:
            continue
        target_sheet = sheets.get(sheet_name)
        if target_sheet is None or target_sheet.is_macro_sheet:
            continue
        try:
            min_row, max_row, min_col, max_col = _bounds(sheet_name, a1, sheet_max_row[sheet_name])
        except ValueError:
            continue
        parent = target_sheet.node_id
        if not (min_row == 1 and max_row == sheet_max_row[sheet_name]):
            try:
                parent = _region_parent(target_sheet, A1Range.parse(sheet_name, a1))
            except ValueError:
                parent = target_sheet.node_id
        column_name = None
        if kind == "table_column":
            column_name = node_id.rsplit(":", 1)[-1].split("!")[-1] or None
        builder.add_node(
            StructureDocument(
                node_id=node_id,
                workbook_id=model.workbook_id,
                version=model.version,
                node_type=node_type,
                sheet_id=target_sheet.node_id,
                sheet_name=sheet_name,
                a1_range=a1,
                row_span={"gte": min_row, "lte": max_row},
                column_span={"gte": min_col, "lte": max_col},
                parent_id=parent,
                column_name=column_name,
                acl_scope=model.acl_scope,
                source_file=model.source_file,
            )
        )


def _node_type_for(reference: Reference) -> NodeType:
    if reference.kind is ReferenceKind.TABLE_COLUMN:
        return NodeType.COLUMN
    if reference.kind is ReferenceKind.NAMED_RANGE:
        return NodeType.NAMED_RANGE
    if reference.kind is ReferenceKind.CELL and ":" not in reference.a1_range:
        return NodeType.CELL
    return NodeType.RANGE


def _build_named_ranges(builder: _Builder, model: WorkbookModel) -> None:
    workbook_node = workbook_node_id(model.workbook_id, model.version)
    for name in model.named_ranges:
        references: list[Reference] = []
        parent = workbook_node
        a1_range = ""
        row_span = dict(_ONE_SPAN)
        column_span = dict(_ONE_SPAN)
        sheet_id = workbook_node
        if name.resolved and name.sheet_name and name.a1:
            bounds = A1Range.parse(name.sheet_name, name.a1)
            a1_range = bounds.a1
            row_span = dict(bounds.row_span)
            column_span = dict(bounds.column_span)
            target_sheet = next(
                (sheet for sheet in model.sheets if sheet.name == name.sheet_name), None
            )
            if target_sheet is not None:
                sheet_id = target_sheet.node_id
                parent = target_sheet.node_id
            single = bounds.min_row == bounds.max_row and bounds.min_col == bounds.max_col
            references.append(
                Reference(
                    target_node_id=name.target_node_id or "",
                    sheet_name=name.sheet_name,
                    a1_range=bounds.a1,
                    kind=ReferenceKind.CELL if single else ReferenceKind.RANGE,
                )
            )
        builder.add_node(
            StructureDocument(
                node_id=name.node_id,
                workbook_id=model.workbook_id,
                version=model.version,
                node_type=NodeType.NAMED_RANGE,
                sheet_id=sheet_id,
                sheet_name=name.sheet_name or "",
                a1_range=a1_range,
                row_span=row_span,
                column_span=column_span,
                parent_id=parent,
                named_range=name.name,
                references=tuple(references),
                acl_scope=model.acl_scope,
                source_file=model.source_file,
            )
        )
