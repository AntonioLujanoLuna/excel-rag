"""The canonical workbook representation.

This exists **only during ingestion** -- the durable copy is the Elasticsearch indices. It is the
hierarchy the design calls for: workbook -> sheet -> region/table -> column/row-group, plus the
typed formula edges, kept in plain dataclasses so it never accidentally becomes a second store.

Every id is created through :func:`excel_rag.models.node_id`, so a reindex of the same workbook at
the same version reproduces byte-identical ids and a stale version can be dropped wholesale. The
helpers below fix the ``kind`` and ``key`` conventions for each node type so ingestion and the
structure documents agree on one scheme.
"""

from __future__ import annotations

import re
from collections.abc import Mapping
from dataclasses import dataclass, field
from enum import StrEnum
from typing import TYPE_CHECKING, Any

from ..models import A1Range, Reference, UnresolvedReference, node_id

if TYPE_CHECKING:
    from .formulas import FormulaContext


class RegionKind(StrEnum):
    """What a detected region is, in the terms the design uses."""

    TITLE = "title"
    TABLE = "table"
    NOTES = "notes"
    GENERIC = "generic"


# -------------------------------------------------------------------------------------------------
# Node id helpers -- one place that decides the kind/key convention.
# -------------------------------------------------------------------------------------------------
def workbook_node_id(workbook_id: str, version: int) -> str:
    return node_id(workbook_id, version, "workbook", workbook_id)


def sheet_node_id(workbook_id: str, version: int, sheet: str) -> str:
    return node_id(workbook_id, version, "sheet", sheet)


def region_node_id(workbook_id: str, version: int, sheet: str, a1: str) -> str:
    return node_id(workbook_id, version, "region", f"{sheet}!{a1}")


def table_node_id(workbook_id: str, version: int, table_name: str) -> str:
    return node_id(workbook_id, version, "table", table_name)


def column_node_id(workbook_id: str, version: int, sheet: str, column_range: str) -> str:
    return node_id(workbook_id, version, "column", f"{sheet}!{column_range}")


def row_group_node_id(workbook_id: str, version: int, sheet: str, a1: str) -> str:
    return node_id(workbook_id, version, "row_group", f"{sheet}!{a1}")


def cell_node_id(workbook_id: str, version: int, sheet: str, coordinate: str) -> str:
    return node_id(workbook_id, version, "cell", f"{sheet}!{coordinate}")


def range_node_id(workbook_id: str, version: int, sheet: str, a1: str) -> str:
    return node_id(workbook_id, version, "range", f"{sheet}!{a1}")


def formula_node_id(workbook_id: str, version: int, sheet: str, coordinate: str) -> str:
    return node_id(workbook_id, version, "formula", f"{sheet}!{coordinate}")


def sheet_object_node_id(workbook_id: str, version: int, kind: str, sheet: str, key: str) -> str:
    """A chart (keyed by its ordinal), pivot table (its name) or validation (its cells)."""
    return node_id(workbook_id, version, kind, f"{sheet}!{key}")


def table_column_node_id(
    workbook_id: str, version: int, sheet: str, table: str, column: str
) -> str:
    return node_id(workbook_id, version, "column", f"{sheet}!{table}!{column}")


def named_range_node_id(
    workbook_id: str, version: int, name: str, scope_sheet: str | None = None
) -> str:
    """A workbook-wide name keys on its name; a sheet-scoped one on ``Sheet!Name``, so the same
    local name on two sheets (``Rate`` on every monthly sheet) is two nodes."""
    key = name if scope_sheet is None else f"{scope_sheet}!{name}"
    return node_id(workbook_id, version, "named_range", key)


def chunk_id(workbook_id: str, version: int, key: str) -> str:
    return node_id(workbook_id, version, "chunk", key)


_COLUMN_LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


def column_letter(index: int) -> str:
    """``1`` -> ``A``, ``27`` -> ``AA``. Dependency-light: no openpyxl import here."""
    if index < 1:
        raise ValueError(f"column index must be >= 1, got {index}")
    letters = ""
    value = index
    while value > 0:
        value, remainder = divmod(value - 1, 26)
        letters = _COLUMN_LETTERS[remainder] + letters
    return letters


def format_value(value: Any) -> str:
    """A stable, human-readable rendering of a cell value for chunk text and examples.

    Dates become ISO strings, floats with no fraction lose the trailing ``.0``, booleans read as
    Excel does. Never used for computation -- only for prose the search index carries.
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


# -------------------------------------------------------------------------------------------------
# Cells and column schemas
# -------------------------------------------------------------------------------------------------
@dataclass(frozen=True, slots=True)
class CellValue:
    """One populated cell, with both its formula and its last-saved cached value kept separately."""

    sheet_name: str
    coordinate: str
    row: int
    column: int
    value: Any = None
    formula: str | None = None
    cached_value: Any = None
    data_type: str = "n"
    number_format: str = "General"
    is_date: bool = False
    is_percentage: bool = False
    is_currency: bool = False
    merged_range: str | None = None
    #: For a cell inside an array formula's saved extent (not its master): the master coordinate.
    #: Its ``cached_value`` is the formula's last-saved result there; ``value`` stays ``None``.
    array_master: str | None = None
    #: For an array formula's master cell: the extent its result covers (``B2:B40``).
    array_range: str | None = None
    #: A formula cell whose text the format does not expose (``.xlsb``): only its last-saved value
    #: (``cached_value``) is known, and it has no precedents.
    formula_unavailable: bool = False

    @property
    def computed(self) -> bool:
        """Whether the cell's content is a formula result rather than input data."""
        return self.formula is not None or self.array_master is not None or self.formula_unavailable


@dataclass(frozen=True, slots=True)
class ColumnSchema:
    """A column of a region: its header, inferred type, aliases, and a few example values."""

    index: int
    letter: str
    name: str
    inferred_type: str
    aliases: tuple[str, ...]
    distinct_count: int
    examples: tuple[str, ...]
    unit: str | None
    a1_range: A1Range
    node_id: str
    header_coordinate: str


@dataclass(frozen=True, slots=True)
class RowGroup:
    """A bounded set of rows of a region, carrying its column headers so a row stands alone."""

    index: int
    a1_range: A1Range
    node_id: str
    headers: tuple[str, ...]
    rows: tuple[tuple[str, ...], ...]


@dataclass(frozen=True, slots=True)
class Region:
    """A logical region of a sheet: a title, a table, a notes block or a headerless grid."""

    kind: RegionKind
    sheet_name: str
    a1_range: A1Range
    node_id: str
    node_type_kind: str
    title: str | None
    table_name: str | None
    header_rows: tuple[int, ...]
    units_row: int | None
    columns: tuple[ColumnSchema, ...]
    row_groups: tuple[RowGroup, ...]
    notes: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class NamedRange:
    """A defined name and where it points, if that can be resolved statically.

    ``scope_sheet`` is the sheet a sheet-scoped name is local to, ``None`` for a workbook-wide one.
    """

    name: str
    scope_sheet: str | None
    sheet_name: str | None
    a1: str | None
    node_id: str
    resolved: bool
    target_node_id: str | None
    detail: str | None = None

    @property
    def scope(self) -> str:
        """``workbook`` or the name of the sheet it is local to."""
        return "workbook" if self.scope_sheet is None else self.scope_sheet

    @property
    def label(self) -> str:
        """How a formula on another sheet would write it: ``Rate`` or ``Jan!Rate``."""
        if self.scope_sheet is None:
            return self.name
        sheet = self.scope_sheet
        quoted = sheet if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", sheet) else f"'{sheet}'"
        return f"{quoted}!{self.name}"


@dataclass(frozen=True, slots=True)
class FormulaEntry:
    """A formula node, either a single cell formula or a cluster of one repeated pattern.

    A cluster is the design's answer to a 500-row formula column: one node and one summary for the
    whole run instead of 500 cell documents.
    """

    node_id: str
    node_type_kind: str
    sheet_name: str
    a1_range: A1Range
    formula: str
    cached_value: Any
    references: tuple[Reference, ...]
    unresolved_references: tuple[UnresolvedReference, ...]
    member_coordinates: tuple[str, ...]
    is_cluster: bool
    #: An array formula's saved extent (``B2:B40``), the cells its result last spilled into.
    array_range: str | None = None


@dataclass(frozen=True, slots=True)
class SheetObject:
    """A chart, pivot table or data validation, and the ranges it reads.

    ``anchor`` is where it sits (a chart's top-left cell, a pivot's location) or what it governs (a
    validation's first range of cells). Its edges are resolved like a formula's, and kept apart
    from formulas because it is not one: nothing here has a cached value or a formula text.
    """

    kind: str  # "chart" | "pivot_table" | "data_validation"
    node_id: str
    sheet_name: str
    name: str
    detail: str
    anchor: A1Range
    sources: tuple[str, ...]
    references: tuple[Reference, ...]
    unresolved_references: tuple[UnresolvedReference, ...]

    @property
    def label(self) -> str:
        word = self.detail[:1].upper() + self.detail[1:]
        return word if self.kind == "data_validation" else f"{word} {self.name!r}"


@dataclass(frozen=True, slots=True)
class SheetModel:
    """A worksheet after region detection."""

    name: str
    visibility: str
    node_id: str
    a1_range: A1Range | None
    regions: tuple[Region, ...]
    formulas: tuple[FormulaEntry, ...]
    cells: Mapping[str, CellValue]
    declared_dimension: str | None
    declared_dimension_flagged: bool
    is_macro_sheet: bool
    table_names: tuple[str, ...] = ()
    objects: tuple[SheetObject, ...] = ()


@dataclass(frozen=True, slots=True)
class WorkbookModel:
    """The whole workbook, in memory, ready to be turned into documents."""

    workbook_id: str
    version: int
    source_file: str
    source_sha256: str
    sheets: tuple[SheetModel, ...]
    named_ranges: tuple[NamedRange, ...]
    macro_sheet_names: tuple[str, ...] = ()
    has_vba: bool = False
    acl_scope: tuple[str, ...] = ()
    warnings: tuple[str, ...] = field(default_factory=tuple)
    #: What each sheet's formulas were resolved against (tables, names, spill extents), so a formula
    #: can be resolved again later the same way -- by the evaluator, which rewrites references.
    formula_contexts: Mapping[str, FormulaContext] = field(
        default_factory=dict, compare=False, repr=False
    )
