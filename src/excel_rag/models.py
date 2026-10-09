"""The index documents and the retrieval contract. Frozen: both workstreams code against these.

Two indices, sharing ``workbook_id``, ``version``, ``sheet_id`` and ``node_id``:

* :data:`~excel_rag.es.INDEX_CHUNKS` carries what is *searchable* -- region descriptions,
  contextualized row groups, column schemas, formula summaries.
* :data:`~excel_rag.es.INDEX_STRUCTURE` carries what is *exact* -- cells, ranges, tables, named
  ranges, formulas, and the typed edges between them.

The split is the design's point: content for embedding is separated from the underlying exact
values, so a hit can be scored semantically and still answer with the A1 range, the cached value
and the formula text. Nothing here generates an answer; the response is retrieval evidence.

Every identifier is deterministic within ``(workbook_id, version)`` -- see :func:`node_id` -- so a
reindex produces the same ids and a stale version can be dropped wholesale.
"""

from __future__ import annotations

import re
from datetime import datetime
from enum import StrEnum
from typing import Any

from pydantic import BaseModel, ConfigDict, Field, computed_field, field_validator, model_validator

# -------------------------------------------------------------------------------------------------
# Identity
# -------------------------------------------------------------------------------------------------
_ID_SAFE = re.compile(r"[^a-z0-9_.!$:-]+")


def node_id(workbook_id: str, version: int, kind: str, key: str) -> str:
    """A deterministic structural id: ``wb42:v3:cell:assumptions!c7``.

    ``kind`` is the node type, ``key`` the sheet-qualified coordinate or symbol name. Everything is
    lowercased and squashed so the id is stable across runs and safe as an Elasticsearch ``_id``.
    """
    slug = _ID_SAFE.sub("_", key.strip().lower())
    return f"{workbook_id}:v{version}:{kind}:{slug}"


def version_key(workbook_id: str, version: int) -> str:
    """The ``(workbook_id, version)`` pair as one keyword: ``wb42:v3``.

    Every document carries it, so pinning a request to the active versions is a single ``terms``
    clause however many workbooks there are, instead of one boolean clause per workbook.
    """
    return f"{workbook_id}:v{version}"


class ChunkType(StrEnum):
    WORKBOOK = "workbook"
    SHEET = "sheet"
    TABLE = "table"
    REGION = "region"
    COLUMN = "column"
    ROW_GROUP = "row_group"
    FORMULA_SUMMARY = "formula_summary"


class NodeType(StrEnum):
    WORKBOOK = "workbook"
    SHEET = "sheet"
    TABLE = "table"
    REGION = "region"
    COLUMN = "column"
    ROW_GROUP = "row_group"
    CELL = "cell"
    RANGE = "range"
    FORMULA = "formula"
    NAMED_RANGE = "named_range"


class ReferenceKind(StrEnum):
    CELL = "cell"
    RANGE = "range"
    TABLE_COLUMN = "table_column"
    NAMED_RANGE = "named_range"


class UnresolvedReason(StrEnum):
    """Why a reference has no edge. Fail explicitly: never invent a graph edge."""

    INDIRECT = "indirect"
    VOLATILE_OFFSET = "volatile_offset"
    EXTERNAL_LINK = "external_link"
    DYNAMIC_ARRAY = "dynamic_array"
    UNSUPPORTED_FUNCTION = "unsupported_function"
    MACRO_SHEET = "macro_sheet"
    MALFORMED = "malformed"
    OUT_OF_RANGE = "out_of_range"


# -------------------------------------------------------------------------------------------------
# Coordinates
# -------------------------------------------------------------------------------------------------
class A1Range(BaseModel):
    """An inclusive A1 rectangle on one sheet, with its parsed bounds.

    Both representations are kept: the A1 string is what a caller can open in Excel, the integer
    bounds are what an Elasticsearch ``integer_range`` query intersects.
    """

    model_config = ConfigDict(frozen=True)

    sheet_name: str
    a1: str = Field(description="Inclusive A1 rectangle, e.g. 'A12:F25' or 'C7'.")
    min_row: int = Field(ge=1)
    max_row: int = Field(ge=1)
    min_col: int = Field(ge=1)
    max_col: int = Field(ge=1)

    @model_validator(mode="after")
    def _ordered(self) -> A1Range:
        if self.min_row > self.max_row or self.min_col > self.max_col:
            raise ValueError(f"inverted A1 range: {self.a1!r}")
        return self

    @classmethod
    def parse(cls, sheet_name: str, a1: str) -> A1Range:
        """Parse ``A1`` or ``A1:B7`` form. Raises ``ValueError`` on anything else."""
        match = re.fullmatch(
            r"\s*\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})(?:\s*:\s*\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6}))?\s*",
            a1,
        )
        if match is None:
            raise ValueError(f"not an A1 range: {a1!r}")
        start_col, start_row, end_col, end_row = match.groups()
        min_col = _column_index(start_col)
        min_row = int(start_row)
        max_col = _column_index(end_col) if end_col else min_col
        max_row = int(end_row) if end_row else min_row
        canonical = (
            f"{start_col.upper()}{min_row}"
            if (max_col, max_row) == (min_col, min_row)
            else f"{start_col.upper()}{min_row}:{end_col.upper()}{max_row}"
            if end_col
            else f"{start_col.upper()}{min_row}"
        )
        return cls(
            sheet_name=sheet_name,
            a1=canonical,
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
        )

    @property
    def row_span(self) -> dict[str, int]:
        return {"gte": self.min_row, "lte": self.max_row}

    @property
    def column_span(self) -> dict[str, int]:
        return {"gte": self.min_col, "lte": self.max_col}

    @property
    def cell_count(self) -> int:
        return (self.max_row - self.min_row + 1) * (self.max_col - self.min_col + 1)

    def intersects(self, other: A1Range) -> bool:
        """Whether two rectangles on the *same* sheet overlap. Callers check the sheet first."""
        if self.sheet_name != other.sheet_name:
            return False
        return not (
            self.max_row < other.min_row
            or other.max_row < self.min_row
            or self.max_col < other.min_col
            or other.max_col < self.min_col
        )

    def contains_row(self, row: int) -> bool:
        return self.min_row <= row <= self.max_row

    def contains_column(self, column: int) -> bool:
        return self.min_col <= column <= self.max_col


def _column_index(letters: str) -> int:
    """``A`` -> 1, ``Z`` -> 26, ``AA`` -> 27."""
    index = 0
    for char in letters.upper():
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index


# -------------------------------------------------------------------------------------------------
# Structural edges
# -------------------------------------------------------------------------------------------------
class Reference(BaseModel):
    """A typed edge from a formula (or named range) to the node it reads.

    A reference covering a large area points at a *range* node, never at one edge per cell: the
    rectangle is indexed once and overlaps are resolved by a range query at retrieval time.
    """

    model_config = ConfigDict(frozen=True)

    target_node_id: str
    sheet_name: str
    a1_range: str
    kind: ReferenceKind
    resolved: bool = True


class UnresolvedReference(BaseModel):
    """A reference that cannot be resolved statically, kept with the reason it failed."""

    model_config = ConfigDict(frozen=True)

    reference_text: str
    reason: UnresolvedReason
    detail: str | None = None


# -------------------------------------------------------------------------------------------------
# Documents
# -------------------------------------------------------------------------------------------------
class SourceRef(BaseModel):
    """Where a document or a hit lives: enough to open the right workbook, sheet and range."""

    model_config = ConfigDict(frozen=True)

    workbook_id: str
    version: int
    sheet: str
    a1_range: str


class ChunkDocument(BaseModel):
    """One searchable document in ``excel_chunks``."""

    id: str
    workbook_id: str
    version: int
    node_id: str
    sheet_id: str
    sheet_name: str
    a1_range: str
    chunk_type: ChunkType
    title: str
    content: str
    headers: tuple[str, ...] = ()
    embedding: tuple[float, ...] | None = None
    embedding_model: str | None = Field(
        default=None, description="Model that produced `embedding`; a different model is a miss."
    )
    colbert: tuple[tuple[float, ...], ...] | None = None
    acl_scope: tuple[str, ...] = ()
    source_file: str | None = None
    source_sha256: str | None = None
    ingested_at: datetime | None = None
    ingest_run: str | None = Field(
        default=None, description="The indexing run that wrote this document; see the indexer."
    )

    @computed_field  # type: ignore[prop-decorator]
    @property
    def version_key(self) -> str:
        return version_key(self.workbook_id, self.version)


class StructureDocument(BaseModel):
    """One structural node in ``excel_structure``: a cell, range, table, formula, named range."""

    node_id: str
    workbook_id: str
    version: int
    node_type: NodeType
    sheet_id: str
    sheet_name: str
    a1_range: str
    row_span: dict[str, int]
    column_span: dict[str, int]
    parent_id: str | None = None
    child_ids: tuple[str, ...] = ()
    formula: str | None = None
    cached_value: Any = Field(
        default=None,
        description="Last value Excel saved. A cached value is not a freshly computed result.",
    )
    display_value: str | None = None
    references: tuple[Reference, ...] = ()
    unresolved_references: tuple[UnresolvedReference, ...] = ()
    table_name: str | None = None
    column_name: str | None = None
    named_range: str | None = None
    acl_scope: tuple[str, ...] = ()
    source_file: str | None = None
    ingest_run: str | None = Field(
        default=None, description="The indexing run that wrote this document; see the indexer."
    )

    @computed_field  # type: ignore[prop-decorator]
    @property
    def version_key(self) -> str:
        return version_key(self.workbook_id, self.version)


class ActiveVersionManifest(BaseModel):
    """The active version per workbook. Replacement versions are indexed before activation."""

    workbook_id: str
    active_version: int = Field(ge=1)
    replaced_at: datetime | None = None


# -------------------------------------------------------------------------------------------------
# Retrieval contract
# -------------------------------------------------------------------------------------------------
class SearchFilters(BaseModel):
    """Filters applied inside the query. Unknown keys are refused, not ignored.

    The tolerance this replaces was not a convenience: a caller that put `expand_references` or
    `reference_depth` here instead of at the top level got a 200 with no expansion and no warning,
    which reads exactly like a broken traversal. A request field in the wrong place is refused here
    (a 400 with the standard error envelope).
    """

    model_config = ConfigDict(frozen=True, extra="forbid")

    workbook_ids: tuple[str, ...] = ()
    sheet_names: tuple[str, ...] = ()
    chunk_types: tuple[ChunkType, ...] = ()
    acl_scopes: tuple[str, ...] = Field(
        default=(),
        description="Scopes the caller holds. An empty tuple means no scope filter is applied, "
        "which is only correct where the deployment has no ACLs; the route layer decides.",
    )


class SearchRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    query: str = Field(description="Natural-language query. Blank input is refused below.")
    filters: SearchFilters = Field(default_factory=SearchFilters)
    top_k: int = Field(default=10, ge=1, le=100)
    include_structure: bool = True
    expand_references: bool = False
    reference_depth: int = Field(default=1, ge=0, le=3)
    max_related_nodes: int = Field(default=20, ge=0, le=200)

    @field_validator("query")
    @classmethod
    def _not_blank(cls, value: str) -> str:
        if not value.strip():
            raise ValueError("query must not be blank")
        return value


class NodePayload(BaseModel):
    """An exact structural node returned next to the semantic hit that referenced it."""

    model_config = ConfigDict(frozen=True)

    node_id: str
    node_type: NodeType
    sheet: str
    a1_range: str
    value: Any = None
    display_value: str | None = None
    formula: str | None = None
    cached_value: Any = None
    table_name: str | None = None
    column_name: str | None = None
    named_range: str | None = None
    references: tuple[Reference, ...] = ()
    depth: int = 0


class Hit(BaseModel):
    model_config = ConfigDict(frozen=True)

    chunk_id: str
    score: float
    content: str
    source: SourceRef
    node_id: str
    title: str | None = None
    headers: tuple[str, ...] = ()
    matched_fields: tuple[str, ...] = ()
    related_node_ids: tuple[str, ...] = ()


class TruncationInfo(BaseModel):
    """What a bounded expansion dropped, and why. Silent omission is a bug in this service."""

    model_config = ConfigDict(frozen=True)

    truncated: bool = False
    reason: str | None = None
    dropped_nodes: int = 0
    depth_limit: int | None = None
    bytes_returned: int = 0


class SearchResponse(BaseModel):
    hits: tuple[Hit, ...]
    nodes: dict[str, NodePayload] = Field(default_factory=dict)
    unresolved_references: tuple[UnresolvedReference, ...] = ()
    truncation: TruncationInfo = Field(default_factory=TruncationInfo)
    took_ms: float = 0.0
    es_requests: int = Field(default=0, description="Elasticsearch calls this request cost.")


class StructureQuery(BaseModel):
    """The optional direct inspection path: structure for a workbook, with no semantic search."""

    model_config = ConfigDict(extra="forbid")

    workbook_ids: tuple[str, ...] = ()
    sheet_names: tuple[str, ...] = ()
    node_types: tuple[NodeType, ...] = ()
    limit: int = Field(default=200, ge=1, le=2000)


class RangeQuery(BaseModel):
    """Which nodes overlap a rectangle, resolved by span intersection, not by scanning cells."""

    model_config = ConfigDict(extra="forbid")

    workbook_id: str
    sheet_name: str
    a1: str
    node_types: tuple[NodeType, ...] = ()
    limit: int = Field(default=50, ge=1, le=500)
