"""Build a synthetic corpus directly from the models -- no ingestion path involved.

The retrieval layer is what is measured, so the corpus is assembled with
:class:`~excel_rag.models.ChunkDocument`, :class:`~excel_rag.models.StructureDocument` and
:class:`~excel_rag.models.ActiveVersionManifest` and handed to the client as index-ready pairs.

Two properties matter for the benchmark:

* **chunk node ids resolve.** Every chunk's ``node_id`` is the id of a structure node that exists,
  so ``include_structure`` returns something real rather than a dangling reference.
* **formula hits carry references.** Formula-summary chunks sit on formula nodes that reference a
  chain (``f5 -> f4 -> ... -> f0``), a cross-sheet cell, and an ``A <-> B`` cycle, so reference
  expansion has depth to walk and the cycle has something to terminate on.
"""

from __future__ import annotations

import random
from collections.abc import Callable
from dataclasses import dataclass
from datetime import UTC, datetime
from typing import Any

from ..models import (
    ActiveVersionManifest,
    ChunkDocument,
    ChunkType,
    NodeType,
    Reference,
    ReferenceKind,
    StructureDocument,
    node_id,
)

DEFAULT_WORKBOOKS = 3
DEFAULT_SHEETS = 4
DEFAULT_CELLS_PER_SHEET = 2500
DEFAULT_CHUNKS_PER_SHEET = 180
DEFAULT_EMBED_DIMS = 64
DEFAULT_ACL_SCOPES = ("finance-team", "hr-team", "board")
_INGESTED_AT = datetime(2026, 10, 9, tzinfo=UTC)
_EMBEDDING_MODEL = "bench-embed-v1"

#: Chunks per sheet outside the row groups: 1 sheet + 1 table + 20 columns + 7 formula summaries.
_FIXED_CHUNKS = 29


@dataclass(frozen=True)
class Corpus:
    """Index-ready documents plus the identifiers a benchmark query would use."""

    chunks: tuple[tuple[str, dict[str, Any]], ...]
    structure: tuple[tuple[str, dict[str, Any]], ...]
    versions: tuple[tuple[str, dict[str, Any]], ...]
    workbook_ids: tuple[str, ...]
    query: str
    formula_node_id: str
    chunk_id: str

    @property
    def chunk_count(self) -> int:
        return len(self.chunks)

    @property
    def structure_count(self) -> int:
        return len(self.structure)


def build_corpus(
    *,
    workbooks: int = DEFAULT_WORKBOOKS,
    sheets: int = DEFAULT_SHEETS,
    cells_per_sheet: int = DEFAULT_CELLS_PER_SHEET,
    chunks_per_sheet: int = DEFAULT_CHUNKS_PER_SHEET,
    embed_dims: int = DEFAULT_EMBED_DIMS,
    acl_scopes: tuple[str, ...] = DEFAULT_ACL_SCOPES,
    seed: int = 20261009,
) -> Corpus:
    """Build the corpus deterministically from ``seed``."""
    rng = random.Random(seed)

    def embedding() -> list[float]:
        return [round(rng.uniform(-1.0, 1.0), 4) for _ in range(embed_dims)]

    chunk_docs: list[tuple[str, dict[str, Any]]] = []
    structure_docs: list[tuple[str, dict[str, Any]]] = []
    version_docs: list[tuple[str, dict[str, Any]]] = []
    workbook_ids: list[str] = []
    formula_node_id = ""
    first_chunk_id = ""

    for wb_index in range(workbooks):
        workbook_id = f"wb{wb_index}"
        workbook_ids.append(workbook_id)
        version = 1
        scope = acl_scopes[wb_index % len(acl_scopes)]
        version_docs.append(
            (
                workbook_id,
                ActiveVersionManifest(workbook_id=workbook_id, active_version=version).model_dump(
                    mode="json"
                ),
            )
        )
        sheet_names = [f"Sheet{wb_index}_{s}" for s in range(sheets)]
        for sheet_index, sheet_name in enumerate(sheet_names):
            sheet = _SheetContext(
                workbook_id=workbook_id,
                version=version,
                sheet_name=sheet_name,
                sheet_id=node_id(workbook_id, version, "sheet", sheet_name),
                next_sheet_name=sheet_names[(sheet_index + 1) % len(sheet_names)],
                acl=[scope],
            )
            formulas, chunk_id = _index_sheet(
                sheet,
                cells_per_sheet=cells_per_sheet,
                chunks_per_sheet=chunks_per_sheet,
                embedding=embedding,
                chunk_docs=chunk_docs,
                structure_docs=structure_docs,
            )
            if not formula_node_id:
                formula_node_id = formulas[1] if len(formulas) > 1 else formulas[0]
            if not first_chunk_id:
                first_chunk_id = chunk_id

    return Corpus(
        chunks=tuple(chunk_docs),
        structure=tuple(structure_docs),
        versions=tuple(version_docs),
        workbook_ids=tuple(workbook_ids),
        query="projected revenue",
        formula_node_id=formula_node_id,
        chunk_id=first_chunk_id,
    )


@dataclass(frozen=True)
class _SheetContext:
    workbook_id: str
    version: int
    sheet_name: str
    sheet_id: str
    next_sheet_name: str
    acl: list[str]


def _index_sheet(
    sheet: _SheetContext,
    *,
    cells_per_sheet: int,
    chunks_per_sheet: int,
    embedding: Callable[[], list[float]],
    chunk_docs: list[tuple[str, dict[str, Any]]],
    structure_docs: list[tuple[str, dict[str, Any]]],
) -> tuple[list[str], str]:
    """Index one worksheet, returning its formula node ids and its first row-group chunk id."""
    sheet_node = _node(structure_docs, sheet, NodeType.SHEET, "A1", 1, 1, 1, 1)
    table_id = _node(
        structure_docs,
        sheet,
        NodeType.TABLE,
        "A1:D1000",
        1,
        1000,
        1,
        4,
        table_name=f"tbl_{sheet.sheet_name.lower()}",
    )
    _node(structure_docs, sheet, NodeType.REGION, "A1:D1000", 1, 1000, 1, 4)
    _node(
        structure_docs,
        sheet,
        NodeType.NAMED_RANGE,
        "C7",
        7,
        7,
        3,
        3,
        named_range=f"rate_{sheet.sheet_name.lower()}",
    )

    _chunk(
        chunk_docs,
        sheet,
        sheet_node,
        "A1",
        ChunkType.SHEET,
        f"{sheet.sheet_name} overview",
        f"Worksheet {sheet.sheet_name} of {sheet.workbook_id}: projected revenue and growth.",
        embedding(),
    )
    _chunk(
        chunk_docs,
        sheet,
        table_id,
        "A1:D1000",
        ChunkType.TABLE,
        f"Revenue table {sheet.sheet_name}",
        "Projected revenue by year, linked to historical actuals and growth assumptions.",
        embedding(),
    )

    for column in range(1, 21):
        column_id = _node(
            structure_docs, sheet, NodeType.COLUMN, f"C{column}", 1, 1001, column, column
        )
        _chunk(
            chunk_docs,
            sheet,
            column_id,
            f"C{column}",
            ChunkType.COLUMN,
            f"Column C{column}",
            f"Column {column} of the revenue table: projected revenue values.",
            embedding(),
        )

    first_chunk_id = ""
    row_groups = max(chunks_per_sheet - _FIXED_CHUNKS, 1)
    for row in range(row_groups):
        top = 2 + row * 8
        group_id = _node(
            structure_docs, sheet, NodeType.ROW_GROUP, f"A{top}:D{top + 7}", top, top + 7, 1, 4
        )
        chunk_id = _chunk(
            chunk_docs,
            sheet,
            group_id,
            f"A{top}:D{top + 7}",
            ChunkType.ROW_GROUP,
            f"Row group {row}",
            "Projected revenue by quarter with the growth assumption applied.",
            embedding(),
        )
        if not first_chunk_id:
            first_chunk_id = chunk_id

    for cell in range(cells_per_sheet):
        row = 2 + cell
        _node(
            structure_docs,
            sheet,
            NodeType.CELL,
            f"C{row}",
            row,
            row,
            3,
            3,
            cached_value=float(cell),
        )

    # -- formula chain f0 <- f1 <- ... <- f5, with a cross-sheet edge on f1 ----------------------
    formulas: list[str] = []
    previous = _node_id(sheet, NodeType.FORMULA, "f0")
    _formula_node(structure_docs, sheet, "f0", "F0", 1, [])
    formulas.append(previous)
    for index in range(1, 6):
        current = _node_id(sheet, NodeType.FORMULA, f"f{index}")
        references = [
            Reference(
                target_node_id=previous,
                sheet_name=sheet.sheet_name,
                a1_range="C2",
                kind=ReferenceKind.CELL,
            ),
            Reference(
                target_node_id=_node_id(sheet, NodeType.CELL, f"c{index + 1}"),
                sheet_name=sheet.sheet_name,
                a1_range=f"C{index + 1}",
                kind=ReferenceKind.CELL,
            ),
        ]
        if index == 1:
            references.append(
                Reference(
                    target_node_id=node_id(
                        sheet.workbook_id, sheet.version, "cell", f"{sheet.next_sheet_name}!c2"
                    ),
                    sheet_name=sheet.next_sheet_name,
                    a1_range="C2",
                    kind=ReferenceKind.CELL,
                )
            )
        if index == 5:
            # A deeper chain that is not itself a chunk, so expansion depth actually grows.
            references.append(
                Reference(
                    target_node_id=_node_id(sheet, NodeType.FORMULA, "deep1"),
                    sheet_name=sheet.sheet_name,
                    a1_range="G1",
                    kind=ReferenceKind.CELL,
                )
            )
        _formula_node(structure_docs, sheet, f"f{index}", f"F{index}", index, references)
        formulas.append(current)
        previous = current

    deep_keys = [f"deep{i}" for i in range(1, 5)]
    for position, key in enumerate(deep_keys):
        deeper = (
            [
                Reference(
                    target_node_id=_node_id(sheet, NodeType.FORMULA, deep_keys[position + 1]),
                    sheet_name=sheet.sheet_name,
                    a1_range=f"G{position + 1}",
                    kind=ReferenceKind.CELL,
                )
            ]
            if position + 1 < len(deep_keys)
            else []
        )
        _formula_node(structure_docs, sheet, key, f"G{position + 1}", 10 + position, deeper)

    node_a = _node_id(sheet, NodeType.FORMULA, "cyc_a")
    node_b = _node_id(sheet, NodeType.FORMULA, "cyc_b")
    _formula_node(
        structure_docs,
        sheet,
        None,
        "E1",
        1,
        [
            Reference(
                target_node_id=node_b,
                sheet_name=sheet.sheet_name,
                a1_range="E1",
                kind=ReferenceKind.CELL,
            )
        ],
        identifier=node_a,
    )
    _formula_node(
        structure_docs,
        sheet,
        None,
        "E2",
        1,
        [
            Reference(
                target_node_id=node_a,
                sheet_name=sheet.sheet_name,
                a1_range="E2",
                kind=ReferenceKind.CELL,
            )
        ],
        identifier=node_b,
    )

    # Formula-summary chunks sit on the formula nodes; their titles match the query in two fields,
    # so they outrank the one-field row-group and column chunks and seed the expansion.
    for index in range(1, 6):
        _chunk(
            chunk_docs,
            sheet,
            _node_id(sheet, NodeType.FORMULA, f"f{index}"),
            f"F{index}",
            ChunkType.FORMULA_SUMMARY,
            f"Projected revenue formula F{index}",
            f"Projected revenue formula =SUM(C2:C9)*(1+rate_{sheet.sheet_name.lower()}).",
            embedding(),
        )
    for identifier, a1 in ((node_a, "E1"), (node_b, "E2")):
        _chunk(
            chunk_docs,
            sheet,
            identifier,
            a1,
            ChunkType.FORMULA_SUMMARY,
            "Projected revenue formula cycle",
            "Projected revenue formula with a circular reference between two formula nodes.",
            embedding(),
        )
    return formulas, first_chunk_id


def _node_id(sheet: _SheetContext, node_type: NodeType, key: str) -> str:
    return node_id(sheet.workbook_id, sheet.version, str(node_type), f"{sheet.sheet_name}!{key}")


def _node(
    structure_docs: list[tuple[str, dict[str, Any]]],
    sheet: _SheetContext,
    node_type: NodeType,
    a1: str,
    min_row: int,
    max_row: int,
    min_col: int,
    max_col: int,
    *,
    cached_value: Any = None,
    table_name: str | None = None,
    named_range: str | None = None,
) -> str:
    identifier = _node_id(sheet, node_type, a1)
    document = StructureDocument(
        node_id=identifier,
        workbook_id=sheet.workbook_id,
        version=sheet.version,
        node_type=node_type,
        sheet_id=sheet.sheet_id,
        sheet_name=sheet.sheet_name,
        a1_range=a1,
        row_span={"gte": min_row, "lte": max_row},
        column_span={"gte": min_col, "lte": max_col},
        cached_value=cached_value,
        table_name=table_name,
        named_range=named_range,
        acl_scope=tuple(sheet.acl),
    )
    structure_docs.append((identifier, document.model_dump(mode="json")))
    return identifier


def _formula_node(
    structure_docs: list[tuple[str, dict[str, Any]]],
    sheet: _SheetContext,
    key: str | None,
    a1: str,
    row: int,
    references: list[Reference],
    *,
    identifier: str | None = None,
) -> str:
    node = identifier or _node_id(sheet, NodeType.FORMULA, key or a1)
    document = StructureDocument(
        node_id=node,
        workbook_id=sheet.workbook_id,
        version=sheet.version,
        node_type=NodeType.FORMULA,
        sheet_id=sheet.sheet_id,
        sheet_name=sheet.sheet_name,
        a1_range=a1,
        row_span={"gte": row, "lte": row},
        column_span={"gte": 6, "lte": 6},
        formula=f"=SUM(C2:C9)*(1+rate_{sheet.sheet_name.lower()})",
        references=tuple(references),
        acl_scope=tuple(sheet.acl),
    )
    structure_docs.append((node, document.model_dump(mode="json")))
    return node


def _chunk(
    chunk_docs: list[tuple[str, dict[str, Any]]],
    sheet: _SheetContext,
    node: str,
    a1: str,
    chunk_type: ChunkType,
    title: str,
    content: str,
    embedding: list[float],
) -> str:
    identifier = node_id(
        sheet.workbook_id, sheet.version, "chunk", f"{chunk_type}:{sheet.sheet_name}:{title}:{a1}"
    )
    document = ChunkDocument(
        id=identifier,
        workbook_id=sheet.workbook_id,
        version=sheet.version,
        node_id=node,
        sheet_id=sheet.sheet_id,
        sheet_name=sheet.sheet_name,
        a1_range=a1,
        chunk_type=chunk_type,
        title=title,
        content=content,
        headers=("Metric", "2027", "2028", "2029", "2030"),
        embedding=tuple(embedding),
        embedding_model=_EMBEDDING_MODEL,
        acl_scope=tuple(sheet.acl),
        ingested_at=_INGESTED_AT,
    )
    chunk_docs.append((identifier, document.model_dump(mode="json")))
    return identifier


__all__ = ["DEFAULT_ACL_SCOPES", "Corpus", "build_corpus"]
