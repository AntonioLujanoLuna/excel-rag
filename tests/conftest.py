"""Shared fixtures: a small explicit corpus built straight from the models.

Nothing here imports ``excel_rag.ingest`` (another workstream owns it) or the benchmark corpus; the
documents are constructed field by field so a test failure points at the retrieval layer, not at a
generator. The corpus covers the retrieval pitfalls on purpose: a stale version, a per-scope chunk,
a related node one scope may read and another may not, a dangling reference and a formula cycle.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import UTC, datetime
from typing import Any

import pytest
from fastapi.testclient import TestClient

from excel_rag.api.deps import get_client
from excel_rag.app import create_app
from excel_rag.es import INDEX_CHUNKS, INDEX_STRUCTURE, INDEX_VERSIONS
from excel_rag.fake_es import InMemoryElasticsearch, in_memory_client
from excel_rag.models import (
    A1Range,
    ActiveVersionManifest,
    ChunkDocument,
    ChunkType,
    NodeType,
    Reference,
    ReferenceKind,
    StructureDocument,
    UnresolvedReason,
    UnresolvedReference,
    node_id,
)
from excel_rag.settings import Settings

# The suite runs without model weights: tests that exercise vectors pass a hashing embedder, or
# opt in to the real model in tests/live. Every `Settings()` is built at test time, after this.
os.environ.setdefault("EXCEL_RAG_EMBEDDING__PROVIDER", "none")

FINANCE = "finance-team"
HR = "hr-team"
EXEC = "exec-only"
EMBED = "test-embed-v1"
INGESTED = datetime(2026, 10, 9, tzinfo=UTC)

WB = "wb42"
OTHER_WB = "wb77"

# Structural node ids used across the tests.
N_REGION = node_id(WB, 1, "region", "Forecast!A12")
N_CELL = node_id(WB, 1, "cell", "Assumptions!C7")
N_SECRET = node_id(WB, 1, "cell", "Assumptions!D7")
N_FORMULA = node_id(WB, 1, "formula", "Forecast!F18")
N_FORMULA2 = node_id(WB, 1, "formula", "Forecast!F19")
N_GHOST = node_id(WB, 1, "cell", "Ghost!Z9")
N_FORECAST_CELL = node_id(WB, 1, "cell", "Forecast!B15")
N_RANGE = node_id(WB, 1, "range", "Actuals!D2:D500")
N_OTHER_TABLE = node_id(OTHER_WB, 1, "table", "Sensitivity!A1")
N_STALE = node_id(WB, 2, "region", "Forecast!A12")

C_REVENUE = f"{WB}:v1:chunk:region:forecast"
C_HR = f"{WB}:v1:chunk:region:hr"
C_SECRET = f"{WB}:v1:chunk:region:exec"
C_FORMULA = f"{WB}:v1:chunk:formula:forecast!f18"
C_STALE = f"{WB}:v2:chunk:region:forecast"
C_OTHER = f"{OTHER_WB}:v1:chunk:table:sensitivity"


@dataclass(frozen=True)
class Seed:
    chunks: tuple[tuple[str, dict[str, Any]], ...]
    structure: tuple[tuple[str, dict[str, Any]], ...]
    versions: tuple[tuple[str, dict[str, Any]], ...]


def _chunk(
    *,
    identifier: str,
    workbook: str,
    version: int,
    node: str,
    sheet: str,
    a1: str,
    chunk_type: ChunkType,
    title: str,
    content: str,
    acl: tuple[str, ...],
    embedding: tuple[float, ...] | None = None,
    embedding_model: str | None = EMBED,
) -> tuple[str, dict[str, Any]]:
    document = ChunkDocument(
        id=identifier,
        workbook_id=workbook,
        version=version,
        node_id=node,
        sheet_id=node_id(workbook, version, "sheet", sheet),
        sheet_name=sheet,
        a1_range=a1,
        chunk_type=chunk_type,
        title=title,
        content=content,
        headers=("Metric", "2027", "2028"),
        embedding=embedding,
        embedding_model=embedding_model,
        acl_scope=acl,
        ingested_at=INGESTED,
    )
    return identifier, document.model_dump(mode="json")


def _node(
    *,
    identifier: str,
    workbook: str,
    version: int,
    node_type: NodeType,
    sheet: str,
    a1: str,
    row_span: tuple[int, int],
    column_span: tuple[int, int],
    acl: tuple[str, ...],
    references: tuple[Reference, ...] = (),
    unresolved: tuple[UnresolvedReference, ...] = (),
    cached_value: Any = None,
    formula: str | None = None,
) -> tuple[str, dict[str, Any]]:
    document = StructureDocument(
        node_id=identifier,
        workbook_id=workbook,
        version=version,
        node_type=node_type,
        sheet_id=node_id(workbook, version, "sheet", sheet),
        sheet_name=sheet,
        a1_range=a1,
        row_span={"gte": row_span[0], "lte": row_span[1]},
        column_span={"gte": column_span[0], "lte": column_span[1]},
        formula=formula,
        cached_value=cached_value,
        references=references,
        unresolved_references=unresolved,
        acl_scope=acl,
    )
    return identifier, document.model_dump(mode="json")


def build_seed() -> Seed:
    """The corpus: two workbooks, one stale version, four ACL scopes, a ghost ref and a cycle."""
    chunks = [
        _chunk(
            identifier=C_REVENUE,
            workbook=WB,
            version=1,
            node=N_REGION,
            sheet="Forecast",
            a1="A12:D25",
            chunk_type=ChunkType.TABLE,
            title="Revenue forecast",
            content="Projected revenue forecast by year for the group.",
            acl=(FINANCE,),
            embedding=(1.0, 0.0, 0.0, 0.0),
        ),
        _chunk(
            identifier=C_HR,
            workbook=WB,
            version=1,
            node=node_id(WB, 1, "region", "HR!A1"),
            sheet="HR",
            a1="A1:C9",
            chunk_type=ChunkType.REGION,
            title="Headcount plan",
            content="Headcount plan and hiring assumptions.",
            acl=(HR,),
            embedding=(0.0, 1.0, 0.0, 0.0),
        ),
        _chunk(
            identifier=C_SECRET,
            workbook=WB,
            version=1,
            node=N_SECRET,
            sheet="Assumptions",
            a1="D7",
            chunk_type=ChunkType.REGION,
            title="Executive override",
            content="Executive revenue override, restricted to the exec scope.",
            acl=(EXEC,),
            embedding=(0.0, 0.0, 1.0, 0.0),
        ),
        _chunk(
            identifier=C_FORMULA,
            workbook=WB,
            version=1,
            node=N_FORMULA,
            sheet="Forecast",
            a1="F18",
            chunk_type=ChunkType.FORMULA_SUMMARY,
            title="Revenue growth formula",
            content="Projected revenue formula using the growth rate cell.",
            acl=(FINANCE,),
            embedding=(0.9, 0.1, 0.0, 0.0),
        ),
        _chunk(
            identifier=C_STALE,
            workbook=WB,
            version=2,
            node=N_STALE,
            sheet="Forecast",
            a1="A12:D25",
            chunk_type=ChunkType.TABLE,
            title="Revenue forecast (v2)",
            content="Projected revenue forecast, replacement version not yet active.",
            acl=(FINANCE,),
            embedding=(1.0, 0.0, 0.0, 0.0),
        ),
        _chunk(
            identifier=C_OTHER,
            workbook=OTHER_WB,
            version=1,
            node=N_OTHER_TABLE,
            sheet="Sensitivity",
            a1="A1:D9",
            chunk_type=ChunkType.TABLE,
            title="Revenue sensitivity",
            content="Projected revenue sensitivity analysis for the board.",
            acl=(FINANCE,),
            embedding=(0.5, 0.5, 0.0, 0.0),
        ),
    ]

    structure = [
        _node(
            identifier=N_REGION,
            workbook=WB,
            version=1,
            node_type=NodeType.REGION,
            sheet="Forecast",
            a1="A12:D25",
            row_span=(12, 25),
            column_span=(1, 4),
            acl=(FINANCE,),
            references=(
                Reference(
                    target_node_id=N_CELL,
                    sheet_name="Assumptions",
                    a1_range="C7",
                    kind=ReferenceKind.CELL,
                ),
                Reference(
                    target_node_id=N_SECRET,
                    sheet_name="Assumptions",
                    a1_range="D7",
                    kind=ReferenceKind.CELL,
                ),
                Reference(
                    target_node_id=N_GHOST,
                    sheet_name="Ghost",
                    a1_range="Z9",
                    kind=ReferenceKind.CELL,
                ),
                Reference(
                    target_node_id=N_FORMULA,
                    sheet_name="Forecast",
                    a1_range="F18",
                    kind=ReferenceKind.CELL,
                ),
            ),
        ),
        _node(
            identifier=N_CELL,
            workbook=WB,
            version=1,
            node_type=NodeType.CELL,
            sheet="Assumptions",
            a1="C7",
            row_span=(7, 7),
            column_span=(3, 3),
            acl=(FINANCE,),
            cached_value=0.05,
        ),
        _node(
            identifier=N_SECRET,
            workbook=WB,
            version=1,
            node_type=NodeType.CELL,
            sheet="Assumptions",
            a1="D7",
            row_span=(7, 7),
            column_span=(4, 4),
            acl=(EXEC,),
            cached_value=999.0,
        ),
        _node(
            identifier=N_FORMULA,
            workbook=WB,
            version=1,
            node_type=NodeType.FORMULA,
            sheet="Forecast",
            a1="F18",
            row_span=(18, 18),
            column_span=(6, 6),
            acl=(FINANCE,),
            formula="=SUM(Actuals!D2:D500)*(1+Assumptions!C7)",
            references=(
                Reference(
                    target_node_id=N_FORMULA2,
                    sheet_name="Forecast",
                    a1_range="F19",
                    kind=ReferenceKind.CELL,
                ),
            ),
            unresolved=(
                UnresolvedReference(
                    reference_text="INDIRECT(A1)",
                    reason=UnresolvedReason.INDIRECT,
                    detail="dynamic",
                ),
            ),
        ),
        _node(
            identifier=N_FORMULA2,
            workbook=WB,
            version=1,
            node_type=NodeType.FORMULA,
            sheet="Forecast",
            a1="F19",
            row_span=(19, 19),
            column_span=(6, 6),
            acl=(FINANCE,),
            formula="=F18*2",
            references=(
                Reference(
                    target_node_id=N_FORMULA,
                    sheet_name="Forecast",
                    a1_range="F18",
                    kind=ReferenceKind.CELL,
                ),
            ),
        ),
        _node(
            identifier=N_FORECAST_CELL,
            workbook=WB,
            version=1,
            node_type=NodeType.CELL,
            sheet="Forecast",
            a1="B15",
            row_span=(15, 15),
            column_span=(2, 2),
            acl=(FINANCE,),
            cached_value=12.0,
        ),
        _node(
            identifier=N_RANGE,
            workbook=WB,
            version=1,
            node_type=NodeType.RANGE,
            sheet="Actuals",
            a1="D2:D500",
            row_span=(2, 500),
            column_span=(4, 4),
            acl=(FINANCE,),
        ),
        _node(
            identifier=N_OTHER_TABLE,
            workbook=OTHER_WB,
            version=1,
            node_type=NodeType.TABLE,
            sheet="Sensitivity",
            a1="A1:D9",
            row_span=(1, 9),
            column_span=(1, 4),
            acl=(FINANCE,),
        ),
    ]

    versions = [
        (WB, ActiveVersionManifest(workbook_id=WB, active_version=1).model_dump(mode="json")),
        (
            OTHER_WB,
            ActiveVersionManifest(workbook_id=OTHER_WB, active_version=1).model_dump(mode="json"),
        ),
    ]
    return Seed(chunks=tuple(chunks), structure=tuple(structure), versions=tuple(versions))


def index_seed(client: InMemoryElasticsearch, seed: Seed | None = None) -> Seed:
    seed = seed or build_seed()
    client.bulk_index(INDEX_CHUNKS, seed.chunks, refresh=True)
    client.bulk_index(INDEX_STRUCTURE, seed.structure, refresh=True)
    client.bulk_index(INDEX_VERSIONS, seed.versions, refresh=True)
    return seed


@pytest.fixture
def settings() -> Settings:
    return Settings()


@pytest.fixture
def empty_client(settings: Settings) -> InMemoryElasticsearch:
    """A client whose three indices already exist but hold nothing."""
    return in_memory_client(settings)


@pytest.fixture
def client(settings: Settings) -> InMemoryElasticsearch:
    instance = in_memory_client(settings)
    index_seed(instance)
    return instance


@pytest.fixture
def app(settings: Settings, client: InMemoryElasticsearch) -> Any:
    """The real app with its client dependency overridden by the seeded in-memory client."""
    application = create_app(settings)
    application.dependency_overrides[get_client] = lambda: client
    return application


@pytest.fixture
def api(app: Any) -> TestClient:
    return TestClient(app)


def range_of(sheet: str, a1: str) -> A1Range:
    return A1Range.parse(sheet, a1)


__all__ = [
    "C_FORMULA",
    "C_HR",
    "C_OTHER",
    "C_REVENUE",
    "C_SECRET",
    "C_STALE",
    "EMBED",
    "EXEC",
    "FINANCE",
    "HR",
    "N_CELL",
    "N_FORECAST_CELL",
    "N_FORMULA",
    "N_FORMULA2",
    "N_GHOST",
    "N_OTHER_TABLE",
    "N_RANGE",
    "N_REGION",
    "N_SECRET",
    "N_STALE",
    "OTHER_WB",
    "WB",
    "Seed",
    "build_seed",
    "index_seed",
    "range_of",
]
