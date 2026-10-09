"""Ingestion: the workbook model turned into the documents Elasticsearch indexes.

Reading, region detection and formula edges live in :mod:`excel_rag.workbook`, shared with the
context renderer (:mod:`excel_rag.context`); this package adds the chunk and structure documents
and the indexer. The public entry point is :func:`ingest_workbook`, which reads a workbook into the
canonical model and returns the exact documents to index. Nothing here evaluates a formula,
executes a macro or refreshes an external link.
"""

from __future__ import annotations

from collections.abc import Sequence
from datetime import datetime
from pathlib import Path

from ..workbook import (
    IngestError,
    NamedRange,
    RawWorkbook,
    Region,
    RegionConfig,
    RegionKind,
    SheetModel,
    WorkbookModel,
    build_model,
    detect_regions,
    normalize_formula,
    read_workbook,
)
from .documents import (
    MAX_ROW_GROUPS_PER_REGION,
    MAX_TEXT_COLUMNS,
    ROW_GROUP_ROWS,
    IngestedWorkbook,
    build_documents,
)
from .indexer import Indexer, IndexResult, build_client, document_body

__all__ = [
    "MAX_ROW_GROUPS_PER_REGION",
    "MAX_TEXT_COLUMNS",
    "ROW_GROUP_ROWS",
    "IndexResult",
    "Indexer",
    "IngestError",
    "IngestedWorkbook",
    "NamedRange",
    "RawWorkbook",
    "Region",
    "RegionConfig",
    "RegionKind",
    "SheetModel",
    "WorkbookModel",
    "build_client",
    "build_documents",
    "build_model",
    "detect_regions",
    "document_body",
    "ingest_workbook",
    "normalize_formula",
    "read_workbook",
]


def ingest_workbook(
    path: str | Path,
    *,
    workbook_id: str,
    version: int,
    acl_scope: Sequence[str] = (),
    config: RegionConfig | None = None,
    ingested_at: datetime | None = None,
) -> IngestedWorkbook:
    """Read ``path`` and return the model plus every chunk and structure document it produces."""
    raw = read_workbook(path)
    model = build_model(
        raw, workbook_id=workbook_id, version=version, acl_scope=acl_scope, config=config
    )
    return build_documents(model, ingested_at=ingested_at)
