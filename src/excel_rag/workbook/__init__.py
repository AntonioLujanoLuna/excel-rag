"""A workbook read into a structured model: no index, no store, no network.

This is the half of excel-rag that both consumers share. :mod:`excel_rag.ingest` turns the model
into Elasticsearch documents; :mod:`excel_rag.context` renders it for a conversation's context
window. Everything here is pure: openpyxl reading with the macro, link, zip-bomb and dimension
guards (:mod:`.reader`), region and header detection (:mod:`.regions`), static formula references
(:mod:`.formulas`), and the assembled :class:`WorkbookModel` (:mod:`.build`).

Nothing here evaluates a formula, executes a macro or refreshes an external link.
"""

from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path

from .build import build_model, normalize_formula
from .canonical import NamedRange, Region, RegionKind, SheetModel, WorkbookModel
from .errors import IngestError, WorkbookError
from .reader import RawWorkbook, read_workbook
from .regions import RegionConfig, detect_regions


def load_workbook(
    source: str | Path | bytes,
    *,
    name: str | None = None,
    workbook_id: str = "workbook",
    version: int = 1,
    acl_scope: Sequence[str] = (),
    config: RegionConfig | None = None,
) -> WorkbookModel:
    """Read a workbook from a path or from bytes (an upload) into the structured model."""
    raw = read_workbook(source, name=name)
    return build_model(
        raw, workbook_id=workbook_id, version=version, acl_scope=acl_scope, config=config
    )


__all__ = [
    "IngestError",
    "NamedRange",
    "RawWorkbook",
    "Region",
    "RegionConfig",
    "RegionKind",
    "SheetModel",
    "WorkbookError",
    "WorkbookModel",
    "build_model",
    "detect_regions",
    "load_workbook",
    "normalize_formula",
    "read_workbook",
]
