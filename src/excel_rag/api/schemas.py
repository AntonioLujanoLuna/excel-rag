"""Response schemas for the endpoints that are not the frozen search contract.

``SearchResponse`` lives in :mod:`excel_rag.models` because both workstreams code against it. The
optional inspection endpoints and the error envelope are the retrieval side's own, so they are
declared here and kept in one place.
"""

from __future__ import annotations

from typing import Any

from pydantic import BaseModel

from ..models import NodePayload, UnresolvedReference


class ErrorBody(BaseModel):
    type: str
    message: str
    details: tuple[Any, ...] | None = None


class ErrorResponse(BaseModel):
    """The one error shape: every non-2xx response is ``{"error": {...}}``."""

    error: ErrorBody


class StructureResponse(BaseModel):
    workbook_id: str
    version: int
    nodes: tuple[NodePayload, ...]
    took_ms: float
    es_requests: int


class RangeResponse(BaseModel):
    workbook_id: str
    sheet_name: str
    a1: str
    version: int
    nodes: tuple[NodePayload, ...]
    unresolved_references: tuple[UnresolvedReference, ...] = ()
    took_ms: float
    es_requests: int


class HealthResponse(BaseModel):
    status: str
    version: str
    use_live_elasticsearch: bool
    indices: dict[str, int]


def error_body(
    error_type: str, message: str, details: tuple[Any, ...] | None = None
) -> dict[str, Any]:
    """Build the standard error envelope as a plain dict (for exception handlers)."""
    payload: dict[str, Any] = {"type": error_type, "message": message}
    if details is not None:
        payload["details"] = details
    return {"error": payload}


__all__ = [
    "ErrorBody",
    "ErrorResponse",
    "HealthResponse",
    "RangeResponse",
    "StructureResponse",
    "error_body",
]
