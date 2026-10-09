"""``POST /api/v1/search/excel`` -- the primary retrieval endpoint.

It is a thin shell over :meth:`~excel_rag.retrieval.service.RetrievalService.search`: parse the
frozen request model, run the workflow, return the frozen response model. No generation, no
follow-up turns, no server state.
"""

from __future__ import annotations

from typing import Annotated

from fastapi import APIRouter, Depends

from ..models import SearchRequest, SearchResponse
from ..retrieval import RetrievalService
from .deps import get_service, require_token

router = APIRouter(
    prefix="/api/v1",
    tags=["search"],
    dependencies=[Depends(require_token)],
)


@router.post("/search/excel", response_model=SearchResponse)
def search_excel(
    request: SearchRequest,
    service: Annotated[RetrievalService, Depends(get_service)],
) -> SearchResponse:
    """Search the semantic index and, if asked, expand related structural nodes.

    The HTTP request carries no query vector -- the frozen :class:`~excel_rag.models.SearchRequest`
    has no such field -- so this endpoint runs the lexical path. A caller that embeds the query
    itself reaches the hybrid path through :meth:`RetrievalService.search`.
    """
    return service.search(request)


__all__ = ["router"]
