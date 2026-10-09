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
from .deps import Caller, get_caller, get_service

router = APIRouter(
    prefix="/api/v1",
    tags=["search"],
    dependencies=[Depends(get_caller)],
)


@router.post("/search/excel", response_model=SearchResponse)
def search_excel(
    request: SearchRequest,
    service: Annotated[RetrievalService, Depends(get_service)],
    caller: Annotated[Caller, Depends(get_caller)],
) -> SearchResponse:
    """Search the semantic index and, if asked, expand related structural nodes.

    The HTTP request carries no query vector -- the frozen :class:`~excel_rag.models.SearchRequest`
    has no such field -- so this endpoint runs the lexical path. A caller that embeds the query
    itself reaches the hybrid path through :meth:`RetrievalService.search`.

    ``filters.acl_scopes`` is checked against the caller: a restricted caller may narrow its scopes
    but never name one it does not hold, and is filtered by all of its scopes when it names none.
    """
    scopes = caller.scopes_for(request.filters.acl_scopes)
    scoped = request.model_copy(
        update={"filters": request.filters.model_copy(update={"acl_scopes": scopes})}
    )
    return service.search(scoped)


__all__ = ["router"]
