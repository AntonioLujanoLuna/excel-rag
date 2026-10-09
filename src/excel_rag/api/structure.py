"""Direct-inspection endpoints: workbook structure, A1 range intersection, and dependents.

These answer deterministic questions without a semantic query, which is the only reason they exist:
a caller that already knows the coordinate should not have to phrase a question. Range overlap is
answered by ``integer_range`` intersection on ``row_span``/``column_span`` -- never by walking
cells -- and both routes go through the same scope filter as search, with the scopes taken from the
caller (see :class:`~excel_rag.api.deps.Caller`), so an unknown workbook is a 404 and an
unauthorized node is never returned.
"""

from __future__ import annotations

import time
from typing import Annotated

from fastapi import APIRouter, Depends, HTTPException, Query

from ..models import A1Range, NodeType, RangeQuery, StructureQuery
from ..retrieval import Repository, UnknownWorkbook, node_payload
from .deps import Caller, get_caller, get_repository
from .schemas import RangeResponse, StructureResponse

router = APIRouter(
    prefix="/api/v1",
    tags=["structure"],
    dependencies=[Depends(get_caller)],
)


@router.get("/excel/{workbook_id}/structure", response_model=StructureResponse)
def workbook_structure(
    workbook_id: str,
    repository: Annotated[Repository, Depends(get_repository)],
    caller: Annotated[Caller, Depends(get_caller)],
    sheet_names: Annotated[list[str] | None, Query()] = None,
    node_types: Annotated[list[NodeType] | None, Query()] = None,
    limit: Annotated[int, Query(ge=1, le=2000)] = 200,
    acl_scopes: Annotated[list[str] | None, Query()] = None,
) -> StructureResponse:
    """Structural nodes for one workbook, at its active version, filtered by scope.

    ``StructureQuery`` validates the bounds and carries the facets; the repository applies the ACL
    and version filters.
    """
    started = time.perf_counter()
    calls_before = repository.es_requests()
    version = repository.active_version(workbook_id)
    if version is None:
        raise UnknownWorkbook(workbook_id)
    query = StructureQuery(
        workbook_ids=(workbook_id,),
        sheet_names=tuple(sheet_names or ()),
        node_types=tuple(node_types or ()),
        limit=limit,
        acl_scopes=tuple(acl_scopes or ()),
    )
    scope = repository.resolve_scope(
        workbook_ids=query.workbook_ids, acl_scopes=caller.scopes_for(query.acl_scopes)
    )
    documents = repository.query_structure(
        scope=scope,
        workbook_ids=query.workbook_ids,
        sheet_names=query.sheet_names,
        node_types=tuple(str(kind) for kind in query.node_types),
        limit=query.limit,
    )
    return StructureResponse(
        workbook_id=workbook_id,
        version=version,
        nodes=tuple(node_payload(document) for document in documents),
        took_ms=(time.perf_counter() - started) * 1000.0,
        es_requests=repository.es_requests() - calls_before,
    )


@router.post("/excel/range", response_model=RangeResponse)
def range_nodes(
    query: RangeQuery,
    repository: Annotated[Repository, Depends(get_repository)],
    caller: Annotated[Caller, Depends(get_caller)],
) -> RangeResponse:
    """Nodes whose spans intersect ``query.a1``, resolved by range intersection."""
    return _region_lookup(query, repository, caller, dependents=False)


@router.post("/excel/dependents", response_model=RangeResponse)
def dependents(
    query: RangeQuery,
    repository: Annotated[Repository, Depends(get_repository)],
    caller: Annotated[Caller, Depends(get_caller)],
) -> RangeResponse:
    """Nodes whose formulas (or named ranges) read a cell of ``query.a1``: the reverse edges.

    "If I change ``Assumptions!C7``, what moves?" Answered by intersecting the spans on each
    nested reference edge, so a formula reading ``D2:D500`` is a dependent of ``D100``.
    """
    return _region_lookup(query, repository, caller, dependents=True)


def _region_lookup(
    query: RangeQuery, repository: Repository, caller: Caller, *, dependents: bool
) -> RangeResponse:
    """Parse the rectangle, pin the active version and the caller's scopes, run one span query."""
    started = time.perf_counter()
    calls_before = repository.es_requests()
    try:
        region = A1Range.parse(query.sheet_name, query.a1)
    except ValueError as exc:
        raise HTTPException(
            status_code=400,
            detail={"type": "invalid_range", "message": str(exc)},
        ) from exc
    version = repository.active_version(query.workbook_id)
    if version is None:
        raise UnknownWorkbook(query.workbook_id)
    scope = repository.resolve_scope(
        workbook_ids=(query.workbook_id,), acl_scopes=caller.scopes_for(query.acl_scopes)
    )
    lookup = repository.query_dependents if dependents else repository.query_range
    documents = lookup(
        scope=scope,
        workbook_id=query.workbook_id,
        sheet_name=query.sheet_name,
        region=region,
        node_types=tuple(str(kind) for kind in query.node_types),
        limit=query.limit,
    )
    unresolved = tuple(ref for document in documents for ref in document.unresolved_references)
    return RangeResponse(
        workbook_id=query.workbook_id,
        sheet_name=query.sheet_name,
        a1=region.a1,
        version=version,
        nodes=tuple(node_payload(document) for document in documents),
        unresolved_references=unresolved,
        took_ms=(time.perf_counter() - started) * 1000.0,
        es_requests=repository.es_requests() - calls_before,
    )


__all__ = ["router"]
