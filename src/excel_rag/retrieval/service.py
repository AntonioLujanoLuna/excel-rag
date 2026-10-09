"""Orchestration: one :class:`~excel_rag.models.SearchRequest` in, one ``SearchResponse`` out.

The service is the only place that knows the order of operations. It resolves the scope once (ACL
scopes and active versions), spends one primary query, re-ranks the candidate window, then -- only
if asked -- expands references under the budgets in :class:`~excel_rag.settings.BudgetSettings`.
Every bound comes from settings; nothing here is hardcoded.

The response is retrieval evidence. It never synthesises an answer, and ``took_ms``/``es_requests``
describe the work this request cost.
"""

from __future__ import annotations

import time

from ..models import (
    Hit,
    NodePayload,
    SearchRequest,
    SearchResponse,
    SourceRef,
)
from ..settings import Settings
from .expansion import ExpansionResult, expand
from .fusion import Candidate, RankedCandidate, matched_fields, rank
from .repository import Repository, Scope

#: The candidate window the lexical query returns before fusion re-ranks it. A multiple of ``top_k``
#: so a vector re-rank has something to promote, floored so a small ``top_k`` still gets a window.
CANDIDATE_WINDOW_FACTOR = 5
CANDIDATE_WINDOW_FLOOR = 50
CANDIDATE_WINDOW_CEILING = 500


class RetrievalService:
    """The retrieval workflow, independent of HTTP."""

    def __init__(self, repository: Repository, settings: Settings) -> None:
        self._repository = repository
        self._settings = settings

    def search(
        self,
        request: SearchRequest,
        *,
        query_vector: tuple[float, ...] | None = None,
        query_embedding_model: str | None = None,
    ) -> SearchResponse:
        """Run one search. ``query_vector`` is optional because the frozen HTTP request has no
        vector field; the endpoint runs lexical-only, and a caller that embeds the query passes a
        vector here to exercise the fusion path."""
        started = time.perf_counter()
        calls_before = self._repository.es_requests()

        scope = self._repository.resolve_scope(
            workbook_ids=request.filters.workbook_ids,
            acl_scopes=request.filters.acl_scopes,
        )
        window = min(
            max(request.top_k * CANDIDATE_WINDOW_FACTOR, CANDIDATE_WINDOW_FLOOR),
            CANDIDATE_WINDOW_CEILING,
        )
        scored = self._repository.search_chunks(
            request.query,
            filters=request.filters,
            scope=scope,
            size=window,
        )
        candidates = [Candidate(chunk=chunk, lexical_score=score) for chunk, score in scored]
        ranked = rank(
            candidates,
            query_vector=query_vector,
            query_model=query_embedding_model,
        )
        top = ranked[: request.top_k]

        expansion = ExpansionResult()
        if request.include_structure:
            expansion = self._expand(request, scope, top)

        related = _related_ids(top, expansion)
        hits = tuple(
            _hit(
                ranked_candidate,
                related.get(ranked_candidate.candidate.chunk.node_id, ()),
                request.query,
            )
            for ranked_candidate in top
        )

        took_ms = (time.perf_counter() - started) * 1000.0
        return SearchResponse(
            hits=hits,
            nodes=expansion.nodes,
            unresolved_references=expansion.unresolved,
            truncation=expansion.truncation,
            took_ms=took_ms,
            es_requests=self._repository.es_requests() - calls_before,
        )

    def _expand(
        self, request: SearchRequest, scope: Scope, top: list[RankedCandidate]
    ) -> ExpansionResult:
        budgets = self._settings.budgets
        explicit_depth = "reference_depth" in request.model_fields_set
        requested_depth = request.reference_depth if explicit_depth else budgets.reference_depth
        max_depth = (
            min(requested_depth, budgets.reference_depth) if request.expand_references else 0
        )

        explicit_nodes = "max_related_nodes" in request.model_fields_set
        requested_nodes = request.max_related_nodes if explicit_nodes else budgets.max_related_nodes
        max_nodes = min(requested_nodes, budgets.max_related_nodes)

        seeds = list(dict.fromkeys(candidate.candidate.chunk.node_id for candidate in top))
        deadline = time.monotonic() + budgets.timeout_seconds
        return expand(
            seed_node_ids=seeds,
            repository=self._repository,
            scope=scope,
            max_depth=max_depth,
            max_nodes=max_nodes,
            max_bytes=budgets.max_payload_bytes,
            deadline=deadline,
        )


def _related_ids(
    top: list[RankedCandidate], expansion: ExpansionResult
) -> dict[str, tuple[str, ...]]:
    """Map each hit's node id to the related node ids the expansion actually returned."""
    known = set(expansion.nodes)
    related: dict[str, tuple[str, ...]] = {}
    for ranked in top:
        node_id = ranked.candidate.chunk.node_id
        payload: NodePayload | None = expansion.nodes.get(node_id)
        if payload is None:
            related[node_id] = ()
            continue
        targets = sorted({reference.target_node_id for reference in payload.references} & known)
        related[node_id] = tuple(targets)
    return related


def _hit(ranked: RankedCandidate, related_node_ids: tuple[str, ...], query: str) -> Hit:
    chunk = ranked.candidate.chunk
    return Hit(
        chunk_id=chunk.id,
        score=ranked.score,
        content=chunk.content,
        source=SourceRef(
            workbook_id=chunk.workbook_id,
            version=chunk.version,
            sheet=chunk.sheet_name,
            a1_range=chunk.a1_range,
        ),
        node_id=chunk.node_id,
        title=chunk.title,
        headers=chunk.headers,
        matched_fields=matched_fields(query, chunk),
        related_node_ids=related_node_ids,
    )


__all__ = [
    "CANDIDATE_WINDOW_CEILING",
    "CANDIDATE_WINDOW_FACTOR",
    "CANDIDATE_WINDOW_FLOOR",
    "RetrievalService",
]
