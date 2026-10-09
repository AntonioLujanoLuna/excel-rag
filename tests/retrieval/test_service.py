"""Service tests: the orchestration order, the response shape, and the fusion hook.

The HTTP path carries no vector, so the hybrid path is reached here by calling the service with one
-- exactly the seam the docs describe.
"""

from __future__ import annotations

import pytest
from conftest import (
    C_FORMULA,
    C_REVENUE,
    C_SECRET,
    C_STALE,
    EMBED,
    FINANCE,
    N_CELL,
    N_REGION,
    N_SECRET,
    WB,
)

from excel_rag.models import SearchFilters, SearchRequest
from excel_rag.retrieval import Repository, RetrievalService, UnknownWorkbook
from excel_rag.settings import Settings


def _service(settings: Settings, client) -> RetrievalService:
    return RetrievalService(Repository(client, settings), settings)


def _request(**overrides) -> SearchRequest:
    payload = {"query": "revenue", "filters": SearchFilters(workbook_ids=(WB,))}
    payload.update(overrides)
    return SearchRequest(**payload)


class TestSearch:
    def test_returns_hits_with_source_provenance(self, settings: Settings, client) -> None:
        response = _service(settings, client).search(_request())
        assert response.hits
        hit = next(hit for hit in response.hits if hit.chunk_id == C_REVENUE)
        assert hit.source.workbook_id == WB
        assert hit.source.version == 1
        assert hit.source.sheet == "Forecast"
        assert hit.source.a1_range == "A12:D25"
        assert hit.matched_fields
        assert response.took_ms >= 0.0
        assert response.es_requests >= 1

    def test_no_matches_is_an_empty_but_well_formed_response(
        self, settings: Settings, client
    ) -> None:
        response = _service(settings, client).search(_request(query="zzzzzzzz"))
        assert response.hits == ()
        assert response.nodes == {}
        assert not response.truncation.truncated

    def test_include_structure_false_skips_the_node_fetch(self, settings: Settings, client) -> None:
        response = _service(settings, client).search(_request(include_structure=False))
        assert response.nodes == {}

    def test_expansion_returns_related_nodes_and_hides_denied_ones(
        self, settings: Settings, client
    ) -> None:
        response = _service(settings, client).search(
            _request(
                filters=SearchFilters(workbook_ids=(WB,), acl_scopes=(FINANCE,)),
                include_structure=True,
                expand_references=True,
                reference_depth=1,
            )
        )
        assert N_REGION in response.nodes
        assert N_CELL in response.nodes
        assert N_SECRET not in response.nodes
        related = {node_id for hit in response.hits for node_id in hit.related_node_ids}
        assert N_CELL in related

    def test_a_request_never_mixes_versions(self, settings: Settings, client) -> None:
        response = _service(settings, client).search(
            SearchRequest(query="revenue", include_structure=False)
        )
        assert C_STALE not in {hit.chunk_id for hit in response.hits}
        assert {hit.source.version for hit in response.hits} == {1}

    def test_an_unknown_workbook_filter_is_refused(self, settings: Settings, client) -> None:
        with pytest.raises(UnknownWorkbook):
            _service(settings, client).search(
                SearchRequest(query="revenue", filters=SearchFilters(workbook_ids=("ghost-wb",)))
            )

    def test_top_k_windows_the_hits(self, settings: Settings, client) -> None:
        response = _service(settings, client).search(_request(top_k=1))
        assert len(response.hits) == 1


class TestFusionHook:
    def test_lexical_run_reports_the_index_score(self, settings: Settings, client) -> None:
        response = _service(settings, client).search(_request())
        assert response.hits[0].score == pytest.approx(2.0)
        assert response.hits[0].chunk_id == C_FORMULA

    def test_a_query_vector_switches_to_fused_scores(self, settings: Settings, client) -> None:
        response = _service(settings, client).search(
            _request(), query_vector=(1.0, 0.0, 0.0, 0.0), query_embedding_model=EMBED
        )
        # RRF scores live on a different scale from the index's ordinal score.
        assert 0.0 < response.hits[0].score < 0.1

    def test_a_query_vector_can_promote_a_low_lexical_hit(self, settings: Settings, client) -> None:
        lexical = _service(settings, client).search(_request())
        fused = _service(settings, client).search(
            _request(top_k=3), query_vector=(0.0, 0.0, 1.0, 0.0), query_embedding_model=EMBED
        )
        lexical_ids = [hit.chunk_id for hit in lexical.hits]
        fused_ids = [hit.chunk_id for hit in fused.hits]
        assert lexical_ids.index(C_SECRET) == len(lexical_ids) - 1
        assert fused_ids.index(C_SECRET) < len(fused_ids) - 1
