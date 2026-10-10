"""Reranking: the head of the fused list rescored by a (query, chunk) model, never truncated."""

from __future__ import annotations

import sys
import types

import pytest
from conftest import WB

from excel_rag.embedding import EmbeddingError
from excel_rag.models import SearchFilters, SearchRequest
from excel_rag.rerank import CrossEncoderReranker, OverlapReranker, build_reranker
from excel_rag.retrieval import Repository, RetrievalService
from excel_rag.settings import RerankSettings, Settings


class _Reversing:
    """Scores the texts in reverse of the order it is given them, and records each call."""

    model_name = "reversing"

    def __init__(self) -> None:
        self.calls: list[tuple[str, int]] = []

    def score(self, query: str, texts):
        self.calls.append((query, len(texts)))
        return [float(index) for index in range(len(texts))]


def _request(**overrides) -> SearchRequest:
    payload = {"query": "revenue", "filters": SearchFilters(workbook_ids=(WB,))}
    payload.update(overrides)
    return SearchRequest(**payload)


def test_a_reranker_reorders_the_fused_head(settings: Settings, client) -> None:
    plain = RetrievalService(Repository(client, settings), settings).search(_request(top_k=10))
    reranker = _Reversing()
    reranked = RetrievalService(Repository(client, settings), settings, reranker=reranker).search(
        _request(top_k=10)
    )
    assert len(plain.hits) > 1
    assert [hit.chunk_id for hit in reranked.hits] == [hit.chunk_id for hit in plain.hits][::-1]
    assert {hit.score_kind for hit in reranked.hits} == {"rerank"}
    assert reranker.calls == [("revenue", len(plain.hits))]


def test_the_window_is_never_smaller_than_top_k(client) -> None:
    settings = Settings(rerank=RerankSettings(window=1))
    reranker = _Reversing()
    service = RetrievalService(Repository(client, settings), settings, reranker=reranker)
    response = service.search(_request(top_k=3))
    assert len(response.hits) == min(3, reranker.calls[0][1])
    assert reranker.calls[0][1] >= len(response.hits)


def test_without_a_reranker_nothing_changes(settings: Settings, client) -> None:
    response = RetrievalService(Repository(client, settings), settings).search(_request())
    assert {hit.score_kind for hit in response.hits} == {"lexical"}


def test_overlap_double_scores_shared_tokens() -> None:
    scores = OverlapReranker().score("projected revenue", ["Revenue forecast", "Cost", ""])
    assert scores == [0.5, 0.0, 0.0]


def test_build_reranker_follows_settings() -> None:
    assert build_reranker(RerankSettings()) is None
    assert isinstance(build_reranker(RerankSettings(provider="overlap")), OverlapReranker)
    built = build_reranker(RerankSettings(provider="cross-encoder", model="m"))
    assert isinstance(built, CrossEncoderReranker)
    assert built.model_name == "m"


def test_cross_encoder_loads_lazily_once_and_scores_pairs(monkeypatch) -> None:
    created: list[tuple[str, dict]] = []

    class FakeCrossEncoder:
        def __init__(self, name: str, **kwargs) -> None:
            created.append((name, kwargs))

        def predict(self, pairs, **kwargs):
            return [float(len(text)) for _, text in pairs]

    module = types.ModuleType("sentence_transformers")
    module.CrossEncoder = FakeCrossEncoder  # type: ignore[attr-defined]
    monkeypatch.setitem(sys.modules, "sentence_transformers", module)
    reranker = CrossEncoderReranker("m", max_length=128, device="cpu")
    assert created == []
    assert reranker.score("q", ["ab", "abcd"]) == [2.0, 4.0]
    assert reranker.score("q", []) == []
    assert reranker.score("q", ["a"]) == [1.0]
    assert created == [("m", {"max_length": 128, "device": "cpu"})]


def test_cross_encoder_without_the_extra_says_how_to_fix_it(monkeypatch) -> None:
    monkeypatch.setitem(sys.modules, "sentence_transformers", None)
    with pytest.raises(EmbeddingError, match="embed"):
        CrossEncoderReranker("m").score("q", ["a"])


def test_the_http_path_reranks_when_configured(client) -> None:
    from fastapi.testclient import TestClient

    from excel_rag.api.deps import get_client
    from excel_rag.app import create_app

    settings = Settings(rerank=RerankSettings(provider="overlap"), embedding={"provider": "none"})
    app = create_app(settings)
    app.dependency_overrides[get_client] = lambda: client
    body = (
        TestClient(app)
        .post("/api/v1/search/excel", json={"query": "revenue", "filters": {"workbook_ids": [WB]}})
        .json()
    )
    assert body["hits"]
    assert {hit["score_kind"] for hit in body["hits"]} == {"rerank"}
