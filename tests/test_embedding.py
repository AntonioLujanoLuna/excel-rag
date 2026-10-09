"""Embedders: the prompts, the dimensionality checks, and the hashing double's guarantees.

The sentence-transformers wrapper is exercised against a stand-in module here, so the suite needs
no weights; ``tests/live/test_mdenseon.py`` runs the real ``lightonai/mDenseOn``.
"""

from __future__ import annotations

import math
from types import SimpleNamespace
from typing import Any

import pytest

from excel_rag import embedding as embedding_module
from excel_rag.embedding import (
    EmbeddingError,
    HashingEmbedder,
    SentenceTransformerEmbedder,
    build_embedder,
    chunk_text,
)
from excel_rag.settings import EmbeddingSettings


class _FakeModel:
    def __init__(self, name: str, device: str | None = None, dims: int = 4) -> None:
        self.name = name
        self.device = device
        self.dims = dims
        self.prompts = {"query": "query: ", "document": "document: "}
        self.max_seq_length = 8192
        self.calls: list[tuple[list[str], str, bool]] = []

    def get_embedding_dimension(self) -> int:
        return self.dims

    def encode(self, texts: list[str], *, prompt_name: str, **kwargs: Any) -> list[list[float]]:
        self.calls.append((texts, prompt_name, kwargs["normalize_embeddings"]))
        return [[1.0] + [0.0] * (self.dims - 1) for _ in texts]


@pytest.fixture
def fake_st(monkeypatch: pytest.MonkeyPatch) -> list[_FakeModel]:
    built: list[_FakeModel] = []

    def factory(name: str, device: str | None = None) -> _FakeModel:
        model = _FakeModel(name, device)
        built.append(model)
        return model

    real_import = embedding_module.importlib.import_module

    def import_module(name: str) -> Any:
        if name == "sentence_transformers":
            return SimpleNamespace(SentenceTransformer=factory)
        return real_import(name)

    monkeypatch.setattr(embedding_module.importlib, "import_module", import_module)
    return built


class TestSentenceTransformerEmbedder:
    def test_queries_and_documents_use_their_own_prompts(self, fake_st: list[_FakeModel]) -> None:
        embedder = SentenceTransformerEmbedder("lightonai/mDenseOn", dims=4)
        embedder.embed_documents(["a region", "a column"])
        embedder.embed_query("revenue")
        (model,) = fake_st
        assert [(texts, prompt) for texts, prompt, _ in model.calls] == [
            (["a region", "a column"], "document"),
            (["revenue"], "query"),
        ]
        assert all(normalized for _, _, normalized in model.calls)

    def test_the_model_loads_once_and_lazily(self, fake_st: list[_FakeModel]) -> None:
        embedder = SentenceTransformerEmbedder("m", dims=4, max_seq_length=512, device="cpu")
        assert fake_st == []
        embedder.embed_query("a")
        embedder.embed_query("b")
        assert len(fake_st) == 1
        assert fake_st[0].max_seq_length == 512
        assert fake_st[0].device == "cpu"

    def test_a_dimension_mismatch_is_refused_on_load(self, fake_st: list[_FakeModel]) -> None:
        with pytest.raises(EmbeddingError, match="768"):
            SentenceTransformerEmbedder("m", dims=768).embed_query("a")

    def test_a_model_without_both_prompts_is_refused(
        self, fake_st: list[_FakeModel], monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(_FakeModel, "__init__", _no_prompts_init)
        with pytest.raises(EmbeddingError, match="prompt"):
            SentenceTransformerEmbedder("m", dims=4).embed_query("a")

    def test_no_texts_is_no_call(self, fake_st: list[_FakeModel]) -> None:
        assert SentenceTransformerEmbedder("m", dims=4).embed_documents([]) == []
        assert fake_st == []

    def test_a_missing_extra_says_how_to_fix_it(self, monkeypatch: pytest.MonkeyPatch) -> None:
        def missing(name: str) -> Any:
            raise ModuleNotFoundError(name)

        monkeypatch.setattr(embedding_module.importlib, "import_module", missing)
        with pytest.raises(EmbeddingError, match=r"excel-rag\[embed\]"):
            SentenceTransformerEmbedder("m", dims=4).embed_query("a")


def _no_prompts_init(self: _FakeModel, name: str, device: str | None = None) -> None:
    self.name, self.device, self.dims, self.prompts, self.calls = name, device, 4, {}, []


class TestHashingEmbedder:
    def test_is_deterministic_unit_length_and_never_zero(self) -> None:
        embedder = HashingEmbedder(32)
        first, empty = embedder.embed_documents(["Projected revenue", ""])
        assert first == embedder.embed_query("Projected revenue")
        assert math.isclose(math.sqrt(sum(v * v for v in first)), 1.0)
        assert any(empty), "cosine similarity refuses an all-zero vector"

    def test_shared_words_score_higher_than_none(self) -> None:
        embedder = HashingEmbedder(64)
        query = embedder.embed_query("revenue forecast")
        near, far = embedder.embed_documents(["Revenue forecast by year", "Headcount plan"])
        assert _dot(query, near) > _dot(query, far)

    def test_is_never_labelled_as_a_real_model(self) -> None:
        settings = EmbeddingSettings(provider="hashing", model="lightonai/mDenseOn", dims=16)
        embedder = build_embedder(settings)
        assert isinstance(embedder, HashingEmbedder)
        assert embedder.model_name != "lightonai/mDenseOn"


class TestBuild:
    def test_mdenseon_is_the_default(self) -> None:
        settings = EmbeddingSettings()
        assert (settings.provider, settings.model, settings.dims) == (
            "sentence-transformers",
            "lightonai/mDenseOn",
            768,
        )
        embedder = build_embedder(settings)
        assert isinstance(embedder, SentenceTransformerEmbedder)
        assert embedder.model_name == "lightonai/mDenseOn"

    def test_none_switches_vectors_off(self) -> None:
        assert build_embedder(EmbeddingSettings(provider="none")) is None


def test_a_chunk_is_embedded_as_title_then_content() -> None:
    assert chunk_text("Revenue", "by year") == "Revenue\nby year"
    assert chunk_text("", "by year") == "by year"


def _dot(left: tuple[float, ...], right: tuple[float, ...]) -> float:
    return sum(a * b for a, b in zip(left, right, strict=True))
