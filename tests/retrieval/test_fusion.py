"""Fusion tests: reciprocal-rank fusion of the lexical and knn lists, and hit provenance."""

from __future__ import annotations

import pytest
from conftest import EMBED

from excel_rag.models import ChunkDocument, ChunkType
from excel_rag.retrieval.fusion import RRF_K, fuse, matched_fields


def _chunk(
    identifier: str,
    *,
    embedding: tuple[float, ...] | None = None,
    model: str | None = EMBED,
    title: str = "t",
    content: str = "c",
    headers: tuple[str, ...] = (),
) -> ChunkDocument:
    return ChunkDocument(
        id=identifier,
        workbook_id="wb",
        version=1,
        node_id="n",
        sheet_id="s",
        sheet_name="Sheet1",
        a1_range="A1",
        chunk_type=ChunkType.TABLE,
        title=title,
        content=content,
        headers=headers,
        embedding=embedding,
        embedding_model=model,
    )


class TestFuse:
    def test_lexical_only_preserves_order_and_the_index_score(self) -> None:
        ranked = fuse([(_chunk("b"), 1.0), (_chunk("a"), 2.0)])
        assert [r.candidate.chunk.id for r in ranked] == ["a", "b"]
        assert ranked[0].score == 2.0
        assert not any(r.vector_matched for r in ranked)

    def test_ties_break_on_id(self) -> None:
        ranked = fuse([(_chunk("b"), 1.0), (_chunk("a"), 1.0)])
        assert [r.candidate.chunk.id for r in ranked] == ["a", "b"]

    def test_a_chunk_both_retrievers_found_outranks_one_found_once(self) -> None:
        lexical = [(_chunk("a"), 3.0), (_chunk("b"), 2.0)]
        vector = [(_chunk("b"), 0.9), (_chunk("c"), 0.8)]
        ranked = fuse(lexical, vector)
        assert [r.candidate.chunk.id for r in ranked] == ["b", "a", "c"]
        assert ranked[0].score == pytest.approx(1 / (RRF_K + 2) + 1 / (RRF_K + 1))
        assert [r.vector_matched for r in ranked] == [True, False, True]

    def test_a_vector_only_hit_is_kept_with_its_scores(self) -> None:
        (only,) = fuse([], [(_chunk("v"), 0.7)])
        assert only.candidate.lexical_score is None
        assert only.candidate.vector_score == 0.7
        assert only.score == pytest.approx(1 / (RRF_K + 1))

    def test_an_empty_vector_list_still_fuses(self) -> None:
        """A vector query that found nothing is fused, not skipped: scores are on the RRF scale."""
        (only,) = fuse([(_chunk("a"), 5.0)], [])
        assert only.score == pytest.approx(1 / (RRF_K + 1))
        assert not only.vector_matched

    def test_a_repeated_chunk_keeps_its_best_rank(self) -> None:
        ranked = fuse([(_chunk("a"), 2.0), (_chunk("a"), 1.0)], [])
        assert len(ranked) == 1
        assert ranked[0].score == pytest.approx(1 / (RRF_K + 1))

    def test_nothing_fuses_to_nothing(self) -> None:
        assert fuse([]) == []
        assert fuse([], []) == []


class TestMatchedFields:
    def test_reports_where_the_tokens_land(self) -> None:
        chunk = _chunk("a", title="Projected revenue", content="by year", headers=("Metric",))
        assert matched_fields("revenue", chunk) == ("title",)
        assert matched_fields("metric", chunk) == ("headers",)
        assert matched_fields("nothing", chunk) == ()

    def test_blank_query_matches_nothing(self) -> None:
        assert matched_fields("   ", _chunk("a")) == ()

    def test_a_vector_match_is_reported_as_embedding(self) -> None:
        chunk = _chunk("a", title="Projected revenue")
        assert matched_fields("income", chunk, vector_matched=True) == ("embedding",)
        assert matched_fields("revenue", chunk, vector_matched=True) == ("title", "embedding")


class TestMatchedFieldsTokens:
    def test_punctuation_splits_tokens_as_the_standard_analyzer_does(self) -> None:
        chunk = _chunk("a", title="x", content="Revenue/cost per Q1-2026 (USD)")
        assert matched_fields("cost", chunk) == ("content",)
        assert matched_fields("2026", chunk) == ("content",)
        assert matched_fields("usd", chunk) == ("content",)

    def test_non_ascii_words_are_whole_tokens(self) -> None:
        chunk = _chunk("a", title="Previsión de ingresos", content="c")
        assert matched_fields("previsión", chunk) == ("title",)
        assert matched_fields("prevision", chunk) == ()
