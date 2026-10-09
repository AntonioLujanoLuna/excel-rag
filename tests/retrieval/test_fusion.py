"""Fusion tests: candidate-window re-rank, and the model-mismatch-is-a-miss rule."""

from __future__ import annotations

import pytest
from conftest import EMBED

from excel_rag.models import ChunkDocument, ChunkType
from excel_rag.retrieval.fusion import Candidate, cosine_similarity, matched_fields, rank


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


class TestCosineSimilarity:
    def test_identical_vectors_score_one(self) -> None:
        assert cosine_similarity((1.0, 0.0), (1.0, 0.0), EMBED, EMBED) == pytest.approx(1.0)

    def test_orthogonal_vectors_score_zero(self) -> None:
        assert cosine_similarity((1.0, 0.0), (0.0, 1.0), EMBED, EMBED) == pytest.approx(0.0)

    def test_zero_vector_is_zero_not_a_crash(self) -> None:
        assert cosine_similarity((0.0, 0.0), (1.0, 0.0), EMBED, EMBED) == 0.0

    def test_different_model_is_a_miss(self) -> None:
        assert cosine_similarity((1.0, 0.0), (1.0, 0.0), "other-model", EMBED) is None

    def test_unknown_query_model_is_a_miss(self) -> None:
        assert cosine_similarity((1.0, 0.0), (1.0, 0.0), EMBED, None) is None

    def test_missing_embedding_is_a_miss(self) -> None:
        assert cosine_similarity((1.0, 0.0), None, EMBED, EMBED) is None

    def test_dimensionality_mismatch_is_a_miss(self) -> None:
        assert cosine_similarity((1.0, 0.0), (1.0, 0.0, 0.0), EMBED, EMBED) is None


class TestRank:
    def test_lexical_only_preserves_order_and_score(self) -> None:
        candidates = [
            Candidate(chunk=_chunk("b"), lexical_score=1.0),
            Candidate(chunk=_chunk("a"), lexical_score=2.0),
        ]
        ranked = rank(candidates)
        assert [r.candidate.chunk.id for r in ranked] == ["a", "b"]
        assert ranked[0].score == 2.0
        assert not any(r.vector_matched for r in ranked)

    def test_ties_break_on_id(self) -> None:
        candidates = [
            Candidate(chunk=_chunk("b"), lexical_score=1.0),
            Candidate(chunk=_chunk("a"), lexical_score=1.0),
        ]
        assert [r.candidate.chunk.id for r in rank(candidates)] == ["a", "b"]

    def test_a_vector_promotes_a_lower_lexical_hit(self) -> None:
        candidates = [
            Candidate(chunk=_chunk("a", embedding=(1.0, 0.0), model="foreign"), lexical_score=2.0),
            Candidate(chunk=_chunk("b", embedding=(1.0, 0.0)), lexical_score=1.0),
        ]
        ranked = rank(candidates, query_vector=(1.0, 0.0), query_model=EMBED)
        # `a` wins lexically but its vector is from another model, so it earns no vector rank.
        assert ranked[0].candidate.chunk.id == "b"
        assert ranked[0].vector_matched is True
        assert ranked[1].vector_matched is False

    def test_a_foreign_model_vector_never_re_ranks(self) -> None:
        candidates = [
            Candidate(chunk=_chunk("a"), lexical_score=1.0),
            Candidate(chunk=_chunk("b", embedding=(1.0, 0.0), model="foreign"), lexical_score=1.0),
        ]
        ranked = rank(candidates, query_vector=(1.0, 0.0), query_model=EMBED)
        assert [r.candidate.chunk.id for r in ranked] == ["a", "b"]
        assert not any(r.vector_matched for r in ranked)

    def test_empty_window_ranks_to_nothing(self) -> None:
        assert rank([]) == []


class TestMatchedFields:
    def test_reports_where_the_tokens_land(self) -> None:
        chunk = _chunk("a", title="Projected revenue", content="by year", headers=("Metric",))
        assert matched_fields("revenue", chunk) == ("title",)
        assert matched_fields("metric", chunk) == ("headers",)
        assert matched_fields("nothing", chunk) == ()

    def test_blank_query_matches_nothing(self) -> None:
        assert matched_fields("   ", _chunk("a")) == ()
