"""Hybrid ranking: reciprocal-rank fusion of the lexical and the dense-vector result lists.

Two retrievers run against ``excel_chunks``: a BM25 ``bool`` query over title/content/headers, and
a ``knn`` query over the chunk embeddings (``lightonai/mDenseOn`` by default). Each returns its own
ranked list from the index, under the same ACL, version and facet filters. This module merges them.

**Reciprocal-rank fusion** (RRF) is used rather than a weighted sum of scores, because the two
scales are incomparable: BM25 is unbounded and corpus-dependent, a cosine ``knn`` score is
``(1 + cos) / 2``, and the in-memory double returns ordinal match counts. RRF needs only each
list's *order*, so it behaves the same whichever client is underneath. A chunk found by both
retrievers earns two contributions; a chunk only one retriever found still ranks.

**A vector from a different model is a miss, not a silent re-ranking.** The ``knn`` query filters
on ``embedding_model``, so a chunk embedded by another model never enters the vector list: it keeps
its lexical rank and earns no vector contribution. Comparing vectors across models would rank
by an artefact of two unrelated spaces.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass

from ..models import ChunkDocument

#: The RRF smoothing constant. 60 is the value from the original Cormack et al. paper; it keeps any
#: single retriever's top hit from dominating and is stable across candidate-window sizes.
RRF_K = 60

#: A scored hit as a repository returns it: the chunk and the index's own score.
Scored = tuple[ChunkDocument, float]


@dataclass(frozen=True)
class Candidate:
    """A chunk in the fused window and each retriever's score for it (``None``: not found)."""

    chunk: ChunkDocument
    lexical_score: float | None = None
    vector_score: float | None = None


@dataclass(frozen=True)
class RankedCandidate:
    """A candidate after fusion, with the score to report and whether a vector contributed."""

    candidate: Candidate
    score: float
    vector_matched: bool


def fuse(
    lexical: Sequence[Scored],
    vector: Sequence[Scored] | None = None,
    *,
    rrf_k: int = RRF_K,
) -> list[RankedCandidate]:
    """Merge the two ranked lists.

    With ``vector`` ``None`` (no embedder, or no query vector) the result is the lexical order with
    the index's own scores -- nothing is fused, so nothing is re-scaled. With a vector list, even an
    empty one, every chunk is scored ``sum(1 / (rrf_k + rank))`` over the lists it appears in, and
    the reported score is that RRF score.
    """
    lexical_order = sorted(lexical, key=lambda hit: (-hit[1], hit[0].id))
    if vector is None:
        return [
            RankedCandidate(
                candidate=Candidate(chunk=chunk, lexical_score=score),
                score=score,
                vector_matched=False,
            )
            for chunk, score in lexical_order
        ]

    vector_order = sorted(vector, key=lambda hit: (-hit[1], hit[0].id))
    chunks: dict[str, ChunkDocument] = {}
    lexical_scores: dict[str, float] = {}
    vector_scores: dict[str, float] = {}
    fused: dict[str, float] = {}
    for scores, order in ((lexical_scores, lexical_order), (vector_scores, vector_order)):
        for position, (chunk, score) in enumerate(order, 1):
            if chunk.id in scores:
                continue  # a list repeating a chunk keeps its best rank only
            chunks.setdefault(chunk.id, chunk)
            scores[chunk.id] = score
            fused[chunk.id] = fused.get(chunk.id, 0.0) + 1.0 / (rrf_k + position)

    ranked = [
        RankedCandidate(
            candidate=Candidate(
                chunk=chunks[chunk_id],
                lexical_score=lexical_scores.get(chunk_id),
                vector_score=vector_scores.get(chunk_id),
            ),
            score=score,
            vector_matched=chunk_id in vector_scores,
        )
        for chunk_id, score in fused.items()
    ]
    ranked.sort(key=lambda item: (-item.score, item.candidate.chunk.id))
    return ranked


def matched_fields(
    query: str, chunk: ChunkDocument, *, vector_matched: bool = False
) -> tuple[str, ...]:
    """Where a hit's provenance is: the fields the query's tokens appear in, and ``embedding``
    when the vector retriever found it.

    It is derived from the document the index returned, not from the index's own term statistics,
    so it describes where the provenance is, not Elasticsearch's scoring internals.
    """
    fields: list[str] = []
    query_tokens = _tokens(query)
    if query_tokens:
        if query_tokens & _tokens(chunk.title):
            fields.append("title")
        if query_tokens & _tokens(chunk.content):
            fields.append("content")
        if query_tokens & _tokens(" ".join(chunk.headers)):
            fields.append("headers")
    if vector_matched:
        fields.append("embedding")
    return tuple(fields)


def _tokens(text: str) -> set[str]:
    return {token.strip(".,;:()?!\"'") for token in text.lower().split() if token.strip()}


__all__ = [
    "RRF_K",
    "Candidate",
    "RankedCandidate",
    "Scored",
    "fuse",
    "matched_fields",
]
