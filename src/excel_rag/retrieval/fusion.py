"""Hybrid scoring over the candidate window the index returned.

The in-memory client does not implement ``knn`` and deliberately refuses it, so there is no
exhaustive vector search here: the lexical query returns a bounded candidate window and the vectors
are scored in Python over exactly that window. That is a **candidate-window re-rank**, not an
approximate-nearest-neighbour search, and the docs say so. The live adapter builds a real ``knn``
query (see :mod:`excel_rag.live`); the same fusion code would then re-rank a candidate window that
Elasticsearch had already narrowed.

**A vector from a different model is a miss, not a silent re-ranking.** Each chunk carries the
``embedding_model`` that produced its vector. If the caller's query vector was produced by a
different model, the chunk is excluded from the vector ranking entirely -- it keeps its lexical
rank and earns no vector contribution -- rather than being scored against an incomparable vector.
The same is true of a dimensionality mismatch: it is a miss, not a zero score that quietly drags
the chunk down.

Ranking is **reciprocal-rank fusion** (RRF), chosen over a weighted sum because the lexical score
this service receives is ordinal: the in-memory client returns counts of matched leaf clauses, not
BM25, and the live client returns BM25 -- two scales that a fixed weight cannot reconcile. RRF only
needs the *order* each retriever produced, so it behaves the same whichever client is underneath.
"""

from __future__ import annotations

import math
from collections.abc import Iterable, Sequence
from dataclasses import dataclass

from ..models import ChunkDocument

#: The RRF smoothing constant. 60 is the value from the original Cormack et al. paper; it keeps any
#: single retriever's top hit from dominating and is stable across candidate-window sizes.
RRF_K = 60


@dataclass(frozen=True)
class Candidate:
    """One chunk in the candidate window, with the lexical score the index gave it."""

    chunk: ChunkDocument
    lexical_score: float


@dataclass(frozen=True)
class RankedCandidate:
    """A candidate after fusion, with the score to report and whether a vector contributed."""

    candidate: Candidate
    score: float
    vector_matched: bool


def cosine_similarity(
    query_vector: Sequence[float],
    embedding: Sequence[float] | None,
    embedding_model: str | None,
    query_model: str | None,
) -> float | None:
    """Cosine similarity, or ``None`` when the vectors are not comparable.

    ``None`` (a miss, not a zero) is returned when the chunk has no embedding, when the model that
    produced the stored vector does not equal ``query_model``, when ``query_model`` is unknown, or
    when the dimensionalities disagree.
    """
    if embedding is None or query_model is None:
        return None
    if embedding_model != query_model:
        return None
    if len(embedding) != len(query_vector) or not embedding:
        return None
    dot = 0.0
    norm_a = 0.0
    norm_b = 0.0
    for left, right in zip(query_vector, embedding, strict=True):
        dot += left * right
        norm_a += left * left
        norm_b += right * right
    if norm_a == 0.0 or norm_b == 0.0:
        return 0.0
    return dot / (math.sqrt(norm_a) * math.sqrt(norm_b))


def rank(
    candidates: Sequence[Candidate],
    *,
    query_vector: Sequence[float] | None = None,
    query_model: str | None = None,
    rrf_k: int = RRF_K,
) -> list[RankedCandidate]:
    """Rank a candidate window by RRF of the lexical order and the (optional) vector order.

    With no ``query_vector`` the result is the lexical order with the index's own scores; that is
    the HTTP path, whose frozen ``SearchRequest`` carries no vector. With one, each candidate earns
    a contribution from its lexical rank and, where its stored vector is comparable, from its
    vector rank.
    """
    lexical_order = sorted(
        candidates, key=lambda candidate: (-candidate.lexical_score, _key(candidate))
    )
    lexical_rank = {
        _key(candidate): position for position, candidate in enumerate(lexical_order, 1)
    }

    if query_vector is None:
        return [
            RankedCandidate(
                candidate=candidate, score=candidate.lexical_score, vector_matched=False
            )
            for candidate in lexical_order
        ]

    vector_scores: dict[str, float] = {}
    for candidate in candidates:
        similarity = cosine_similarity(
            query_vector,
            candidate.chunk.embedding,
            candidate.chunk.embedding_model,
            query_model,
        )
        if similarity is not None:
            vector_scores[_key(candidate)] = similarity
    vector_order = sorted(vector_scores, key=lambda key: (-vector_scores[key], key))
    vector_rank = {key: position for position, key in enumerate(vector_order, 1)}

    fused: list[RankedCandidate] = []
    for candidate in candidates:
        key = _key(candidate)
        score = 1.0 / (rrf_k + lexical_rank[key])
        matched = key in vector_rank
        if matched:
            score += 1.0 / (rrf_k + vector_rank[key])
        fused.append(RankedCandidate(candidate=candidate, score=score, vector_matched=matched))
    fused.sort(key=lambda ranked: (-ranked.score, _key(ranked.candidate)))
    return fused


def matched_fields(query: str, chunk: ChunkDocument) -> tuple[str, ...]:
    """Which searchable fields of a chunk a query's tokens appear in, for the hit's provenance.

    It is derived from the document the index returned, not from the index's own term statistics,
    so it describes where the provenance is, not Elasticsearch's scoring internals.
    """
    query_tokens = _tokens(query)
    if not query_tokens:
        return ()
    fields: list[str] = []
    if query_tokens & _tokens(chunk.title):
        fields.append("title")
    if query_tokens & _tokens(chunk.content):
        fields.append("content")
    if query_tokens & _tokens(" ".join(chunk.headers)):
        fields.append("headers")
    return tuple(fields)


def _key(candidate: Candidate) -> str:
    return candidate.chunk.id


def _tokens(text: str) -> set[str]:
    return {token.strip(".,;:()?!\"'") for token in text.lower().split() if token.strip()}


def iter_vector_models(candidates: Iterable[Candidate]) -> set[str]:
    """The distinct embedding models present in a candidate window (diagnostics/tests)."""
    return {c.chunk.embedding_model for c in candidates if c.chunk.embedding_model is not None}


__all__ = [
    "RRF_K",
    "Candidate",
    "RankedCandidate",
    "cosine_similarity",
    "iter_vector_models",
    "matched_fields",
    "rank",
]
