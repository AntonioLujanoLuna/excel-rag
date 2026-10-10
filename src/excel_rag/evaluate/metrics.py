"""Retrieval-quality metrics over evaluation cases: hit rate at k, range recall at k, and MRR.

* **hit@k** -- the fraction of cases with at least one relevant hit in the top ``k``.
* **recall@k** -- per case, the fraction of its expected rectangles that some relevant hit in the
  top ``k`` intersects, averaged over cases. A case that expects two places and finds one scores
  0.5 however high that one ranks.
* **MRR** -- the mean reciprocal rank of the first relevant hit (0 when none is found), over the
  ranked list the run returned.

Relevance is citation-shaped: a hit is relevant when it is from the case's workbook, on an expected
sheet, and its A1 range intersects an expected rectangle. Workbook and sheet summaries are never
relevant -- a summary intersects everything on its sheet and is not a citation.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass

from ..models import A1Range, Hit
from .cases import EvalCase

#: Node kinds whose chunks summarise rather than cite.
SUMMARY_KINDS = frozenset({"workbook", "sheet"})


def _node_kind(hit: Hit) -> str:
    prefix = f"{hit.source.workbook_id}:v{hit.source.version}:"
    remainder = hit.node_id.removeprefix(prefix)
    return remainder.split(":", 1)[0]


def _hit_range(hit: Hit) -> A1Range | None:
    if not hit.source.sheet or not hit.source.a1_range:
        return None
    try:
        return A1Range.parse(hit.source.sheet, hit.source.a1_range)
    except ValueError:
        return None


def matched_expectations(hit: Hit, case: EvalCase) -> frozenset[int]:
    """The indices of ``case.expected`` this hit cites (empty when it is not relevant)."""
    if hit.source.workbook_id != case.workbook_id or _node_kind(hit) in SUMMARY_KINDS:
        return frozenset()
    bounds = _hit_range(hit)
    if bounds is None:
        return frozenset()
    return frozenset(
        index
        for index, expected in enumerate(case.expected)
        if expected.sheet == bounds.sheet_name and expected.bounds().intersects(bounds)
    )


@dataclass(frozen=True, slots=True)
class CaseResult:
    """How one case went: the rank of its first relevant hit, and what each rank cited."""

    case: EvalCase
    first_rank: int | None
    #: For each returned hit, in rank order, the expected indices it cites.
    cited: tuple[frozenset[int], ...]
    top_hits: tuple[str, ...]

    def found_within(self, k: int) -> bool:
        return self.first_rank is not None and self.first_rank <= k

    def recall_at(self, k: int) -> float:
        covered: set[int] = set()
        for cited in self.cited[:k]:
            covered |= cited
        return len(covered) / len(self.case.expected)


def score_case(case: EvalCase, hits: Sequence[Hit]) -> CaseResult:
    cited = tuple(matched_expectations(hit, case) for hit in hits)
    first = next((rank for rank, found in enumerate(cited, start=1) if found), None)
    top = tuple(f"{hit.source.sheet}!{hit.source.a1_range}" for hit in hits[:3])
    return CaseResult(case=case, first_rank=first, cited=cited, top_hits=top)


@dataclass(frozen=True, slots=True)
class Metrics:
    """Aggregate metrics for one configuration over a set of cases."""

    cases: int
    hit_at: dict[int, float]
    recall_at: dict[int, float]
    mrr: float


def aggregate(results: Sequence[CaseResult], k_values: Sequence[int]) -> Metrics:
    count = len(results)
    if count == 0:
        return Metrics(0, dict.fromkeys(k_values, 0.0), dict.fromkeys(k_values, 0.0), 0.0)
    return Metrics(
        cases=count,
        hit_at={k: sum(result.found_within(k) for result in results) / count for k in k_values},
        recall_at={k: sum(result.recall_at(k) for result in results) / count for k in k_values},
        mrr=sum(1.0 / result.first_rank for result in results if result.first_rank) / count,
    )


__all__ = [
    "SUMMARY_KINDS",
    "CaseResult",
    "Metrics",
    "aggregate",
    "matched_expectations",
    "score_case",
]
