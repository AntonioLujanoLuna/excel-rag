"""Run evaluation cases against the real retrieval stack over the in-memory index.

Each workbook is ingested by the real parser and indexed (with the configuration's embedder) into
the in-memory Elasticsearch double; each case is then one :meth:`RetrievalService.search`. That
measures what ingestion, chunking, embeddings, fusion and reranking do to *ranking*. It does not
measure a cluster's approximate ``knn`` (the double's is exact) or its BM25 scoring (the double's
lexical score is a match count), so figures from a live cluster can differ, and the report says so.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from pathlib import Path

from ..embedding import Embedder
from ..fake_es import in_memory_client
from ..ingest import ingest_workbook
from ..ingest.indexer import Indexer
from ..models import SearchFilters, SearchRequest
from ..rerank import Reranker
from ..retrieval import Repository, RetrievalService
from ..settings import Settings
from .cases import EvalCase
from .metrics import CaseResult, Metrics, aggregate, score_case

DEFAULT_K = (1, 3, 5, 10)
NOTE = (
    "ranking measured on the IN-MEMORY index (exact knn, match-count lexical scores); "
    "a live cluster's BM25 and approximate knn can rank differently"
)


@dataclass(frozen=True)
class Configuration:
    """One way of answering: lexical, hybrid, hybrid + rerank, or anything else to compare."""

    name: str
    embedder: Embedder | None = None
    reranker: Reranker | None = None


@dataclass(frozen=True)
class RunResult:
    configuration: str
    metrics: Metrics
    results: tuple[CaseResult, ...]
    embedding_model: str | None = None
    reranker_model: str | None = None


@dataclass(frozen=True)
class EvalReport:
    runs: tuple[RunResult, ...]
    k_values: tuple[int, ...]
    note: str = NOTE
    workbooks: dict[str, str] = field(default_factory=dict)


def run_configuration(
    configuration: Configuration,
    cases: Sequence[EvalCase],
    workbooks: Mapping[str, str | Path],
    *,
    settings: Settings | None = None,
    k_values: Sequence[int] = DEFAULT_K,
) -> RunResult:
    """Index ``workbooks`` (workbook id -> path) and answer every case under one configuration."""
    resolved = settings or Settings(embedding={"provider": "none"})
    if configuration.embedder is not None:
        resolved = resolved.model_copy(
            update={
                "embedding": resolved.embedding.model_copy(
                    update={"dims": configuration.embedder.dims}
                )
            }
        )
    client = in_memory_client(resolved)
    indexer = Indexer(client, resolved, embedder=configuration.embedder)
    for workbook_id, path in workbooks.items():
        indexer.index_workbook(ingest_workbook(path, workbook_id=workbook_id, version=1))
    service = RetrievalService(
        Repository(client, resolved),
        resolved,
        configuration.embedder,
        configuration.reranker,
    )
    top_k = max(k_values)
    results: list[CaseResult] = []
    for case in cases:
        if case.workbook_id not in workbooks:
            raise ValueError(f"case {case.id!r} names workbook {case.workbook_id!r}, not given")
        response = service.search(
            SearchRequest(
                query=case.question,
                filters=SearchFilters(workbook_ids=(case.workbook_id,)),
                top_k=top_k,
                include_structure=False,
            )
        )
        results.append(score_case(case, response.hits))
    return RunResult(
        configuration=configuration.name,
        metrics=aggregate(results, k_values),
        results=tuple(results),
        embedding_model=configuration.embedder.model_name if configuration.embedder else None,
        reranker_model=configuration.reranker.model_name if configuration.reranker else None,
    )


def evaluate(
    configurations: Sequence[Configuration],
    cases: Sequence[EvalCase],
    workbooks: Mapping[str, str | Path],
    *,
    settings: Settings | None = None,
    k_values: Sequence[int] = DEFAULT_K,
) -> EvalReport:
    """Run every configuration over the same cases and workbooks."""
    ks = tuple(sorted(set(k_values)))
    return EvalReport(
        runs=tuple(
            run_configuration(configuration, cases, workbooks, settings=settings, k_values=ks)
            for configuration in configurations
        ),
        k_values=ks,
        workbooks={key: str(value) for key, value in workbooks.items()},
    )


def format_report(report: EvalReport, *, show_misses: bool = True) -> str:
    """A plain-text table, then (optionally) each configuration's cases with no relevant hit."""
    ks = report.k_values
    header = ["configuration", "cases", *(f"hit@{k}" for k in ks), f"recall@{ks[-1]}", "MRR"]
    rows = [
        [
            run.configuration,
            str(run.metrics.cases),
            *(f"{run.metrics.hit_at[k]:.2f}" for k in ks),
            f"{run.metrics.recall_at[ks[-1]]:.2f}",
            f"{run.metrics.mrr:.3f}",
        ]
        for run in report.runs
    ]
    widths = [max(len(row[index]) for row in [header, *rows]) for index in range(len(header))]
    lines = [
        "  ".join(cell.ljust(width) for cell, width in zip(row, widths, strict=True))
        for row in [header, *rows]
    ]
    lines.insert(1, "  ".join("-" * width for width in widths))
    for run in report.runs:
        model = ", ".join(
            part
            for part in (
                f"embedder {run.embedding_model}" if run.embedding_model else "",
                f"reranker {run.reranker_model}" if run.reranker_model else "",
            )
            if part
        )
        if model:
            lines.append(f"{run.configuration}: {model}")
    if show_misses:
        for run in report.runs:
            misses = [result for result in run.results if not result.found_within(ks[-1])]
            if misses:
                lines.append(f"\n{run.configuration}: no relevant hit in the top {ks[-1]} for")
                lines.extend(
                    f"  - {result.case.id}: {result.case.question!r} "
                    f"(top: {', '.join(result.top_hits) or 'nothing'})"
                    for result in misses
                )
    lines.append(f"\nNote: {report.note}.")
    return "\n".join(lines)


__all__ = [
    "DEFAULT_K",
    "Configuration",
    "EvalReport",
    "RunResult",
    "evaluate",
    "format_report",
    "run_configuration",
]
