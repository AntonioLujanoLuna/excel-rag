"""``python -m excel_rag.evaluate``: compare lexical, hybrid and reranked retrieval on evidence."""

from __future__ import annotations

import argparse
import json
import sys
import tempfile
from collections.abc import Callable
from dataclasses import asdict
from pathlib import Path

from ..embedding import Embedder, EmbeddingError, HashingEmbedder, build_embedder
from ..rerank import OverlapReranker, Reranker, build_reranker
from ..settings import Settings
from ..workbook import load_workbook
from ..workbook.errors import WorkbookError
from .cases import CaseFileError, EvalCase, load_cases
from .mining import mine_cases
from .runner import (
    DEFAULT_K,
    Configuration,
    Threshold,
    check_thresholds,
    evaluate,
    format_report,
)
from .sample import SAMPLE_CASES, SAMPLE_WORKBOOK_ID, build_sample_workbook


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="python -m excel_rag.evaluate", description=__doc__)
    parser.add_argument("--cases", help="JSON Lines case file (default: the built-in sample)")
    parser.add_argument(
        "--workbook",
        action="append",
        default=[],
        metavar="ID=PATH",
        help="a workbook the cases name, repeatable (required with --cases)",
    )
    parser.add_argument(
        "--embedder",
        choices=("settings", "hashing", "none"),
        default="settings",
        help="the hybrid run's embedder: the configured one (EXCEL_RAG_EMBEDDING__*), the hashing "
        "test double, or none (lexical only)",
    )
    parser.add_argument(
        "--rerank",
        choices=("settings", "overlap", "none"),
        default="settings",
        help="the reranked run's reranker: the configured one (EXCEL_RAG_RERANK__*), the overlap "
        "test double, or none",
    )
    parser.add_argument("-k", type=int, action="append", help=f"cut-offs (default {DEFAULT_K})")
    parser.add_argument("--json", action="store_true", help="print the report as JSON")
    parser.add_argument("--no-misses", action="store_true", help="omit the per-case misses")
    parser.add_argument(
        "--mine",
        action="store_true",
        help="ask questions mined from the workbooks' own labels (labelled formulas and inputs) "
        "instead of --cases; with no --workbook, from the built-in sample workbook",
    )
    parser.add_argument("--save-cases", metavar="PATH", help="write the cases run as JSON Lines")
    parser.add_argument(
        "--min",
        action="append",
        default=[],
        metavar="[CONFIG:]METRIC=VALUE",
        help="fail (exit 3) when a run's metric is below VALUE, e.g. mrr=0.5 or hybrid:hit@5=0.7; "
        "repeatable",
    )
    return parser


def _embedder(choice: str, settings: Settings) -> Embedder | None:
    if choice == "none":
        return None
    if choice == "hashing":
        return HashingEmbedder(settings.embedding.dims)
    return build_embedder(settings.embedding)


def _reranker(choice: str, settings: Settings) -> Reranker | None:
    if choice == "none":
        return None
    if choice == "overlap":
        return OverlapReranker()
    return build_reranker(settings.rerank)


def _usable(probe: Callable[[], object] | None, label: str) -> bool:
    """Load a model once up front, so a missing extra skips a run instead of failing mid-way."""
    if probe is None:
        return False
    try:
        probe()
    except EmbeddingError as error:
        print(f"excel-rag evaluate: skipping the {label} run: {error}", file=sys.stderr)
        return False
    return True


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    settings = Settings()
    with tempfile.TemporaryDirectory() as scratch:
        try:
            thresholds = [Threshold.parse(text) for text in args.min]
            workbooks: dict[str, str | Path] = {}
            for item in args.workbook:
                workbook_id, sep, path = item.partition("=")
                if not sep or not workbook_id or not path:
                    print(f"excel-rag evaluate: --workbook wants ID=PATH, got {item!r}")
                    return 2
                workbooks[workbook_id] = path
            if args.cases and args.mine:
                print("excel-rag evaluate: --cases and --mine are alternatives", file=sys.stderr)
                return 2
            if not workbooks and not args.cases:
                workbooks = {SAMPLE_WORKBOOK_ID: build_sample_workbook(Path(scratch))}
            if args.cases:
                cases = load_cases(args.cases)
            elif args.mine:
                cases = tuple(
                    case
                    for workbook_id, path in workbooks.items()
                    for case in mine_cases(load_workbook(path, workbook_id=workbook_id))
                )
                if not cases:
                    print(
                        "excel-rag evaluate: no labelled formulas or inputs to mine",
                        file=sys.stderr,
                    )
                    return 1
            else:
                cases = SAMPLE_CASES
            if args.save_cases:
                Path(args.save_cases).write_text(
                    "".join(
                        json.dumps(_case_json(case), ensure_ascii=False) + "\n" for case in cases
                    ),
                    encoding="utf-8",
                )
        except (CaseFileError, OSError, ValueError, WorkbookError) as error:
            print(f"excel-rag evaluate: {error}", file=sys.stderr)
            return 1

        embedder = _embedder(args.embedder, settings)
        reranker = _reranker(args.rerank, settings)
        configurations = [Configuration("lexical")]
        hybrid = _usable(
            (lambda: embedder.embed_query("probe")) if embedder is not None else None, "hybrid"
        )
        if hybrid:
            configurations.append(Configuration("hybrid", embedder=embedder))
        if _usable(
            (lambda: reranker.score("probe", ["probe"])) if reranker is not None else None,
            "reranked",
        ):
            configurations.append(
                Configuration(
                    "hybrid+rerank" if hybrid else "lexical+rerank",
                    embedder=embedder if hybrid else None,
                    reranker=reranker,
                )
            )
        try:
            report = evaluate(
                configurations, cases, workbooks, settings=settings, k_values=args.k or DEFAULT_K
            )
        except ValueError as error:
            print(f"excel-rag evaluate: {error}", file=sys.stderr)
            return 1

    if args.json:
        payload = asdict(report)
        for run in payload["runs"]:
            for result in run["results"]:
                result["cited"] = [sorted(cited) for cited in result["cited"]]
        print(json.dumps(payload, indent=2, sort_keys=True, default=str))
    else:
        print(format_report(report, show_misses=not args.no_misses))
    try:
        failures = check_thresholds(report, thresholds)
    except ValueError as error:
        print(f"excel-rag evaluate: {error}", file=sys.stderr)
        return 1
    for failure in failures:
        print(f"excel-rag evaluate: {failure}", file=sys.stderr)
    return 3 if failures else 0


def _case_json(case: EvalCase) -> dict[str, object]:
    return {
        "id": case.id,
        "question": case.question,
        "workbook_id": case.workbook_id,
        "expected": [{"sheet": item.sheet, "a1": item.a1} for item in case.expected],
    }


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
