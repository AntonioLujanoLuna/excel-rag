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
from .cases import CaseFileError, load_cases
from .runner import DEFAULT_K, Configuration, evaluate, format_report
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
            if args.cases:
                cases = load_cases(args.cases)
                workbooks: dict[str, str | Path] = {}
                for item in args.workbook:
                    workbook_id, sep, path = item.partition("=")
                    if not sep or not workbook_id or not path:
                        print(f"excel-rag evaluate: --workbook wants ID=PATH, got {item!r}")
                        return 2
                    workbooks[workbook_id] = path
            else:
                cases = SAMPLE_CASES
                workbooks = {SAMPLE_WORKBOOK_ID: build_sample_workbook(Path(scratch))}
        except (CaseFileError, OSError) as error:
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
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
