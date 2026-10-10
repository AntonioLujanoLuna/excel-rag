"""Retrieval-quality evaluation: does the right rectangle come back, and how high?

``python -m excel_rag.evaluate`` indexes a workbook corpus into the in-memory double, asks each
case's question, and reports hit@k, recall@k and MRR for each configuration it can run -- lexical,
hybrid with the configured embedder, and hybrid plus the configured reranker. With no arguments it
uses a built-in planning workbook and sixteen questions (:mod:`.sample`); ``--cases`` and
``--workbook`` run your own. This is the harness for tuning chunking, the RRF constant, the
embedding model or a reranker on evidence rather than on intuition.
"""

from __future__ import annotations

from .cases import CaseFileError, EvalCase, Expected, load_cases, parse_cases
from .metrics import CaseResult, Metrics, aggregate, matched_expectations, score_case
from .runner import Configuration, EvalReport, RunResult, evaluate, format_report

__all__ = [
    "CaseFileError",
    "CaseResult",
    "Configuration",
    "EvalCase",
    "EvalReport",
    "Expected",
    "Metrics",
    "RunResult",
    "aggregate",
    "evaluate",
    "format_report",
    "load_cases",
    "matched_expectations",
    "parse_cases",
    "score_case",
]
