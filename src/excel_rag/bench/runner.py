"""Measure the retrieval endpoints and expansion behaviour, and report the numbers honestly.

Latency is wall-clock around an in-process :class:`~fastapi.testclient.TestClient` call, so it
includes the FastAPI/pydantic overhead and a Python dict as the "index". It is not an Elasticsearch
latency and is not a production latency. Elasticsearch call counts come from the response envelope
and are the one number here that would carry over to a cluster.
"""

from __future__ import annotations

import math
import time
from dataclasses import dataclass
from typing import Any

from fastapi.testclient import TestClient

from .. import __version__
from ..app import create_app
from ..es import ElasticsearchLike
from ..settings import Settings
from .corpus import Corpus, build_corpus

DEFAULT_ITERATIONS = 30
NOTE = (
    "latency measured on the IN-MEMORY client (a Python dict), in-process via TestClient; "
    "NOT an Elasticsearch latency and NOT a production latency"
)


@dataclass(frozen=True)
class Latency:
    name: str
    iterations: int
    p50_ms: float
    p95_ms: float
    p99_ms: float
    mean_ms: float
    es_requests: int


@dataclass(frozen=True)
class ExpansionRow:
    depth: int
    nodes_returned: int
    truncated: bool
    p50_ms: float
    p95_ms: float
    mean_nodes: float
    es_requests: int


@dataclass(frozen=True)
class BenchResult:
    version: str
    chunk_count: int
    structure_count: int
    embedding_dims: int
    endpoints: tuple[Latency, ...]
    expansions: tuple[ExpansionRow, ...]
    note: str = NOTE


def prepare(settings: Settings, corpus: Corpus) -> TestClient:
    """Build the app, index the corpus into its client, and return a request client."""
    app = create_app(settings)
    client: ElasticsearchLike = app.state.client
    elastic = settings.elasticsearch
    client.bulk_index(elastic.chunks_index, corpus.chunks, refresh=True)
    client.bulk_index(elastic.structure_index, corpus.structure, refresh=True)
    client.bulk_index(elastic.versions_index, corpus.versions, refresh=True)
    return TestClient(app)


def percentile(samples: list[float], percent: float) -> float:
    """Nearest-rank percentile: the smallest value at or above ``percent`` of the samples."""
    if not samples:
        return 0.0
    ordered = sorted(samples)
    rank = math.ceil(percent / 100.0 * len(ordered))
    index = min(max(rank - 1, 0), len(ordered) - 1)
    return ordered[index]


def _time_call(call: Any) -> tuple[float, dict[str, Any]]:
    started = time.perf_counter()
    response = call()
    elapsed = (time.perf_counter() - started) * 1000.0
    if response.status_code != 200:  # pragma: no cover - a broken bench should fail loudly
        raise RuntimeError(f"benchmark request failed: {response.status_code} {response.text}")
    return elapsed, response.json()


def run_benchmark(
    *,
    iterations: int = DEFAULT_ITERATIONS,
    workbooks: int = 3,
    sheets: int = 4,
    cells_per_sheet: int = 2500,
    chunks_per_sheet: int = 180,
    embed_dims: int = 64,
    seed: int = 20261009,
) -> BenchResult:
    """Build a corpus, index it, and measure every endpoint plus bounded expansion."""
    # The bench raises the traversal budget so depth 0/1/2 are all actually reachable; the service
    # still caps every request at whatever the settings allow.
    settings = Settings(budgets={"reference_depth": 3, "max_related_nodes": 200})
    corpus = build_corpus(
        workbooks=workbooks,
        sheets=sheets,
        cells_per_sheet=cells_per_sheet,
        chunks_per_sheet=chunks_per_sheet,
        embed_dims=embed_dims,
        seed=seed,
    )
    test = prepare(settings, corpus)
    workbook = corpus.workbook_ids[0]
    sheet = "Sheet0_0"

    search_lexical = {
        "query": corpus.query,
        "top_k": 10,
        "include_structure": False,
    }
    search_expanded = {
        "query": corpus.query,
        "top_k": 10,
        "include_structure": True,
        "expand_references": True,
        "reference_depth": 1,
        "max_related_nodes": 20,
    }
    range_body = {"workbook_id": workbook, "sheet_name": sheet, "a1": "B2:D40", "limit": 50}

    endpoints = (
        _measure(
            "POST /api/v1/search/excel (lexical)",
            lambda: test.post("/api/v1/search/excel", json=search_lexical),
            iterations,
        ),
        _measure(
            "POST /api/v1/search/excel (structure + expansion depth 1)",
            lambda: test.post("/api/v1/search/excel", json=search_expanded),
            iterations,
        ),
        _measure(
            "GET /api/v1/excel/{workbook}/structure",
            lambda: test.get(f"/api/v1/excel/{workbook}/structure", params={"limit": 200}),
            iterations,
        ),
        _measure(
            "POST /api/v1/excel/range",
            lambda: test.post("/api/v1/excel/range", json=range_body),
            iterations,
        ),
    )
    expansions = _measure_expansions(test, corpus.query, iterations)
    return BenchResult(
        version=__version__,
        chunk_count=corpus.chunk_count,
        structure_count=corpus.structure_count,
        embedding_dims=embed_dims,
        endpoints=endpoints,
        expansions=expansions,
    )


def _measure(name: str, call: Any, iterations: int) -> Latency:
    samples: list[float] = []
    es_requests = 0
    for _ in range(iterations):
        elapsed, body = _time_call(call)
        samples.append(elapsed)
        es_requests = int(body.get("es_requests", es_requests))
    return Latency(
        name=name,
        iterations=iterations,
        p50_ms=percentile(samples, 50),
        p95_ms=percentile(samples, 95),
        p99_ms=percentile(samples, 99),
        mean_ms=sum(samples) / len(samples),
        es_requests=es_requests,
    )


def _measure_expansions(test: TestClient, query: str, iterations: int) -> tuple[ExpansionRow, ...]:
    """Depth 0 is "no related nodes"; depths 1 and 2 follow references."""
    rows: list[ExpansionRow] = []
    for depth, expand in ((0, False), (1, True), (2, True)):
        body = {
            "query": query,
            "top_k": 10,
            "include_structure": True,
            "expand_references": expand,
            "reference_depth": depth,
            "max_related_nodes": 50,
        }
        samples: list[float] = []
        nodes_counts: list[int] = []
        truncated = False
        es_requests = 0
        for _ in range(iterations):
            elapsed, response = _time_call(
                lambda body=body: test.post("/api/v1/search/excel", json=body)
            )
            samples.append(elapsed)
            nodes_counts.append(len(response.get("nodes", {})))
            truncated = bool(response.get("truncation", {}).get("truncated", False))
            es_requests = int(response.get("es_requests", es_requests))
        rows.append(
            ExpansionRow(
                depth=depth,
                nodes_returned=max(nodes_counts) if nodes_counts else 0,
                truncated=truncated,
                p50_ms=percentile(samples, 50),
                p95_ms=percentile(samples, 95),
                mean_nodes=sum(nodes_counts) / len(nodes_counts),
                es_requests=es_requests,
            )
        )
    return tuple(rows)


def format_report(result: BenchResult) -> str:
    lines = [
        f"excel-rag benchmark  (version {result.version})",
        f"corpus: {result.chunk_count} chunk docs, {result.structure_count} structure nodes, "
        f"{result.embedding_dims}-dim embeddings",
        "",
        "per-endpoint latency (ms) and Elasticsearch calls per request",
        f"{'endpoint':<56}{'p50':>8}{'p95':>8}{'p99':>8}{'mean':>8}{'es_req':>8}",
        "-" * 96,
    ]
    for row in result.endpoints:
        lines.append(
            f"{row.name:<56}{row.p50_ms:>8.2f}{row.p95_ms:>8.2f}{row.p99_ms:>8.2f}"
            f"{row.mean_ms:>8.2f}{row.es_requests:>8}"
        )
    lines += [
        "",
        "reference expansion (depth 0 = no related nodes; 1 and 2 follow references)",
        f"{'depth':>6}{'nodes':>8}{'mean_nodes':>12}{'truncated':>11}{'p50_ms':>9}"
        f"{'p95_ms':>9}{'es_req':>8}",
        "-" * 96,
    ]
    for expansion in result.expansions:
        lines.append(
            f"{expansion.depth:>6}{expansion.nodes_returned:>8}{expansion.mean_nodes:>12.1f}"
            f"{expansion.truncated!s:>11}{expansion.p50_ms:>9.2f}"
            f"{expansion.p95_ms:>9.2f}{expansion.es_requests:>8}"
        )
    lines += ["", f"NOTE: {result.note}."]
    return "\n".join(lines)


def main() -> int:  # pragma: no cover - entry point, exercised by running the module
    result = run_benchmark()
    print(format_report(result))
    return 0


__all__ = [
    "BenchResult",
    "ExpansionRow",
    "Latency",
    "format_report",
    "main",
    "percentile",
    "prepare",
    "run_benchmark",
]
