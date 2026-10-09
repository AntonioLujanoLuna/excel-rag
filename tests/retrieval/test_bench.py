"""Benchmark-harness tests: the corpus is well-formed and the harness runs and reports."""

from __future__ import annotations

from excel_rag.bench.corpus import build_corpus
from excel_rag.bench.runner import format_report, percentile, prepare, run_benchmark
from excel_rag.settings import Settings


class TestCorpus:
    def test_counts_match_the_parameters(self) -> None:
        corpus = build_corpus(
            workbooks=1, sheets=1, cells_per_sheet=10, chunks_per_sheet=40, embed_dims=8
        )
        assert corpus.chunk_count == 40
        assert corpus.structure_count > 40
        assert corpus.workbook_ids == ("wb0",)

    def test_every_chunk_node_id_resolves_to_a_structure_node(self) -> None:
        corpus = build_corpus(
            workbooks=1, sheets=1, cells_per_sheet=10, chunks_per_sheet=40, embed_dims=8
        )
        structure_ids = {identifier for identifier, _ in corpus.structure}
        assert all(document["node_id"] in structure_ids for _, document in corpus.chunks)

    def test_is_deterministic_for_a_seed(self) -> None:
        first = build_corpus(
            workbooks=1, sheets=1, cells_per_sheet=5, chunks_per_sheet=30, embed_dims=4
        )
        second = build_corpus(
            workbooks=1, sheets=1, cells_per_sheet=5, chunks_per_sheet=30, embed_dims=4
        )
        assert first.chunks == second.chunks


class TestRunner:
    def test_prepare_indexes_and_serves_a_search(self) -> None:
        settings = Settings()
        corpus = build_corpus(
            workbooks=1, sheets=1, cells_per_sheet=10, chunks_per_sheet=40, embed_dims=8
        )
        test = prepare(settings, corpus)
        response = test.post(
            "/api/v1/search/excel", json={"query": "revenue", "include_structure": False}
        )
        assert response.status_code == 200
        assert response.json()["hits"]

    def test_run_benchmark_measures_every_endpoint(self) -> None:
        result = run_benchmark(
            iterations=2,
            workbooks=1,
            sheets=1,
            cells_per_sheet=20,
            chunks_per_sheet=40,
            embed_dims=8,
        )
        assert len(result.endpoints) == 4
        assert len(result.expansions) == 3
        assert all(row.iterations == 2 for row in result.endpoints)
        assert result.chunk_count == 40

    def test_report_states_the_in_memory_caveat(self) -> None:
        result = run_benchmark(
            iterations=1,
            workbooks=1,
            sheets=1,
            cells_per_sheet=20,
            chunks_per_sheet=40,
            embed_dims=8,
        )
        report = format_report(result)
        assert "IN-MEMORY" in report
        assert "NOT an Elasticsearch latency" in report


class TestPercentile:
    def test_nearest_rank(self) -> None:
        samples = [1.0, 2.0, 3.0, 4.0]
        assert percentile(samples, 50) == 2.0
        assert percentile(samples, 99) == 4.0
        assert percentile(samples, 100) == 4.0

    def test_empty_is_zero(self) -> None:
        assert percentile([], 95) == 0.0
