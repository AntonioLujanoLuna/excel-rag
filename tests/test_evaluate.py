"""The retrieval-quality harness: case files, relevance, metrics, and the sample run."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from excel_rag.cli import main as cli_main
from excel_rag.embedding import HashingEmbedder
from excel_rag.evaluate import (
    CaseFileError,
    Configuration,
    EvalCase,
    Expected,
    aggregate,
    evaluate,
    format_report,
    parse_cases,
    score_case,
)
from excel_rag.evaluate.__main__ import main as evaluate_main
from excel_rag.evaluate.sample import SAMPLE_CASES, SAMPLE_WORKBOOK_ID, build_sample_workbook
from excel_rag.fake_es import _text_matches
from excel_rag.models import Hit, SourceRef
from excel_rag.rerank import OverlapReranker

CASE = EvalCase(
    id="c",
    question="q",
    workbook_id="wb",
    expected=(Expected("Forecast", "A3:E7"), Expected("Assumptions", "B4")),
)


def _hit(sheet: str, a1: str, kind: str = "region", workbook: str = "wb") -> Hit:
    return Hit(
        chunk_id=f"{sheet}{a1}",
        score=1.0,
        content="",
        source=SourceRef(workbook_id=workbook, version=1, sheet=sheet, a1_range=a1),
        node_id=f"{workbook}:v1:{kind}:{sheet.lower()}!{a1.lower()}",
    )


class TestCases:
    def test_parses_json_lines_and_skips_comments(self) -> None:
        lines = [
            "# a comment",
            "",
            json.dumps(
                {"question": "q", "workbook_id": "wb", "expected": [{"sheet": "S", "a1": "A1"}]}
            ),
        ]
        (case,) = parse_cases(lines)
        assert case.id == "line-3"
        assert case.expected == (Expected("S", "A1"),)

    @pytest.mark.parametrize(
        "line",
        [
            "not json",
            json.dumps({"question": "q", "workbook_id": "wb"}),
            json.dumps({"question": "q", "workbook_id": "wb", "expected": []}),
            json.dumps(
                {"question": "q", "workbook_id": "wb", "expected": [{"sheet": "S", "a1": "7C"}]}
            ),
        ],
    )
    def test_a_bad_line_names_its_line(self, line: str) -> None:
        with pytest.raises(CaseFileError, match=":1:"):
            parse_cases([line])


class TestRelevance:
    def test_intersection_on_the_expected_sheet_is_relevant(self) -> None:
        result = score_case(CASE, [_hit("Forecast", "D4:D7")])
        assert result.first_rank == 1

    def test_other_sheets_workbooks_and_summaries_are_not(self) -> None:
        hits = [
            _hit("Actuals", "A3:E7"),
            _hit("Forecast", "A3:E7", workbook="other"),
            _hit("Forecast", "A1:E12", kind="sheet"),
            _hit("Forecast", "", kind="workbook"),
            _hit("Forecast", "G1:H2"),
            _hit("Assumptions", "A4:B7"),
        ]
        result = score_case(CASE, hits)
        assert result.first_rank == 6
        assert result.recall_at(10) == 0.5

    def test_metrics_aggregate_over_cases(self) -> None:
        found_second = score_case(CASE, [_hit("Actuals", "A1"), _hit("Forecast", "A3")])
        missed = score_case(CASE, [_hit("Actuals", "A1")])
        both = score_case(CASE, [_hit("Forecast", "A3"), _hit("Assumptions", "B4")])
        metrics = aggregate([found_second, missed, both], (1, 2))
        assert metrics.hit_at == {1: pytest.approx(1 / 3), 2: pytest.approx(2 / 3)}
        assert metrics.recall_at[2] == pytest.approx((0.5 + 0 + 1) / 3)
        assert metrics.mrr == pytest.approx((0.5 + 0 + 1) / 3)

    def test_no_cases_is_zero_not_a_division_error(self) -> None:
        assert aggregate([], (1,)).mrr == 0.0


class TestSampleRun:
    def test_sample_cases_cite_real_ranges(self, tmp_path: Path) -> None:
        build_sample_workbook(tmp_path)
        assert len(SAMPLE_CASES) >= 12
        assert {case.workbook_id for case in SAMPLE_CASES} == {SAMPLE_WORKBOOK_ID}

    def test_every_configuration_runs_and_finds_most_lexical_cases(self, tmp_path: Path) -> None:
        workbooks = {SAMPLE_WORKBOOK_ID: build_sample_workbook(tmp_path)}
        report = evaluate(
            [
                Configuration("lexical"),
                Configuration("hybrid", embedder=HashingEmbedder(64)),
                Configuration(
                    "hybrid+rerank", embedder=HashingEmbedder(64), reranker=OverlapReranker()
                ),
            ],
            SAMPLE_CASES,
            workbooks,
            k_values=(1, 5),
        )
        assert [run.configuration for run in report.runs] == ["lexical", "hybrid", "hybrid+rerank"]
        lexical = report.runs[0].metrics
        assert lexical.cases == len(SAMPLE_CASES)
        # The English, word-overlapping half of the sample is a lexical floor.
        assert lexical.hit_at[5] >= 0.5
        assert report.runs[1].embedding_model == "hashing-test-double"
        text = format_report(report)
        assert "hit@5" in text and "IN-MEMORY" in text

    def test_a_case_naming_an_unknown_workbook_is_refused(self, tmp_path: Path) -> None:
        stray = EvalCase("x", "q", "missing", (Expected("S", "A1"),))
        with pytest.raises(ValueError, match="missing"):
            evaluate([Configuration("lexical")], [stray], {})


class TestCommandLine:
    def test_runs_the_sample_with_the_doubles(self, capsys) -> None:
        assert evaluate_main(["--embedder", "hashing", "--rerank", "overlap", "-k", "3"]) == 0
        out = capsys.readouterr().out
        assert "hybrid+rerank" in out and "hit@3" in out

    def test_json_output(self, capsys) -> None:
        assert evaluate_main(["--embedder", "none", "--rerank", "none", "--json"]) == 0
        payload = json.loads(capsys.readouterr().out)
        assert [run["configuration"] for run in payload["runs"]] == ["lexical"]

    def test_a_missing_extra_skips_the_run_instead_of_failing(self, capsys, monkeypatch) -> None:
        import sys

        monkeypatch.setitem(sys.modules, "sentence_transformers", None)
        monkeypatch.setenv("EXCEL_RAG_EMBEDDING__PROVIDER", "sentence-transformers")
        assert evaluate_main(["--rerank", "none", "--no-misses"]) == 0
        captured = capsys.readouterr()
        assert "skipping the hybrid run" in captured.err
        assert "lexical" in captured.out and "hybrid" not in captured.out.split("Note:")[0]

    def test_own_cases_and_workbooks(self, tmp_path: Path, capsys) -> None:
        path = build_sample_workbook(tmp_path)
        cases = tmp_path / "cases.jsonl"
        cases.write_text(
            json.dumps(
                {
                    "id": "units",
                    "question": "units sold",
                    "workbook_id": "mine",
                    "expected": [{"sheet": "Actuals", "a1": "E1:E25"}],
                }
            )
        )
        argv = ["evaluate", "--cases", str(cases), "--workbook", f"mine={path}"]
        assert cli_main([*argv, "--embedder", "none", "--rerank", "none"]) == 0
        assert "1.00" in capsys.readouterr().out

    def test_a_malformed_workbook_flag(self, tmp_path: Path) -> None:
        cases = tmp_path / "cases.jsonl"
        cases.write_text("")
        assert evaluate_main(["--cases", str(cases), "--workbook", "no-equals"]) == 2


def test_the_in_memory_match_splits_like_the_standard_analyzer() -> None:
    assert _text_matches("Department=Engineering; Employees=42", "engineering?")
    assert _text_matches("Revenue/cost", "cost")
    assert not _text_matches("Engineering", "engineer")


class TestMining:
    def test_labelled_formulas_and_inputs_become_cases(self, tmp_path: Path) -> None:
        from excel_rag.evaluate import mine_cases
        from excel_rag.workbook import load_workbook

        model = load_workbook(build_sample_workbook(tmp_path), workbook_id=SAMPLE_WORKBOOK_ID)
        cases = {case.question: case for case in mine_cases(model)}
        # A column of identical formulas is one case under its header.
        assert cases["How is EBITDA calculated?"].expected == (Expected("Forecast", "D4:D7"),)
        # A formula of its own is named by its row label and its header.
        assert cases["How is Q1 Revenue calculated?"].expected == (Expected("Forecast", "B4"),)
        # An input a formula reads is answered by its label and value together.
        assert cases["What is the Tax rate?"].expected == (Expected("Assumptions", "A5:B5"),)
        # An input nothing reads is not asked about.
        assert "What is the EUR/USD exchange rate?" not in cases
        assert all(case.workbook_id == SAMPLE_WORKBOOK_ID for case in cases.values())
        assert len({case.id for case in cases.values()}) == len(cases)

    def test_mining_is_deterministic_and_bounded(self, tmp_path: Path) -> None:
        from excel_rag.evaluate import mine_cases
        from excel_rag.workbook import load_workbook

        model = load_workbook(build_sample_workbook(tmp_path))
        assert mine_cases(model) == mine_cases(model)
        assert len(mine_cases(model, limit=3)) == 3

    def test_mine_and_save_from_the_command_line(self, tmp_path: Path, capsys) -> None:
        saved = tmp_path / "mined.jsonl"
        argv = ["--mine", "--embedder", "none", "--rerank", "none", "--save-cases", str(saved)]
        assert evaluate_main(argv) == 0
        cases = parse_cases(saved.read_text().splitlines())
        assert any(case.question == "How is EBITDA calculated?" for case in cases)
        assert f"lexical        {len(cases)}" in capsys.readouterr().out

    def test_cases_and_mine_are_alternatives(self, tmp_path: Path) -> None:
        cases = tmp_path / "cases.jsonl"
        cases.write_text("")
        assert evaluate_main(["--cases", str(cases), "--mine"]) == 2


class TestThresholds:
    def test_parse(self) -> None:
        from excel_rag.evaluate import Threshold

        assert Threshold.parse("mrr=0.5") == Threshold("mrr", 0.5)
        assert Threshold.parse("hybrid+rerank:HIT@5=0.7") == Threshold(
            "hit@5", 0.7, "hybrid+rerank"
        )
        for bad in ("mrr", "ndcg=0.5", "mrr=high", "hit@=1"):
            with pytest.raises(ValueError):
                Threshold.parse(bad)

    def test_a_run_below_a_floor_fails_the_command(self, capsys) -> None:
        argv = ["--embedder", "none", "--rerank", "none", "--no-misses"]
        assert evaluate_main([*argv, "--min", "mrr=0.0", "--min", "lexical:hit@10=0.1"]) == 0
        assert evaluate_main([*argv, "--min", "lexical:mrr=1.01"]) == 3
        assert "lexical mrr" in capsys.readouterr().err

    def test_a_floor_on_a_run_that_did_not_happen_fails(self, capsys) -> None:
        argv = ["--embedder", "none", "--rerank", "none", "--min", "hybrid:mrr=0.1"]
        assert evaluate_main(argv) == 3
        assert "no 'hybrid' run" in capsys.readouterr().err

    def test_a_cut_off_that_was_not_measured_is_an_error(self) -> None:
        argv = ["--embedder", "none", "--rerank", "none", "-k", "3", "--min", "hit@5=0.1"]
        assert evaluate_main(argv) == 1

    def test_a_malformed_floor_is_an_error(self) -> None:
        assert evaluate_main(["--embedder", "none", "--rerank", "none", "--min", "mrr"]) == 1
