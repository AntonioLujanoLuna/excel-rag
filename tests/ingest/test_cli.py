"""The CLI: inspect, index, and serve (with the retrieval app named by string only)."""

from __future__ import annotations

import json

import pytest

import excel_rag.cli as cli


def test_inspect_reports_sheets_and_document_counts(build) -> None:
    report = cli.inspect_workbook(build.path("two_tables_one_sheet"))
    assert report["sheets"][0]["name"] == "Report"
    assert (
        report["documents"]["total"]
        == report["documents"]["chunks"] + report["documents"]["structure"]
    )
    assert any(region["kind"] == "table" for region in report["regions"])


def test_inspect_command_prints_json(build, capsys) -> None:
    code = cli.main(["inspect", str(build.path("two_tables_one_sheet"))])
    assert code == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["workbook"] == "two_tables.xlsx"


def test_index_command_writes_documents(build, capsys) -> None:
    code = cli.main(
        [
            "index",
            str(build.path("table_object")),
            "--workbook-id",
            "wb",
            "--version",
            "1",
            "--acl",
            "finance",
        ]
    )
    assert code == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["workbook_id"] == "wb"
    assert payload["version"] == 1
    assert payload["chunks_written"] > 0


def test_index_command_fails_cleanly_on_a_bad_workbook(build, capsys) -> None:
    code = cli.main(
        ["index", str(build.path("corrupt_file")), "--workbook-id", "wb", "--version", "1"]
    )
    assert code == 1
    assert "not a valid" in capsys.readouterr().err


def test_serve_names_the_app_by_string(build, monkeypatch) -> None:
    captured: dict[str, object] = {}

    def fake_run(app: str, **kwargs: object) -> None:
        captured["app"] = app
        captured.update(kwargs)

    monkeypatch.setattr("uvicorn.run", fake_run)
    code = cli.main(["serve", "--host", "0.0.0.0", "--port", "1234"])
    assert code == 0
    assert captured["app"] == "excel_rag.app:create_app"
    assert captured["factory"] is True
    assert captured["host"] == "0.0.0.0"
    assert captured["port"] == 1234


def test_serve_uses_settings_defaults(build, monkeypatch) -> None:
    captured: dict[str, object] = {}
    monkeypatch.setattr("uvicorn.run", lambda app, **kwargs: captured.update(kwargs))
    cli.main(["serve"])
    assert captured["host"] == "127.0.0.1"
    assert captured["port"] == 8080


def test_parser_requires_a_subcommand() -> None:
    with pytest.raises(SystemExit):
        cli.build_parser().parse_args([])


def test_render_command_prints_the_workbook_and_its_cost(build, capsys) -> None:
    code = cli.main(["render", str(build.path("large_region")), "--budget", "400", "--tools-hint"])
    assert code == 0
    captured = capsys.readouterr()
    assert captured.out.startswith("# Workbook: ")
    assert "Use the workbook tools" in captured.out
    assert "detail rows-" in captured.err and "partial" in captured.err


def test_render_command_fails_cleanly_on_a_bad_workbook(build, capsys) -> None:
    assert cli.main(["render", str(build.path("corrupt_file"))]) == 1
    assert "not a valid" in capsys.readouterr().err
