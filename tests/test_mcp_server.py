"""The MCP server: the workbook tools over MCP, confined to the roots it is given."""

from __future__ import annotations

import asyncio
import json
import os
import shutil
import sys
from pathlib import Path

import pytest

pytest.importorskip("mcp.server.mcpserver")

from mcp import Client

sys.path.insert(0, str(Path(__file__).parent / "fixtures"))

import make_fixtures as mk

from excel_rag.mcp_server import PathRefused, WorkbookCatalog, build_server
from excel_rag.workbook.calc import calculation_available


@pytest.fixture
def root(tmp_path: Path) -> Path:
    books = tmp_path / "books"
    books.mkdir()
    mk.cross_sheet_formula(books)
    mk.named_range(books)
    (books / "notes.txt").write_text("not a workbook")
    (books / "~$cross_sheet.xlsx").write_bytes(b"lock")
    return books


def _call(server, name: str, arguments: dict) -> tuple[str, bool]:
    async def run() -> tuple[str, bool]:
        async with Client(server) as client:
            result = await client.call_tool(name, arguments)
            return "".join(getattr(part, "text", "") for part in result.content), result.is_error

    return asyncio.run(run())


def _tools(server) -> dict[str, object]:
    async def run():
        async with Client(server) as client:
            return {tool.name: tool for tool in (await client.list_tools()).tools}

    return asyncio.run(run())


def test_tools_are_listed_read_only(root: Path) -> None:
    tools = _tools(build_server([root]))
    assert set(tools) == {
        "list_workbooks",
        "render_workbook",
        "read_range",
        "find",
        "precedents",
        "dependents",
        "diff_workbooks",
        *(["calculate"] if calculation_available() else []),
    }
    assert all(tool.annotations.read_only_hint for tool in tools.values())
    assert "calculate" not in _tools(build_server([root], calculate=False))


def test_list_workbooks_names_only_workbooks(root: Path) -> None:
    text, error = _call(build_server([root]), "list_workbooks", {})
    assert not error
    assert text.splitlines() == ["cross_sheet.xlsx", "named_range.xlsx"]


def test_render_read_and_trace(root: Path) -> None:
    server = build_server([root])
    rendered, _ = _call(server, "render_workbook", {"path": "cross_sheet.xlsx"})
    assert "Forecast" in rendered and "Actuals" in rendered
    grid, _ = _call(
        server, "read_range", {"path": "cross_sheet.xlsx", "sheet": "Actuals", "range": "A1:D3"}
    )
    assert "| 2 |" in grid
    reads, _ = _call(
        server, "precedents", {"path": "cross_sheet.xlsx", "sheet": "Forecast", "range": "B2"}
    )
    assert "Actuals!D2:D500" in reads
    readers, _ = _call(
        server, "dependents", {"path": "cross_sheet.xlsx", "sheet": "Actuals", "range": "D100"}
    )
    assert "Forecast!B2" in readers
    found, _ = _call(server, "find", {"path": "named_range.xlsx", "query": "growthrate"})
    assert "GrowthRate" in found


def test_a_bad_call_is_a_tool_error_the_model_can_read(root: Path) -> None:
    server = build_server([root])
    text, error = _call(
        server, "read_range", {"path": "cross_sheet.xlsx", "sheet": "Nope", "range": "A1"}
    )
    assert error and "Nope" in text
    text, error = _call(server, "render_workbook", {"path": "missing.xlsx"})
    assert error and "no such workbook" in text
    text, error = _call(server, "render_workbook", {"path": "notes.txt"})
    assert error and "not an .xlsx/.xlsm/.xlsb/.csv workbook" in text


def test_paths_outside_the_roots_are_refused(root: Path, tmp_path: Path) -> None:
    outside = tmp_path / "secret.xlsx"
    shutil.copyfile(root / "cross_sheet.xlsx", outside)
    server = build_server([root])
    text, error = _call(server, "render_workbook", {"path": str(outside)})
    assert error and "outside the directories" in text
    text, error = _call(server, "render_workbook", {"path": "../secret.xlsx"})
    assert error and "no such workbook" in text
    assert "Forecast" not in text


@pytest.mark.skipif(os.name == "nt", reason="symlinks need privileges on Windows")
def test_a_symlink_out_of_the_root_is_refused(root: Path, tmp_path: Path) -> None:
    outside = tmp_path / "elsewhere.xlsx"
    shutil.copyfile(root / "cross_sheet.xlsx", outside)
    (root / "link.xlsx").symlink_to(outside)
    with pytest.raises(PathRefused):
        WorkbookCatalog([root]).resolve("link.xlsx")


def test_the_cache_reloads_an_edited_file(root: Path) -> None:
    catalog = WorkbookCatalog([root], cache_size=1)
    first = catalog.session("cross_sheet.xlsx")
    assert catalog.session("cross_sheet.xlsx") is first
    path = root / "cross_sheet.xlsx"
    stat = path.stat()
    os.utime(path, ns=(stat.st_atime_ns, stat.st_mtime_ns + 1_000_000_000))
    assert catalog.session("cross_sheet.xlsx") is not first
    catalog.session("named_range.xlsx")
    assert len(catalog._cache) == 1


def test_diff_tool(root: Path, tmp_path: Path) -> None:
    import openpyxl

    edited = openpyxl.load_workbook(root / "cross_sheet.xlsx")
    edited["Assumptions"]["C7"] = 0.07
    edited.save(root / "cross_sheet_v2.xlsx")
    text, error = _call(
        build_server([root]),
        "diff_workbooks",
        {"before": "cross_sheet.xlsx", "after": "cross_sheet_v2.xlsx"},
    )
    assert not error
    assert "Assumptions!C7" in text and "read by Forecast!B2" in text


def test_search_index_tool_runs_the_retrieval_service(root: Path) -> None:
    from excel_rag.fake_es import in_memory_client
    from excel_rag.ingest import ingest_workbook
    from excel_rag.ingest.indexer import Indexer
    from excel_rag.retrieval import Repository, RetrievalService
    from excel_rag.settings import Settings

    settings = Settings(embedding={"provider": "none"})
    client = in_memory_client(settings)
    Indexer(client, settings, embedder=None).index_workbook(
        ingest_workbook(root / "cross_sheet.xlsx", workbook_id="wb", version=1)
    )
    service = RetrievalService(Repository(client, settings), settings)
    server = build_server([root], search_service=lambda: service)
    assert "search_index" in _tools(server)
    text, error = _call(server, "search_index", {"query": "revenue", "workbook_ids": ["wb"]})
    assert not error
    hits = json.loads(text)["hits"]
    assert hits and hits[0]["source"]["workbook_id"] == "wb"
    text, error = _call(server, "search_index", {"query": "revenue", "workbook_ids": ["ghost"]})
    assert error and "ghost" in text


def test_roots_must_be_directories(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="not a directory"):
        WorkbookCatalog([tmp_path / "missing"])
    with pytest.raises(ValueError, match="at least one"):
        WorkbookCatalog([])


def test_the_cli_serves_over_stdio(root: Path) -> None:
    from mcp import StdioServerParameters

    parameters = StdioServerParameters(
        command=sys.executable,
        args=["-m", "excel_rag.cli", "mcp", "--root", str(root)],
        env={**os.environ, "EXCEL_RAG_EMBEDDING__PROVIDER": "none"},
    )

    async def run() -> str:
        async with Client(parameters, read_timeout_seconds=60) as client:
            result = await client.call_tool("list_workbooks", {})
            return "".join(getattr(part, "text", "") for part in result.content)

    assert "cross_sheet.xlsx" in asyncio.run(asyncio.wait_for(run(), timeout=90))
