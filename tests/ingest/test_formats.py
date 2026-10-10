"""CSV and binary .xlsb: read into the same raw sheets, with what they cannot carry stated."""

from __future__ import annotations

import importlib.util
import sys
import time
from pathlib import Path

import pytest

from excel_rag.context import WorkbookSession, render_workbook
from excel_rag.ingest import ingest_workbook
from excel_rag.workbook.errors import WorkbookError
from excel_rag.workbook.formats import sheet_name_for
from excel_rag.workbook.reader import read_workbook


@pytest.mark.skipif(importlib.util.find_spec("pyxlsb") is None, reason="needs the xlsb extra")
class TestXlsb:
    def test_values_types_and_shared_strings(self, build) -> None:
        raw = read_workbook(build.path("xlsb_workbook"))
        assert [sheet.name for sheet in raw.sheets] == ["Sales", "Empty"]
        cells = raw.sheets[0].cells
        assert cells["A1"].value == "Region"
        assert cells["B2"].value == 100 and cells["B2"].data_type == "n"
        assert cells["B3"].value == 150.5
        assert cells["A5"].value is True and cells["A5"].data_type == "b"
        assert raw.sheets[1].cells == {}

    def test_formula_cells_are_computed_with_saved_values_only(self, build) -> None:
        cells = read_workbook(build.path("xlsb_workbook")).sheets[0].cells
        for coordinate, saved in (("C2", 0.4), ("C3", "#DIV/0!"), ("B5", "ok")):
            cell = cells[coordinate]
            assert cell.computed and cell.formula_unavailable
            assert cell.value is None and cell.formula is None
            assert cell.cached_value == saved

    def test_warnings_say_what_was_not_read(self, build) -> None:
        warnings = " ".join(read_workbook(build.path("xlsb_workbook")).warnings)
        assert "dates appear as Excel serial numbers" in warnings
        assert "3 formula cell(s)" in warnings
        assert "1 defined name(s) not read" in warnings

    def test_the_declared_dimension_allocates_nothing(self, build) -> None:
        # The fixture declares A1:XFD1048576; reading must not pad rows to it.
        started = time.perf_counter()
        read_workbook(build.path("xlsb_workbook"))
        assert time.perf_counter() - started < 5

    def test_ingests_renders_and_answers_tools(self, build) -> None:
        path = build.path("xlsb_workbook")
        ingested = ingest_workbook(path, workbook_id="wb", version=1)
        assert any(chunk.sheet_name == "Sales" for chunk in ingested.chunks)
        text = render_workbook(path).text
        assert "Region" in text and "0.4 ƒ" in text
        assert "formula text is not read" in text
        assert "North" in WorkbookSession.load(path).find("north")

    def test_without_the_extra_the_error_says_how_to_fix_it(self, build, monkeypatch) -> None:
        path = build.path("xlsb_workbook")
        monkeypatch.setitem(sys.modules, "pyxlsb", None)
        with pytest.raises(WorkbookError, match="xlsb' extra"):
            read_workbook(path)


class TestCsv:
    def test_sniffs_semicolons_and_types_fields(self, build) -> None:
        raw = read_workbook(build.path("csv_file"))
        (sheet,) = raw.sheets
        assert sheet.name == "ventas región"
        cells = sheet.cells
        assert cells["B1"].value == "Región"  # the BOM is not part of the first header
        assert cells["A1"].value == "id"
        assert cells["A2"].value == "007" and cells["A2"].data_type == "s"
        assert cells["A3"].value == 8
        assert cells["C2"].value == 1200.5  # a comma decimal in a semicolon file
        assert cells["C4"].value == 1500.0
        assert cells["D3"].value is False
        assert cells["E2"].value == "=SUM(A1:A3)" and cells["E2"].formula is None
        assert cells["E4"].value == "3/4"
        assert "E3" not in cells

    def test_commas_tabs_and_bytes_with_a_name(self) -> None:
        raw = read_workbook(b'a,b\n1,"2,5"\n', name="upload.csv")
        cells = raw.sheets[0].cells
        assert cells["A2"].value == 1 and cells["B2"].value == "2,5"
        tabbed = read_workbook(b"x\ty\n3\t4\n", name="t.tsv").sheets[0].cells
        assert tabbed["B2"].value == 4

    def test_windows_1252_falls_back(self) -> None:
        cells = read_workbook("nom;prix\ncafé;2\n".encode("cp1252"), name="c.csv").sheets[0].cells
        assert cells["A2"].value == "café"

    def test_ingests_into_regions(self, build) -> None:
        ingested = ingest_workbook(build.path("csv_file"), workbook_id="wb", version=1)
        tables = [chunk for chunk in ingested.chunks if chunk.chunk_type.value == "table"]
        assert tables and tables[0].a1_range == "A1:E4"

    def test_bytes_named_xlsx_that_are_not_a_zip_still_fail(self) -> None:
        with pytest.raises(WorkbookError, match="not a zip archive"):
            read_workbook(b"a,b\n1,2\n", name="book.xlsx")


@pytest.mark.parametrize(
    ("label", "expected"),
    [("sales.csv", "sales"), ("dir/q[1]:2.csv", "q_1__2"), ("", "Sheet1"), ("x" * 40, "x" * 31)],
)
def test_sheet_names_are_valid(label: str, expected: str) -> None:
    assert sheet_name_for(label) == expected


def test_catalog_lists_the_new_formats(build, tmp_path: Path) -> None:
    pytest.importorskip("mcp.server.mcpserver")
    from excel_rag.mcp_server import WorkbookCatalog

    build.path("xlsb_workbook")
    build.path("csv_file")
    names = WorkbookCatalog([tmp_path]).list_workbooks()
    assert "sales.xlsb" in names and "ventas región.csv" in names
