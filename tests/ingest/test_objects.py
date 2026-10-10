"""Charts, pivot tables and data validations: sheet objects that read ranges without formulas."""

from __future__ import annotations

from types import SimpleNamespace

import pytest

from excel_rag.context import WorkbookSession, render_workbook
from excel_rag.ingest import ingest_workbook
from excel_rag.models import ChunkType, NodeType, UnresolvedReason
from excel_rag.workbook import load_workbook
from excel_rag.workbook.canonical import SheetObject
from excel_rag.workbook.objects import read_sheet_objects


@pytest.fixture
def path(build):
    return build.path("sheet_objects")


def _objects(path) -> dict[str, SheetObject]:
    model = load_workbook(path, workbook_id="wb")
    return {item.name: item for sheet in model.sheets for item in sheet.objects}


def _reads(item: SheetObject) -> set[tuple[str, str]]:
    return {(reference.sheet_name, reference.a1_range) for reference in item.references}


class TestReading:
    def test_a_chart_reads_its_series_categories_and_title(self, path) -> None:
        chart = _objects(path)["Revenue by month"]
        assert (chart.kind, chart.detail, chart.anchor.a1) == ("chart", "bar chart", "D2")
        assert _reads(chart) == {("Data", "B1"), ("Data", "A2:A5"), ("Data", "B2:B5")}
        assert chart.unresolved_references == ()

    def test_a_pivot_reads_its_source_range_or_table(self, path) -> None:
        objects = _objects(path)
        assert objects["RevenuePivot"].anchor.a1 == "A3:B8"
        assert _reads(objects["RevenuePivot"]) == {("Data", "A1:B5")}
        # By table name: the whole table, header included (it names the pivot's fields).
        assert _reads(objects["TablePivot"]) == {("Data", "A1:B5")}

    def test_a_pivot_over_an_external_connection_is_a_gap(self, path) -> None:
        cube = _objects(path)["CubePivot"]
        assert cube.references == ()
        assert [gap.reason for gap in cube.unresolved_references] == [
            UnresolvedReason.EXTERNAL_LINK
        ]

    def test_a_list_validation_reads_its_source(self, path) -> None:
        validation = _objects(path)["list validation on C2:C5"]
        assert validation.anchor.a1 == "C2:C5"
        assert _reads(validation) == {("Lists", "A1:A2")}

    def test_a_constant_list_reads_nothing_and_is_not_an_object(self, path) -> None:
        assert "list validation on E2" not in _objects(path)

    def test_an_indirect_validation_is_a_gap_not_a_guess(self, path) -> None:
        validation = _objects(path)["list validation on F2"]
        assert validation.references == ()
        assert {gap.reason for gap in validation.unresolved_references} == {
            UnresolvedReason.INDIRECT
        }

    def test_a_malformed_object_is_a_warning_not_a_failure(self) -> None:
        class Broken:
            @property
            def series(self):
                raise ValueError("bad series")

        sheet = SimpleNamespace(title="S", _charts=[Broken()], _pivots=[], data_validations=None)
        objects, warnings = read_sheet_objects(sheet)
        assert objects == ()
        assert warnings == ("sheet 'S': chart 1 not read (bad series)",)


class TestIndexing:
    def test_each_object_is_a_node_with_edges_and_a_chunk(self, path) -> None:
        ingested = ingest_workbook(path, workbook_id="wb", version=1)
        nodes = {node.node_id: node for node in ingested.structure}
        chart = nodes["wb:v1:chart:data!1"]
        assert chart.node_type is NodeType.CHART
        assert chart.a1_range == "D2"
        assert {reference.a1_range for reference in chart.references} == {"B1", "A2:A5", "B2:B5"}
        # Every edge target exists, so expansion and dependents never dangle.
        for node in ingested.structure:
            for reference in node.references:
                assert reference.target_node_id in nodes
        chunk = next(item for item in ingested.chunks if item.node_id == chart.node_id)
        assert chunk.chunk_type is ChunkType.CHART
        assert "plots Data!B1, Data!A2:A5, Data!B2:B5" in chunk.content
        pivot = next(item for item in ingested.chunks if item.chunk_type is ChunkType.PIVOT_TABLE)
        assert "summarises Data!A1:B5" in pivot.content


class TestTools:
    def test_dependents_include_the_objects_reading_a_cell(self, path) -> None:
        session = WorkbookSession.load(path)
        revenue = session.dependents("Data", "B3")
        assert "Bar chart 'Revenue by month' at Data!D2" in revenue
        assert "Pivot table 'RevenuePivot' at Report!A3:B8" in revenue
        assert "List validation" not in revenue
        region = session.dependents("Lists", "A2")
        assert "List validation on Data!C2:C5: values from Lists!A1:A2" in region
        assert session.dependents("Lists", "A3").startswith("No formula, chart")

    def test_precedents_of_an_objects_cells_say_what_it_reads(self, path) -> None:
        session = WorkbookSession.load(path)
        answer = session.precedents("Report", "A3")
        assert "Pivot table 'RevenuePivot' at Report!A3:B8: summarises Data!A1:B5" in answer

    def test_the_rendering_lists_them(self, path) -> None:
        text = render_workbook(path).text
        assert "Charts, pivot tables and validations:" in text
        assert "- Bar chart 'Revenue by month' at D2: plots" in text
        assert "- Pivot table 'CubePivot' at G3:H8; unresolved: external source" in text
