"""The acceptance cases named in the task, each asserted end to end."""

from __future__ import annotations

from fixtures import make_fixtures as mk

from excel_rag.es import INDEX_CHUNKS, INDEX_STRUCTURE
from excel_rag.fake_es import in_memory_client
from excel_rag.ingest import ingest_workbook, read_workbook
from excel_rag.ingest.canonical import RegionKind
from excel_rag.ingest.indexer import Indexer
from excel_rag.models import NodeType, UnresolvedReason
from excel_rag.settings import Settings


def test_two_independent_tables_on_one_sheet(build) -> None:
    ingested = ingest_workbook(build.path("two_tables_one_sheet"), workbook_id="wb", version=1)
    tables = [
        region
        for sheet in ingested.model.sheets
        for region in sheet.regions
        if region.kind is RegionKind.TABLE
    ]
    assert len(tables) == 2
    assert [column.name for column in tables[0].columns] == ["Region", "Q1", "Q2"]
    assert [column.name for column in tables[1].columns] == ["Product", "Units"]


def test_merged_multilevel_header(build) -> None:
    ingested = build.ingest("merged_multilevel_header")
    region = ingested.model.sheets[0].regions[0]
    assert region.header_rows == (1, 2)
    assert "Sales / Q1" in [column.name for column in region.columns]


def test_units_and_notes_rows(build) -> None:
    ingested = build.ingest("units_and_notes")
    table = next(
        region for region in ingested.model.sheets[0].regions if region.kind is RegionKind.TABLE
    )
    assert table.units_row == 2
    assert any(region.kind is RegionKind.NOTES for region in ingested.model.sheets[0].regions)


def test_cross_sheet_formula_resolves_without_reasoning(build) -> None:
    ingested = build.ingest("cross_sheet_formula")
    node_ids = {node.node_id for node in ingested.structure}
    formula = next(node for node in ingested.structure if node.node_type is NodeType.FORMULA)
    targets = {
        (reference.sheet_name, reference.a1_range, reference.kind.value)
        for reference in formula.references
    }
    assert ("Actuals", "D2:D500", "range") in targets
    assert ("Assumptions", "C7", "cell") in targets
    assert all(reference.target_node_id in node_ids for reference in formula.references)


def test_named_range_resolves(build) -> None:
    ingested = build.ingest("named_range")
    assert ingested.model.named_ranges[0].resolved
    formula = next(node for node in ingested.structure if node.node_type is NodeType.FORMULA)
    assert any(reference.kind.value == "named_range" for reference in formula.references)


def test_indirect_and_offset_reasons(build) -> None:
    ingested = build.ingest("indirect_offset")
    reasons = {gap.reason for node in ingested.structure for gap in node.unresolved_references}
    assert {UnresolvedReason.INDIRECT, UnresolvedReason.VOLATILE_OFFSET} <= reasons


def test_500_by_6_region_is_bounded(build) -> None:
    ingested = ingest_workbook(build.path("large_region"), workbook_id="wb", version=1)
    total = len(ingested.chunks) + len(ingested.structure)
    assert total == 40  # measured; independent of the 3000 cells in the grid
    assert total < 3000 / 10


def test_cached_value_is_surfaced_as_cached(build) -> None:
    ingested = build.ingest("cached_value")
    formula = next(node for node in ingested.structure if node.node_type is NodeType.FORMULA)
    assert formula.formula == "=A1+B1"
    assert formula.cached_value == 999  # stale value from the file, not A1+B1=5


def test_macro_and_vba_are_never_loaded(build) -> None:
    raw = read_workbook(build.path("macro_workbook"))
    assert raw.has_vba and raw.macro_sheet_names == ("Macro1",)
    ingested = ingest_workbook(build.path("macro_workbook"), workbook_id="wb", version=1)
    reasons = {gap.reason for node in ingested.structure for gap in node.unresolved_references}
    assert UnresolvedReason.MACRO_SHEET in reasons


def test_oversized_declared_dimension_not_allocated(build) -> None:
    raw = read_workbook(build.path("hostile_dimension"))
    assert raw.sheets[0].declared_dimension_flagged is True
    assert raw.sheets[0].cells and max(cell.row for cell in raw.sheets[0].cells.values()) == 2


def test_manifest_flip_leaves_only_the_new_version(build) -> None:
    settings = Settings()
    client = in_memory_client(settings)
    indexer = Indexer(client, settings)
    indexer.index_workbook(
        ingest_workbook(build.path("two_tables_one_sheet"), workbook_id="wb", version=1)
    )
    second = ingest_workbook(build.path("merged_multilevel_header"), workbook_id="wb", version=2)
    indexer.index_workbook(second)

    assert indexer.active_version("wb") == 2
    assert client.count(INDEX_CHUNKS, {"term": {"version": 1}}) == 0
    assert client.count(INDEX_STRUCTURE, {"term": {"version": 1}}) == 0
    assert client.count(INDEX_CHUNKS, {"term": {"version": 2}}) == len(second.chunks)
    hits = client.search(INDEX_CHUNKS, {"match_all": {}}, size=200)["hits"]["hits"]
    assert hits and all(hit["_source"]["version"] == 2 for hit in hits)


def test_realistic_workbook_is_bounded(build) -> None:
    """A small multi-sheet workbook with formulas stays well inside the budget."""
    path = mk.cross_sheet_formula(build.directory)
    ingested = ingest_workbook(path, workbook_id="wb", version=1)
    total = len(ingested.chunks) + len(ingested.structure)
    assert 0 < total < 100
