"""Region detection: two tables on one sheet, merged headers, units and notes blocks."""

from __future__ import annotations

from excel_rag.ingest.canonical import RegionKind
from excel_rag.ingest.reader import read_workbook
from excel_rag.ingest.regions import RegionConfig, detect_regions


def _regions(build, name: str, **kwargs):
    raw = read_workbook(build.path(name))
    config = RegionConfig(**kwargs) if kwargs else None
    return detect_regions(raw.sheets[0], workbook_id="wb", version=1, config=config)


def test_two_tables_on_one_sheet_are_two_regions(build) -> None:
    regions = _regions(build, "two_tables_one_sheet")
    kinds = [(region.kind, region.a1_range.a1) for region in regions]
    assert kinds == [
        (RegionKind.TITLE, "A1:C1"),
        (RegionKind.TABLE, "A3:C5"),
        (RegionKind.TABLE, "A7:C9"),
    ]
    first, second = regions[1], regions[2]
    assert [column.name for column in first.columns] == ["Region", "Q1", "Q2"]
    assert [column.name for column in second.columns] == ["Product", "Units"]
    assert first.a1_range != second.a1_range


def test_merged_multilevel_header_is_one_table_with_two_header_rows(build) -> None:
    regions = _regions(build, "merged_multilevel_header")
    assert len(regions) == 1
    region = regions[0]
    assert region.kind is RegionKind.TABLE
    assert region.header_rows == (1, 2)
    assert [column.name for column in region.columns] == [
        "Sales / Q1",
        "Sales / Q2",
        "Costs / Q1",
        "Costs / Q2",
    ]
    assert region.columns[0].aliases == ("Sales", "Q1")


def test_units_row_is_peeled_from_the_header(build) -> None:
    regions = _regions(build, "units_and_notes")
    table = next(region for region in regions if region.kind is RegionKind.TABLE)
    assert table.units_row == 2
    assert table.header_rows == (1,)
    assert {column.unit for column in table.columns} == {"USD"}


def test_notes_block_is_its_own_region(build) -> None:
    regions = _regions(build, "units_and_notes")
    notes = next(region for region in regions if region.kind is RegionKind.NOTES)
    assert notes.notes == ("Notes:", "Revenue is reported in constant currency.")


def test_table_object_names_the_region(build) -> None:
    regions = _regions(build, "table_object")
    assert len(regions) == 1
    assert regions[0].table_name == "SalesTable"
    assert regions[0].node_type_kind == "table"


def test_row_group_count_is_bounded(build) -> None:
    regions = _regions(build, "large_region")
    table = regions[0]
    assert len(table.row_groups) <= RegionConfig().max_row_groups_per_region
    assert sum(len(group.rows) for group in table.row_groups) == 500


def test_row_groups_are_capped_however_tall_the_region(build) -> None:
    regions = _regions(build, "large_region", row_group_rows=5, max_row_groups_per_region=4)
    assert len(regions[0].row_groups) == 4


def test_empty_sheet_has_no_regions(build) -> None:
    import openpyxl

    wb = openpyxl.Workbook()
    path = build.directory / "empty.xlsx"
    wb.save(path)
    raw = read_workbook(path)
    assert detect_regions(raw.sheets[0], workbook_id="wb", version=1) == ()


def test_headerless_grid_still_detects_columns(build) -> None:
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Raw"
    for row in range(1, 4):
        for col in range(1, 3):
            ws.cell(row=row, column=col, value=row * col)
    path = build.directory / "headerless.xlsx"
    wb.save(path)
    raw = read_workbook(path)
    regions = detect_regions(raw.sheets[0], workbook_id="wb", version=1)
    assert len(regions) == 1
    assert len(regions[0].columns) == 2
    assert all(column.name.startswith("Column") for column in regions[0].columns)


def test_column_schema_carries_type_and_examples(build) -> None:
    regions = _regions(build, "units_and_notes")
    table = next(region for region in regions if region.kind is RegionKind.TABLE)
    revenue = next(column for column in table.columns if column.name == "Revenue")
    assert revenue.inferred_type == "currency" or revenue.inferred_type == "number"
    assert revenue.distinct_count == 2
    assert revenue.a1_range.a1 == "A3:A4"
    assert revenue.node_id.endswith("column:report!a3:a4") or "column:" in revenue.node_id
