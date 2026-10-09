"""The canonical helpers: deterministic ids, column letters and value formatting."""

from __future__ import annotations

import pytest

from excel_rag.workbook.canonical import (
    cell_node_id,
    chunk_id,
    column_letter,
    column_node_id,
    format_value,
    formula_node_id,
    named_range_node_id,
    range_node_id,
    region_node_id,
    row_group_node_id,
    sheet_node_id,
    table_column_node_id,
    table_node_id,
    workbook_node_id,
)


@pytest.mark.parametrize(
    ("index", "letter"),
    [(1, "A"), (26, "Z"), (27, "AA"), (52, "AZ"), (702, "ZZ"), (703, "AAA")],
)
def test_column_letter(index: int, letter: str) -> None:
    assert column_letter(index) == letter


def test_column_letter_rejects_zero() -> None:
    with pytest.raises(ValueError, match=">= 1"):
        column_letter(0)


@pytest.mark.parametrize(
    ("value", "expected"),
    [(None, ""), (True, "TRUE"), (False, "FALSE"), (5.0, "5"), (5.5, "5.5"), (7, "7"), ("x", "x")],
)
def test_format_value(value: object, expected: str) -> None:
    assert format_value(value) == expected


def test_node_id_helpers_are_deterministic_and_readable() -> None:
    assert workbook_node_id("wb42", 3) == "wb42:v3:workbook:wb42"
    assert sheet_node_id("wb42", 3, "Forecast") == "wb42:v3:sheet:forecast"
    assert cell_node_id("wb42", 3, "Assumptions", "C7") == "wb42:v3:cell:assumptions!c7"
    assert range_node_id("wb42", 3, "Actuals", "D2:D500") == "wb42:v3:range:actuals!d2:d500"
    assert region_node_id("wb42", 3, "Forecast", "A1:F25") == "wb42:v3:region:forecast!a1:f25"
    assert table_node_id("wb42", 3, "SalesTable") == "wb42:v3:table:salestable"
    assert column_node_id("wb42", 3, "Sales", "B2:B9") == "wb42:v3:column:sales!b2:b9"
    assert row_group_node_id("wb42", 3, "Sales", "A2:F50") == "wb42:v3:row_group:sales!a2:f50"
    assert formula_node_id("wb42", 3, "Forecast", "F18") == "wb42:v3:formula:forecast!f18"
    assert named_range_node_id("wb42", 3, "GrowthRate") == "wb42:v3:named_range:growthrate"
    assert (
        table_column_node_id("wb42", 3, "Sales", "SalesTable", "Revenue")
        == "wb42:v3:column:sales!salestable!revenue"
    )
    assert chunk_id("wb42", 3, "region:Report!A3:C5") == "wb42:v3:chunk:region:report!a3:c5"


def test_different_versions_never_collide() -> None:
    assert cell_node_id("wb", 1, "S", "A1") != cell_node_id("wb", 2, "S", "A1")
