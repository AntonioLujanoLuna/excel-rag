"""The workbook tools: what a model can ask about an attached workbook beyond its rendering."""

from __future__ import annotations

from typing import Any

import pytest

from excel_rag.context import (
    FORMULA_MARK,
    MAX_RANGE_CELLS,
    TOOL_DEFINITIONS,
    WorkbookSession,
)


@pytest.fixture
def session(fixtures) -> WorkbookSession:
    return WorkbookSession.load(fixtures.cross_sheet_formula())


class TestDefinitions:
    def test_every_tool_is_strict_with_every_field_required(self) -> None:
        for definition in TOOL_DEFINITIONS:
            schema = definition["input_schema"]
            assert definition["strict"] is True
            assert schema["additionalProperties"] is False
            assert set(schema["required"]) == set(schema["properties"])
            assert len(definition["description"]) > 80, "a model picks tools by description"

    def test_definitions_are_copies(self) -> None:
        definitions = WorkbookSession.tool_definitions()
        definitions[0]["name"] = "changed"
        assert TOOL_DEFINITIONS[0]["name"] == "read_range"


class TestReadRange:
    def test_a_grid_with_coordinates(self, session: WorkbookSession) -> None:
        text = session.read_range("Actuals", "D1:D6")
        assert text.startswith("Actuals!D1:D6:")
        assert "| 2 | 10 |" in text and "| 6 | 50 |" in text

    def test_saved_formula_values_are_marked_and_listed(self, fixtures) -> None:
        text = WorkbookSession.load(fixtures.cached_value()).read_range("Calc", "A1:C1")
        assert f"| 1 | 2 | 3 | 999 {FORMULA_MARK} |" in text
        assert "- Calc!C1: =A1+B1 → 999; reads Calc!A1, Calc!B1" in text

    def test_a_large_range_is_cut_by_rows_and_names_the_rest(self, fixtures) -> None:
        text = WorkbookSession.load(fixtures.large_region()).read_range("Data", "A1:F501")
        cells = sum(line.count(" | ") for line in text.splitlines() if line.startswith("| "))
        assert cells <= MAX_RANGE_CELLS + 400
        assert "read_range A334:F501 for the rest" in text

    def test_an_empty_range_says_so(self, session: WorkbookSession) -> None:
        assert session.read_range("Actuals", "H1:H5") == "Actuals!H1:H5 is empty."

    def test_sheet_names_match_case_insensitively(self, session: WorkbookSession) -> None:
        assert session.read_range("actuals", "D2").startswith("Actuals!D2:")


class TestFind:
    def test_values_and_formulas_are_searched(self, session: WorkbookSession) -> None:
        assert "- Forecast!A2: Revenue" in session.find("revenue")
        assert "Forecast!B2" in session.find("Assumptions!C7")

    def test_named_ranges_are_found(self, fixtures) -> None:
        text = WorkbookSession.load(fixtures.named_range()).find("growthrate")
        assert "named range GrowthRate = Assumptions!C7" in text

    def test_matches_are_capped_and_counted(self, fixtures) -> None:
        text = WorkbookSession.load(fixtures.large_region()).find("1")
        assert "more not shown" in text

    def test_no_match_is_said(self, session: WorkbookSession) -> None:
        assert session.find("zebra") == "No cell contains 'zebra'."


class TestPrecedentsAndDependents:
    def test_precedents_list_reads_with_values(self, session: WorkbookSession) -> None:
        text = session.precedents("Forecast", "B2")
        assert "=SUM(Actuals!D2:D500)*(1+Assumptions!C7)" in text
        assert "    - Actuals!D2:D500 (499 cells)" in text
        assert "    - Assumptions!C7 = 0.05" in text

    def test_an_input_range_has_no_precedents(self, session: WorkbookSession) -> None:
        assert "holds no formulas" in session.precedents("Actuals", "D2:D6")

    def test_a_cell_inside_a_read_range_finds_its_dependent(self, session: WorkbookSession) -> None:
        assert "Forecast!B2: =SUM(" in session.dependents("Actuals", "D100")
        assert "Forecast!B2" in session.dependents("Assumptions", "C7")

    def test_an_unread_cell_has_no_dependents(self, session: WorkbookSession) -> None:
        assert session.dependents("Assumptions", "C8") == "No formula reads Assumptions!C8."

    def test_unresolvable_references_are_named(self, fixtures) -> None:
        text = WorkbookSession.load(fixtures.indirect_offset()).precedents("Calc", "A3:A4")
        assert "unresolved: INDIRECT(...) (indirect)" in text
        assert "unresolved: OFFSET(...) (volatile_offset)" in text


class TestDispatch:
    @pytest.mark.parametrize(
        ("name", "tool_input", "message"),
        [
            ("read_range", {"sheet": "Nope", "range": "A1"}, "no sheet named 'Nope'"),
            ("read_range", {"sheet": "Forecast", "range": "A1:B"}, "not an A1 cell"),
            ("read_range", {"sheet": "Forecast"}, "non-empty string 'range'"),
            ("find", {"query": "x", "extra": 1}, "takes no 'extra'"),
            ("find", {"query": "  "}, "non-empty string 'query'"),
            ("find", ["not", "an", "object"], "must be a JSON object"),
            ("delete", {}, "unknown tool 'delete'"),
        ],
    )
    def test_a_bad_call_is_an_error_result_not_an_exception(
        self, session: WorkbookSession, name: str, tool_input: Any, message: str
    ) -> None:
        block = session.tool_result("toolu_1", name, tool_input)
        assert block["type"] == "tool_result"
        assert block["tool_use_id"] == "toolu_1"
        assert block["is_error"] is True
        assert message in block["content"]

    def test_a_good_call_has_no_error_flag(self, session: WorkbookSession) -> None:
        block = session.tool_result("toolu_2", "find", {"query": "revenue"})
        assert "is_error" not in block
        assert "Forecast!A2" in block["content"]

    def test_a_sheet_prefixed_range_is_accepted(self, session: WorkbookSession) -> None:
        outcome = session.call("read_range", {"sheet": "Actuals", "range": "Actuals!$D$2"})
        assert not outcome.is_error and "| 2 | 10 |" in outcome.content

    def test_a_macro_sheet_cannot_be_read(self, fixtures) -> None:
        session = WorkbookSession.load(fixtures.macro_workbook())
        outcome = session.call("read_range", {"sheet": "Macro1", "range": "A1"})
        assert outcome.is_error and "macro sheet" in outcome.content
