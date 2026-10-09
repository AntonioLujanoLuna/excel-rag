"""The context renderer: what a model reads when a workbook is attached to a conversation."""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path

import pytest

from excel_rag.context import FORMULA_MARK, approx_tokens, render_workbook
from excel_rag.workbook import WorkbookError, load_workbook


class TestComplete:
    def test_grids_carry_row_numbers_column_letters_and_values(self, fixtures) -> None:
        rendered = render_workbook(fixtures.units_and_notes())
        assert rendered.complete and rendered.detail == "full"
        text = rendered.text
        assert '### Table "Revenue" - Sales!A1:B4, header row(s) 1' in text
        assert "Columns: A Revenue (number, USD); B Cost (number, USD)" in text
        assert "| | A | B |" in text
        assert "| 3 | 100 | 50 |" in text
        assert "> Revenue is reported in constant currency." in text
        assert "partial" not in text

    def test_formulas_list_their_saved_value_and_what_they_read(self, fixtures) -> None:
        text = render_workbook(fixtures.cached_value()).text
        assert f"| 1 | 2 | 3 | 999 {FORMULA_MARK} |" in text
        assert "- C1: =A1+B1 → 999; reads Calc!A1, Calc!B1" in text
        assert "last saved; it is not recalculated" in text

    def test_no_saved_values_is_said_not_shown_as_blanks(self, fixtures) -> None:
        text = render_workbook(fixtures.cross_sheet_formula()).text
        assert "carry no saved values" in text
        assert "| 2 | Revenue | =SUM(Actuals!D2:D500)*(1+Assumptions!C7) |" in text
        assert "→ no saved value; reads Actuals!D2:D500, Assumptions!C7" in text

    def test_unresolvable_references_are_named(self, fixtures) -> None:
        text = render_workbook(fixtures.indirect_offset()).text
        assert "unresolved: INDIRECT(...) (indirect)" in text
        assert "unresolved: OFFSET(...) (volatile_offset)" in text

    def test_named_ranges_and_macros_are_stated(self, fixtures) -> None:
        assert "Named ranges: GrowthRate = Assumptions!C7" in (
            render_workbook(fixtures.named_range()).text
        )
        macro = render_workbook(fixtures.macro_workbook()).text
        assert "Contains macros; they were not loaded or run." in macro
        assert "macro sheet, not loaded" in macro

    def test_values_read_as_excel_shows_them(self, typed_values: Path) -> None:
        text = render_workbook(typed_values).text
        assert "| 2 | Launch | 2026-03-31 | TRUE | 25% |" in text
        assert "| 3 | Review | 2026-04-02 15:30:00 | FALSE | 12.5% | ok |" in text
        assert "go \\| no-go" in text, "a pipe inside a cell must not split the grid"

    def test_a_lone_cell_is_rendered(self, typed_values: Path) -> None:
        assert '### Title "stray note" - Plan!H10' in render_workbook(typed_values).text

    def test_a_cell_no_region_covers_is_still_listed(self, typed_values: Path) -> None:
        """Region detection is a heuristic; a cell it leaves out must not vanish from the text."""
        model = load_workbook(typed_values)
        sheet = model.sheets[0]
        tables_only = tuple(region for region in sheet.regions if region.a1_range.a1 != "H10")
        uncovered = replace(model, sheets=(replace(sheet, regions=tables_only),))
        assert "Other cells: H10 = stray note" in render_workbook(uncovered).text

    def test_bytes_and_a_loaded_model_render_the_same(self, fixtures) -> None:
        path = fixtures.units_and_notes()
        from_path = render_workbook(path).text
        from_bytes = render_workbook(path.read_bytes(), name=path.name).text
        from_model = render_workbook(load_workbook(path)).text
        assert from_path == from_bytes == from_model


class TestBudget:
    def test_a_large_table_keeps_its_header_first_and_last_rows(self, fixtures) -> None:
        rendered = render_workbook(fixtures.large_region(), token_budget=1_500)
        assert not rendered.complete
        assert rendered.tokens <= 1_500
        text = rendered.text
        assert "| 1 | col_0 | col_1 |" in text
        assert "| 2 | 1 | 2 |" in text
        assert "| 501 | 2995 |" in text
        assert "omitted" in text and "rows with data" in text
        assert "This rendering is partial" in text

    def test_the_most_detailed_rendering_that_fits_is_chosen(self, fixtures) -> None:
        path = fixtures.large_region()
        roomy = render_workbook(path, token_budget=50_000)
        tight = render_workbook(path, token_budget=600)
        assert roomy.complete and roomy.detail == "full"
        assert not tight.complete
        assert len(tight.text) < len(roomy.text)

    def test_every_budget_is_honoured(self, fixtures) -> None:
        path = fixtures.large_region()
        for budget in (60, 150, 400, 1_000, 3_000):
            assert render_workbook(path, token_budget=budget).tokens <= budget

    def test_long_text_is_shortened_only_when_detail_drops(self, typed_values: Path) -> None:
        full = render_workbook(typed_values).text
        assert "x" * 200 in full
        short = render_workbook(typed_values, token_budget=170).text
        assert "x" * 200 not in short

    def test_the_tools_hint_is_added_only_when_something_was_left_out(self, fixtures) -> None:
        assert "read_range" not in render_workbook(fixtures.units_and_notes(), tools_hint=True).text
        partial = render_workbook(fixtures.large_region(), token_budget=800, tools_hint=True)
        assert "Use the workbook tools (read_range" in partial.text

    def test_a_custom_counter_is_used(self, fixtures) -> None:
        calls: list[int] = []

        def counter(text: str) -> int:
            calls.append(len(text))
            return approx_tokens(text)

        render_workbook(fixtures.units_and_notes(), count_tokens=counter)
        assert calls

    def test_a_budget_too_small_for_anything_is_refused(self, fixtures) -> None:
        with pytest.raises(ValueError, match="50"):
            render_workbook(fixtures.units_and_notes(), token_budget=10)


def test_an_unreadable_upload_is_a_workbook_error() -> None:
    with pytest.raises(WorkbookError):
        render_workbook(b"not a zip", name="upload.xlsx")
