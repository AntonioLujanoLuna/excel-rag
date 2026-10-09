"""Formula reference extraction: typed edges and explicit gaps, never an invented edge."""

from __future__ import annotations

from excel_rag.models import A1Range, ReferenceKind, UnresolvedReason
from excel_rag.workbook.formulas import (
    FormulaContext,
    NamedRangeInfo,
    TableInfo,
    parse_formula,
)


def _context(**overrides) -> FormulaContext:
    base = dict(
        workbook_id="wb",
        version=1,
        sheet_name="Calc",
        known_sheets=frozenset({"Calc", "Actuals", "Assumptions"}),
        macro_sheets=frozenset({"Macro1"}),
        sheet_max_row={"Calc": 100, "Actuals": 500, "Assumptions": 50},
        tables={},
        defined_names={},
    )
    base.update(overrides)
    return FormulaContext(**base)


def test_cross_sheet_range_and_cell_are_single_edges() -> None:
    parsed = parse_formula("=SUM(Actuals!D2:D500)*(1+Assumptions!C7)", _context())
    edges = {
        (edge.reference.sheet_name, edge.reference.a1_range, edge.reference.kind)
        for edge in parsed.references
    }
    assert ("Actuals", "D2:D500", ReferenceKind.RANGE) in edges
    assert ("Assumptions", "C7", ReferenceKind.CELL) in edges
    assert parsed.unresolved == ()
    # one range edge, not 499 cell edges
    assert sum(1 for edge in parsed.references if edge.reference.a1_range == "D2:D500") == 1


def test_same_sheet_and_absolute_references() -> None:
    parsed = parse_formula("=A1+$B$2*C3:C5", _context())
    keys = {(edge.reference.a1_range, edge.reference.kind) for edge in parsed.references}
    assert ("A1", ReferenceKind.CELL) in keys
    assert ("B2", ReferenceKind.CELL) in keys
    assert ("C3:C5", ReferenceKind.RANGE) in keys
    absolute = {edge.reference.a1_range: edge.absolute for edge in parsed.references}
    assert absolute["B2"] is True
    assert absolute["A1"] is False


def test_indirect_and_offset_are_explicit_gaps() -> None:
    indirect = parse_formula('=INDIRECT("A2")', _context())
    assert [gap.reason for gap in indirect.unresolved] == [UnresolvedReason.INDIRECT]
    offset = parse_formula("=OFFSET(A3,1,0)", _context())
    assert UnresolvedReason.VOLATILE_OFFSET in {gap.reason for gap in offset.unresolved}
    # OFFSET still resolves the explicit A3 it anchors on
    assert any(edge.reference.a1_range == "A3" for edge in offset.references)


def test_named_range_reference() -> None:
    context = _context(
        defined_names={
            "growthrate": NamedRangeInfo(
                "GrowthRate", "Assumptions", A1Range.parse("Assumptions", "C7")
            )
        }
    )
    parsed = parse_formula("=100*GrowthRate", context)
    kinds = {edge.reference.kind for edge in parsed.references}
    assert ReferenceKind.NAMED_RANGE in kinds
    named = next(
        edge.reference
        for edge in parsed.references
        if edge.reference.kind is ReferenceKind.NAMED_RANGE
    )
    assert named.target_node_id == "wb:v1:named_range:growthrate"


def test_macro_sheet_reference_is_a_gap() -> None:
    parsed = parse_formula("=Macro1!A1", _context())
    assert [gap.reason for gap in parsed.unresolved] == [UnresolvedReason.MACRO_SHEET]


def test_missing_sheet_is_out_of_range() -> None:
    parsed = parse_formula("=Missing!A1", _context())
    assert [gap.reason for gap in parsed.unresolved] == [UnresolvedReason.OUT_OF_RANGE]


def test_external_link_is_a_gap_and_its_sheet_is_not_parsed() -> None:
    parsed = parse_formula("='[Budget.xlsx]Sheet1'!A1", _context())
    assert UnresolvedReason.EXTERNAL_LINK in {gap.reason for gap in parsed.unresolved}
    assert all(edge.reference.sheet_name != "Budget.xlsx" for edge in parsed.references)


def test_dynamic_forms_are_gaps() -> None:
    let = parse_formula("=LET(x,1,x+1)", _context())
    assert UnresolvedReason.DYNAMIC_ARRAY in {gap.reason for gap in let.unresolved}
    spill = parse_formula("=A1#", _context())
    assert UnresolvedReason.DYNAMIC_ARRAY in {gap.reason for gap in spill.unresolved}
    implicit = parse_formula("=@A1", _context())
    assert UnresolvedReason.DYNAMIC_ARRAY in {gap.reason for gap in implicit.unresolved}


def test_unsupported_function_is_a_gap() -> None:
    parsed = parse_formula("=_xlfn.CONCAT(A1,A2)", _context())
    assert UnresolvedReason.UNSUPPORTED_FUNCTION in {gap.reason for gap in parsed.unresolved}


def test_table_column_reference() -> None:
    context = _context(
        tables={
            "SalesTable": TableInfo(
                "SalesTable", "Actuals", A1Range.parse("Actuals", "A1:B9"), ("Region", "Revenue")
            )
        }
    )
    parsed = parse_formula("=SUM(SalesTable[Revenue])", context)
    assert len(parsed.references) == 1
    reference = parsed.references[0].reference
    assert reference.kind is ReferenceKind.TABLE_COLUMN
    assert reference.sheet_name == "Actuals"


def test_whole_column_reference_expands_to_the_used_height() -> None:
    parsed = parse_formula("=SUM(A:B)", _context())
    assert any(edge.reference.a1_range == "A1:B100" for edge in parsed.references)


def test_unknown_table_reference_is_a_gap() -> None:
    parsed = parse_formula("=SUM(Nope[Amount])", _context())
    assert UnresolvedReason.OUT_OF_RANGE in {gap.reason for gap in parsed.unresolved}


def test_string_literals_are_not_parsed_as_references() -> None:
    parsed = parse_formula('=IF(A1="B2", "C3", D4)', _context())
    found = {edge.reference.a1_range for edge in parsed.references}
    assert found == {"A1", "D4"}


def test_references_are_deduplicated() -> None:
    parsed = parse_formula("=A1+A1+A1", _context())
    assert len(parsed.references) == 1


def test_a_function_name_that_looks_like_a_cell_is_not_an_edge() -> None:
    """``LOG10``, ``ATAN2``, ``DAYS360`` spell valid A1 cells; followed by ``(`` they are calls."""
    parsed = parse_formula("=LOG10(A1)+ATAN2(B1, C1)+DAYS360(D1, E1)", _context())
    targets = {edge.reference.a1_range for edge in parsed.references}
    assert targets == {"A1", "B1", "C1", "D1", "E1"}
    assert parsed.unresolved == ()


def test_a_quoted_sheet_name_with_an_escaped_apostrophe_resolves() -> None:
    context = _context(
        known_sheets=frozenset({"Calc", "Bob's Inputs"}),
        sheet_max_row={"Calc": 100, "Bob's Inputs": 10},
    )
    parsed = parse_formula("='Bob''s Inputs'!B2*2", context)
    assert [(edge.reference.sheet_name, edge.reference.a1_range) for edge in parsed.references] == [
        ("Bob's Inputs", "B2")
    ]
    assert parsed.unresolved == ()


def _ordered_context() -> FormulaContext:
    return _context(
        known_sheets=frozenset({"Calc", "Jan", "Feb", "Mar", "Apr"}),
        sheet_order=("Calc", "Jan", "Feb", "Mar", "Apr"),
        sheet_max_row={"Calc": 100, "Jan": 40, "Feb": 40, "Mar": 40, "Apr": 40},
        sheet_max_col={"Calc": 8, "Jan": 5},
    )


def test_a_three_d_reference_is_one_edge_per_sheet_in_the_span() -> None:
    parsed = parse_formula("=SUM(Jan:Mar!B2)+SUM('Feb:Apr'!C3:C4)", _ordered_context())
    found = {(edge.reference.sheet_name, edge.reference.a1_range) for edge in parsed.references}
    assert found == {
        ("Jan", "B2"),
        ("Feb", "B2"),
        ("Mar", "B2"),
        ("Feb", "C3:C4"),
        ("Mar", "C3:C4"),
        ("Apr", "C3:C4"),
    }
    assert parsed.unresolved == ()


def test_a_three_d_reference_without_a_sheet_order_is_a_gap() -> None:
    parsed = parse_formula("=SUM(Actuals:Assumptions!B2)", _context())
    assert parsed.references == ()
    assert {gap.reason for gap in parsed.unresolved} == {UnresolvedReason.OUT_OF_RANGE}


def test_a_whole_row_reference_spans_the_used_width() -> None:
    parsed = parse_formula("=SUM(2:3)+SUM(Jan!$5:$5)", _ordered_context())
    found = {(edge.reference.sheet_name, edge.reference.a1_range) for edge in parsed.references}
    assert found == {("Calc", "A2:H3"), ("Jan", "A5:E5")}
    assert parsed.unresolved == ()


def test_a_range_naming_its_sheet_twice_is_one_rectangle() -> None:
    parsed = parse_formula("=SUM(Actuals!A1:Actuals!B2)", _context())
    assert [(edge.reference.sheet_name, edge.reference.a1_range) for edge in parsed.references] == [
        ("Actuals", "A1:B2")
    ]


def test_a_broken_reference_is_a_gap() -> None:
    parsed = parse_formula("=#REF!+Actuals!#REF!+A1", _context())
    assert [edge.reference.a1_range for edge in parsed.references] == ["A1"]
    assert [gap.reason for gap in parsed.unresolved] == [UnresolvedReason.MALFORMED] * 2


def test_a_range_bounded_by_a_function_is_a_gap_and_its_arguments_are_edges() -> None:
    parsed = parse_formula("=SUM(A1:INDEX(B:B,5))", _context())
    assert {edge.reference.a1_range for edge in parsed.references} == {"B1:B100"}
    assert UnresolvedReason.UNSUPPORTED_FUNCTION in {gap.reason for gap in parsed.unresolved}


def test_an_undefined_name_is_a_gap_but_a_let_variable_is_not() -> None:
    undefined = parse_formula("=Rate*B2", _context())
    assert [edge.reference.a1_range for edge in undefined.references] == ["B2"]
    assert [gap.reference_text for gap in undefined.unresolved] == ["Rate"]
    let = parse_formula("=LET(rate, B2, rate*2)", _context())
    assert [edge.reference.a1_range for edge in let.references] == ["B2"]
    assert {gap.reason for gap in let.unresolved} == {UnresolvedReason.DYNAMIC_ARRAY}


def test_an_untokenisable_formula_is_one_gap_and_no_edges() -> None:
    parsed = parse_formula('=SUM(A1,"unterminated)', _context())
    assert parsed.references == ()
    assert [gap.reason for gap in parsed.unresolved] == [UnresolvedReason.MALFORMED]
