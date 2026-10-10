"""Evaluation cases mined from a workbook's own labels: a question per labelled formula and input.

Writing cases by hand is the slow part of measuring retrieval. A workbook already says what most of
its cells are: a formula sits under a column header (``EBITDA``) or beside a row label (``Q1``), and
an input a formula reads sits beside its label (``Tax rate``). Each becomes a question whose answer
is that cell or range:

* a formula, or a column of identical formulas: ``How is EBITDA calculated?`` (header and row
  label together when both exist: ``How is Q1 Revenue calculated?``);
* an input cell that some formula reads: ``What is the Tax rate?``, answered by the label and
  the value together (``A5:B5``).

Mined questions reuse the workbook's words, so they are easy for lexical search: they make a
regression set for chunking and ranking changes (did a chunk stop carrying its header?), not a
measure of semantic recall. Hand-written cases with no word in common with the cells remain the
way to measure that. The cases are deterministic, so a mined file can be saved and reviewed.
"""

from __future__ import annotations

import re
from typing import Any

from ..models import A1Range
from ..workbook.canonical import CellValue, SheetModel, WorkbookModel, column_letter
from .cases import EvalCase, Expected

#: Labels longer than this are prose, not names, and make poor questions.
MAX_LABEL_LENGTH = 60


def mine_cases(
    model: WorkbookModel, *, workbook_id: str | None = None, limit: int = 100
) -> tuple[EvalCase, ...]:
    """Questions over ``model``'s labelled formulas and inputs, at most ``limit``, sheet order."""
    target = workbook_id or model.workbook_id
    cases: list[EvalCase] = []
    seen: set[str] = set()

    def add(sheet: SheetModel, a1: str, question: str) -> None:
        key = question.lower()
        if key in seen:
            return  # the same words over two places: neither is the answer to it alone
        seen.add(key)
        cases.append(
            EvalCase(
                id=f"mined:{sheet.name}!{a1}",
                question=question,
                workbook_id=target,
                expected=(Expected(sheet.name, a1),),
            )
        )

    read_cells = _cells_read_by_formulas(model)
    for sheet in model.sheets:
        if sheet.is_macro_sheet:
            continue
        for entry in sheet.formulas:
            first = sheet.cells.get(entry.member_coordinates[0])
            if first is None:
                continue
            header = _label_above(sheet, first)
            row_label = None if entry.is_cluster else _label_left(sheet, first)[0]
            subject = " ".join(part for part in (row_label, header) if part)
            if subject:
                add(sheet, entry.a1_range.a1, f"How is {subject} calculated?")
        for cell in sorted(sheet.cells.values(), key=lambda cell: (cell.row, cell.column)):
            if (sheet.name, cell.row, cell.column) not in read_cells or cell.computed:
                continue
            if not isinstance(cell.value, (int, float)) or isinstance(cell.value, bool):
                continue
            label, label_column = _label_left(sheet, cell)
            if label:
                # The label is part of the citation: a hit on the label column beside the value
                # answers "what is the tax rate?" as well as one on the value itself.
                span = f"{column_letter(label_column)}{cell.row}:{cell.coordinate}"
                add(sheet, span, f"What is the {label}?")
    return tuple(cases[:limit])


def _cells_read_by_formulas(model: WorkbookModel) -> set[tuple[str, int, int]]:
    """Every single cell some formula names on its own (an input, not a column it sums)."""
    read: set[tuple[str, int, int]] = set()
    for sheet in model.sheets:
        for entry in sheet.formulas:
            for reference in entry.references:
                if ":" in reference.a1_range:
                    continue
                try:
                    cell = A1Range.parse(reference.sheet_name, reference.a1_range)
                except ValueError:
                    continue
                read.add((reference.sheet_name, cell.min_row, cell.min_col))
    return read


def _label(value: Any) -> str | None:
    if not isinstance(value, str):
        return None
    text = " ".join(value.split()).rstrip(":").strip()
    if len(text) < 2 or len(text) > MAX_LABEL_LENGTH or not re.search(r"[^\W\d_]", text):
        return None
    return text


def _label_left(sheet: SheetModel, cell: CellValue) -> tuple[str | None, int]:
    """The nearest text label to the left on the same row, past the formulas beside it, and its
    column."""
    for column in range(cell.column - 1, 0, -1):
        neighbour = _cell_at(sheet, cell.row, column)
        if neighbour is None or neighbour.computed:
            continue
        return _label(neighbour.value), column
    return None, cell.column


def _label_above(sheet: SheetModel, cell: CellValue) -> str | None:
    """The nearest text label above in the same column, past the formulas stacked over it."""
    for row in range(cell.row - 1, 0, -1):
        neighbour = _cell_at(sheet, row, cell.column)
        if neighbour is None:
            return None
        if neighbour.computed:
            continue
        return _label(neighbour.value)
    return None


def _cell_at(sheet: SheetModel, row: int, column: int) -> CellValue | None:
    cell = sheet.cells.get(f"{column_letter(column)}{row}")
    if cell is None or (cell.value is None and not cell.computed):
        return None
    return cell


__all__ = ["MAX_LABEL_LENGTH", "mine_cases"]
