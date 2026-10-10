"""Tools a model calls to read what a rendering left out: ranges, matches, precedents, dependents.

:class:`WorkbookSession` answers them over the in-memory :class:`~excel_rag.workbook.WorkbookModel`
-- the same operations as the service's ``/excel/range`` and ``/excel/dependents`` endpoints, with
no index behind them. The tool definitions (:data:`TOOL_DEFINITIONS`) are plain dicts in the
Messages API shape (``name``, ``description``, ``input_schema``, ``strict``), so they work with
any client; :meth:`WorkbookSession.tool_result` turns a ``tool_use`` block into the matching
``tool_result`` block, with ``is_error`` set when the call could not be answered.

Every answer is bounded (:data:`MAX_RANGE_CELLS`, :data:`MAX_MATCHES`) and says what it cut, so a
tool call cannot flood the context window it exists to protect.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from ..models import A1Range
from ..workbook import WorkbookModel, load_workbook
from ..workbook.calc import (
    CalcInputError,
    CalcUnavailable,
    calculate,
    calculation_available,
    format_calculation,
)
from ..workbook.canonical import FormulaEntry, SheetModel, SheetObject, column_letter
from .render import (
    DETAIL_LADDER,
    FORMULA_MARK,
    _display,
    _qualified,
    formula_line,
    object_line,
)

#: The most cells one ``read_range`` returns; a larger range is cut by rows, and the reply says so.
MAX_RANGE_CELLS = 2_000
#: The most matches one ``find`` returns.
MAX_MATCHES = 50
#: The most formulas one ``precedents`` / ``dependents`` call lists.
MAX_FORMULAS = 100

_FULL = DETAIL_LADDER[0]
_SHEET = {
    "type": "string",
    "description": "Worksheet name exactly as shown in the workbook rendering, e.g. Forecast.",
}
_RANGE = {
    "type": "string",
    "description": "An A1 cell or rectangle on that sheet, e.g. C7 or A1:F40. No sheet prefix.",
}


def _schema(properties: dict[str, Any]) -> dict[str, Any]:
    return {
        "type": "object",
        "properties": properties,
        "required": list(properties),
        "additionalProperties": False,
    }


TOOL_DEFINITIONS: tuple[dict[str, Any], ...] = (
    {
        "name": "read_range",
        "description": (
            "Read the cells of a rectangle in the attached workbook, as a grid with row numbers "
            "and column letters. Formula cells show the value Excel last saved, marked "
            f"{FORMULA_MARK}, and the formulas in the range are listed with what they read. Use it "
            "for rows or columns the workbook rendering marked as omitted, or to check exact "
            f"values before citing them. Returns at most {MAX_RANGE_CELLS} cells; a larger range "
            "is cut by rows and the reply names the next rows to request."
        ),
        "input_schema": _schema({"sheet": _SHEET, "range": _RANGE}),
        "strict": True,
    },
    {
        "name": "find",
        "description": (
            "Search every sheet of the attached workbook for cells whose value or formula "
            "contains the text, case-insensitively, and return their coordinates and values. "
            "Use it to locate a label, a name or a number before reading around it with "
            f"read_range. Returns at most {MAX_MATCHES} matches and says how many more exist."
        ),
        "input_schema": _schema(
            {
                "query": {
                    "type": "string",
                    "description": "Text or number to look for, e.g. 'EBITDA' or '1240'.",
                }
            }
        ),
        "strict": True,
    },
    {
        "name": "precedents",
        "description": (
            "List the formulas inside a range of the attached workbook, each with its saved "
            "value and the cells and ranges it reads (its precedents), resolved statically - "
            "references such as INDIRECT or external links are reported as unresolved, never "
            "guessed; a chart, pivot table or data validation in the range is listed with the "
            "ranges it reads. Use it to explain how a value is calculated."
        ),
        "input_schema": _schema({"sheet": _SHEET, "range": _RANGE}),
        "strict": True,
    },
    {
        "name": "dependents",
        "description": (
            "List the formulas anywhere in the attached workbook that read any cell of a range "
            "(its direct dependents), with their saved values, and the charts, pivot tables and "
            "data validations that read it. Use it to answer what changes if an input changes; "
            "call it again on a dependent to follow the chain further."
        ),
        "input_schema": _schema({"sheet": _SHEET, "range": _RANGE}),
        "strict": True,
    },
)

#: The what-if tool, offered only when the ``calc`` extra (formula evaluation) is installed.
CALCULATE_DEFINITION: dict[str, Any] = {
    "name": "calculate",
    "description": (
        "Compute the cells of a range in the attached workbook, optionally after changing input "
        "cells (a what-if: 'what is net income if growth is 7%?'). Formulas no change reaches "
        "keep the value Excel saved (marked saved); those a change reaches are recomputed by "
        "excel-rag, not by Excel, and marked recalculated with Excel's saved value beside them. "
        "A value that cannot be computed - INDIRECT, OFFSET, a user-defined function, a circular "
        "reference - is reported unknown with the reason, never guessed. Changes are not kept: "
        "every call starts from the workbook as saved, so pass all the changes a scenario needs "
        f"at once. Computes at most {MAX_RANGE_CELLS} cells."
    ),
    "input_schema": _schema(
        {
            "sheet": _SHEET,
            "range": _RANGE,
            "changes": {
                "type": "array",
                "description": "Input cells to set before computing; empty for none.",
                "items": _schema(
                    {
                        "cell": {
                            "type": "string",
                            "description": "The cell with its sheet, e.g. Assumptions!B4.",
                        },
                        "value": {
                            "type": "string",
                            "description": (
                                "Its new value as typed into Excel: 0.07, 7%, 1200, TRUE, "
                                "text, or empty to clear it. Not a formula."
                            ),
                        },
                    }
                ),
            },
        }
    ),
    "strict": True,
}


@dataclass(frozen=True)
class ToolOutcome:
    """A tool's answer, and whether it is an error the model should correct and retry."""

    content: str
    is_error: bool = False


class ToolInputError(ValueError):
    """The call named a sheet, range or argument the workbook cannot answer."""


class WorkbookSession:
    """One attached workbook, answering tool calls for the length of a conversation."""

    def __init__(self, model: WorkbookModel) -> None:
        self.model = model
        self._sheets = {sheet.name: sheet for sheet in model.sheets}

    @classmethod
    def load(cls, source: str | Path | bytes, *, name: str | None = None) -> WorkbookSession:
        return cls(load_workbook(source, name=name))

    # -- tool plumbing ----------------------------------------------------------------------
    @staticmethod
    def tool_definitions(*, calculate: bool | None = None) -> list[dict[str, Any]]:
        """The tool definitions to pass as ``tools`` in a Messages API request.

        ``calculate`` (formula evaluation) is included when the ``calc`` extra is installed,
        unless ``calculate=False``.
        """
        definitions = [dict(definition) for definition in TOOL_DEFINITIONS]
        if calculate if calculate is not None else calculation_available():
            definitions.append(dict(CALCULATE_DEFINITION))
        return definitions

    def call(self, name: str, tool_input: Any) -> ToolOutcome:
        """Run one tool call. A bad call is an ``is_error`` outcome, never an exception."""
        try:
            arguments = _arguments(name, tool_input)
            if name == "read_range":
                return ToolOutcome(self.read_range(arguments["sheet"], arguments["range"]))
            if name == "find":
                return ToolOutcome(self.find(arguments["query"]))
            if name == "precedents":
                return ToolOutcome(self.precedents(arguments["sheet"], arguments["range"]))
            if name == "calculate":
                changes = _changes(tool_input.get("changes"))
                return ToolOutcome(self.calculate(arguments["sheet"], arguments["range"], changes))
            return ToolOutcome(self.dependents(arguments["sheet"], arguments["range"]))
        except ToolInputError as error:
            return ToolOutcome(f"Error: {error}", is_error=True)

    def tool_result(self, tool_use_id: str, name: str, tool_input: Any) -> dict[str, Any]:
        """The ``tool_result`` content block answering one ``tool_use`` block."""
        outcome = self.call(name, tool_input)
        block: dict[str, Any] = {
            "type": "tool_result",
            "tool_use_id": tool_use_id,
            "content": outcome.content,
        }
        if outcome.is_error:
            block["is_error"] = True
        return block

    # -- the tools --------------------------------------------------------------------------
    def read_range(self, sheet_name: str, a1: str) -> str:
        sheet = self._sheet(sheet_name)
        region = _parse(sheet.name, a1)
        width = region.max_col - region.min_col + 1
        rows = [
            row
            for row in range(region.min_row, region.max_row + 1)
            if any(
                f"{column_letter(column)}{row}" in sheet.cells
                for column in range(region.min_col, region.max_col + 1)
            )
        ]
        where = _qualified(sheet.name, region.a1)
        if not rows:
            return f"{where} is empty."
        if width > MAX_RANGE_CELLS:
            raise ToolInputError(
                f"{region.a1} is {width} columns wide; request at most {MAX_RANGE_CELLS} columns"
            )
        limit = max(1, MAX_RANGE_CELLS // width)
        shown, rest = rows[:limit], rows[limit:]
        columns = range(region.min_col, region.max_col + 1)
        lines = [
            f"{where}:",
            "| | " + " | ".join(column_letter(column) for column in columns) + " |",
            "|---|" + "---|" * width,
        ]
        for row in shown:
            cells = [
                _display(sheet.cells.get(f"{column_letter(column)}{row}"), _FULL)
                for column in columns
            ]
            lines.append(f"| {row} | " + " | ".join(cells) + " |")
        skipped = (shown[-1] - region.min_row + 1) - len(shown)
        if skipped:
            lines.append(f"({skipped} empty row(s) not shown.)")
        if rest:
            next_range = (
                f"{column_letter(region.min_col)}{rest[0]}:"
                f"{column_letter(region.max_col)}{region.max_row}"
            )
            lines.append(
                f"(Cut at {MAX_RANGE_CELLS} cells: rows {rest[0]}-{rest[-1]} not shown; "
                f"read_range {next_range} for the rest.)"
            )
        formulas = [
            entry
            for entry in sheet.formulas
            if entry.a1_range.intersects(region)
            and (entry.a1_range.min_row <= shown[-1] or entry.is_cluster)
        ]
        if formulas:
            lines.append("Formulas:")
            lines.extend(
                f"- {formula_line(entry, qualified=True)}" for entry in formulas[:MAX_FORMULAS]
            )
            if len(formulas) > MAX_FORMULAS:
                lines.append(f"({len(formulas) - MAX_FORMULAS} more formula(s) not shown.)")
        return "\n".join(lines)

    def find(self, query: str) -> str:
        needle = query.strip().lower()
        if not needle:
            raise ToolInputError("query is empty")
        matches: list[str] = []
        total = 0
        for sheet in self.model.sheets:
            for cell in sorted(sheet.cells.values(), key=lambda item: (item.row, item.column)):
                shown = _display(cell, _FULL)
                haystack = f"{shown} {cell.formula or ''}".lower()
                if needle in haystack:
                    total += 1
                    if len(matches) < MAX_MATCHES:
                        matches.append(f"- {_qualified(sheet.name, cell.coordinate)}: {shown}")
        for named in self.model.named_ranges:
            if needle in named.name.lower() and named.sheet_name and named.a1:
                total += 1
                if len(matches) < MAX_MATCHES:
                    matches.append(
                        f"- named range {named.label} = {_qualified(named.sheet_name, named.a1)}"
                    )
        if not matches:
            return f"No cell contains {query!r}."
        header = f"{total} match(es) for {query!r}:"
        footer = (
            [f"({total - len(matches)} more not shown; narrow the query.)"]
            if total > len(matches)
            else []
        )
        return "\n".join([header, *matches, *footer])

    def precedents(self, sheet_name: str, a1: str) -> str:
        sheet = self._sheet(sheet_name)
        region = _parse(sheet.name, a1)
        where = _qualified(sheet.name, region.a1)
        entries = [entry for entry in sheet.formulas if entry.a1_range.intersects(region)]
        objects = [item for item in sheet.objects if item.anchor.intersects(region)]
        if not entries and not objects:
            return f"{where} holds no formulas: its values are inputs, not calculations."
        lines = [f"Formulas in {where} and what they read:"] if entries else []
        for entry in entries[:MAX_FORMULAS]:
            lines.append(f"- {formula_line(entry, qualified=True)}")
            for reference in dict.fromkeys(
                (reference.sheet_name, reference.a1_range) for reference in entry.references
            ):
                lines.append(f"    - {self._preview(*reference)}")
        if len(entries) > MAX_FORMULAS:
            lines.append(f"({len(entries) - MAX_FORMULAS} more formula(s) not shown.)")
        if objects:
            lines.append(f"Charts, pivot tables and validations at {where} and what they read:")
            lines.extend(
                f"- {object_line(item, qualified=True)}" for item in objects[:MAX_FORMULAS]
            )
            if len(objects) > MAX_FORMULAS:
                lines.append(f"({len(objects) - MAX_FORMULAS} more not shown.)")
        return "\n".join(lines)

    def dependents(self, sheet_name: str, a1: str) -> str:
        sheet = self._sheet(sheet_name)
        region = _parse(sheet.name, a1)
        where = _qualified(sheet.name, region.a1)
        found: list[FormulaEntry] = [
            entry
            for other in self.model.sheets
            for entry in other.formulas
            if any(_reads(reference, region) for reference in entry.references)
        ]
        readers: list[SheetObject] = [
            item
            for other in self.model.sheets
            for item in other.objects
            if any(_reads(reference, region) for reference in item.references)
        ]
        if not found and not readers:
            return f"No formula, chart, pivot table or validation reads {where}."
        lines: list[str] = []
        if found:
            lines.append(f"Formulas that read {where}:")
            lines.extend(
                f"- {formula_line(entry, qualified=True)}" for entry in found[:MAX_FORMULAS]
            )
            if len(found) > MAX_FORMULAS:
                lines.append(f"({len(found) - MAX_FORMULAS} more formula(s) not shown.)")
        if readers:
            lines.append(f"Charts, pivot tables and validations that read {where}:")
            lines.extend(
                f"- {object_line(item, qualified=True)}" for item in readers[:MAX_FORMULAS]
            )
            if len(readers) > MAX_FORMULAS:
                lines.append(f"({len(readers) - MAX_FORMULAS} more not shown.)")
        return "\n".join(lines)

    def calculate(self, sheet_name: str, a1: str, changes: Mapping[str, str] | None = None) -> str:
        """The range's values, computed under ``changes`` (cell -> value as typed)."""
        sheet = self._sheet(sheet_name)
        region = _parse(sheet.name, a1)
        if region.cell_count > MAX_RANGE_CELLS:
            raise ToolInputError(
                f"{region.cell_count} cells is more than {MAX_RANGE_CELLS}; ask for a smaller range"
            )
        try:
            calculation = calculate(self.model, _qualified(sheet.name, region.a1), changes or {})
        except (CalcInputError, CalcUnavailable) as error:
            raise ToolInputError(str(error)) from error
        return format_calculation(calculation)

    # -- helpers ----------------------------------------------------------------------------
    def _sheet(self, name: str) -> SheetModel:
        sheet = self._sheets.get(name)
        if sheet is None:
            folded = {key.lower(): value for key, value in self._sheets.items()}
            sheet = folded.get(name.strip().lower())
        if sheet is None:
            names = ", ".join(repr(key) for key in self._sheets)
            raise ToolInputError(f"no sheet named {name!r}; the sheets are {names}")
        if sheet.is_macro_sheet:
            raise ToolInputError(f"{sheet.name!r} is a macro sheet; its cells are not loaded")
        return sheet

    def _preview(self, sheet_name: str, a1: str) -> str:
        """A precedent with its value when it is one cell, or its size when it is a range."""
        where = _qualified(sheet_name, a1)
        sheet = self._sheets.get(sheet_name)
        try:
            bounds = A1Range.parse(sheet_name, a1)
        except ValueError:
            return where
        if sheet is not None and bounds.cell_count == 1:
            cell = sheet.cells.get(bounds.a1)
            return f"{where} = {_display(cell, _FULL) or '(empty)'}"
        return f"{where} ({bounds.cell_count} cells)"


def _arguments(name: str, tool_input: Any) -> dict[str, str]:
    definitions = (*TOOL_DEFINITIONS, CALCULATE_DEFINITION)
    definition = next((item for item in definitions if item["name"] == name), None)
    if definition is None:
        known = ", ".join(item["name"] for item in definitions)
        raise ToolInputError(f"unknown tool {name!r}; the tools are {known}")
    if not isinstance(tool_input, Mapping):
        raise ToolInputError("the tool input must be a JSON object")
    required: Sequence[str] = definition["input_schema"]["required"]
    arguments: dict[str, str] = {}
    for key in required:
        if key == "changes":
            continue  # a list, checked by _changes
        value = tool_input.get(key)
        if not isinstance(value, str) or not value.strip():
            raise ToolInputError(f"{name} needs a non-empty string {key!r}")
        arguments[key] = value.strip()
    extra = sorted(set(tool_input) - set(required))
    if extra:
        raise ToolInputError(f"{name} takes no {', '.join(map(repr, extra))}")
    return arguments


def _parse(sheet_name: str, a1: str) -> A1Range:
    text = a1.split("!", 1)[1] if "!" in a1 else a1
    try:
        return A1Range.parse(sheet_name, text.replace("$", ""))
    except ValueError as error:
        raise ToolInputError(
            f"{a1!r} is not an A1 cell or rectangle such as C7 or A1:F40"
        ) from error


def _reads(reference: Any, region: A1Range) -> bool:
    if reference.sheet_name != region.sheet_name or not reference.resolved:
        return False
    rows, columns = reference.row_span, reference.column_span
    if rows is None or columns is None:
        return False
    return not (
        rows["lte"] < region.min_row
        or rows["gte"] > region.max_row
        or columns["lte"] < region.min_col
        or columns["gte"] > region.max_col
    )


def _changes(raw: Any) -> dict[str, str]:
    """The ``changes`` argument as cell -> value, or a correctable error."""
    if not isinstance(raw, list):
        raise ToolInputError("calculate needs 'changes', a list of {cell, value} (empty for none)")
    changes: dict[str, str] = {}
    for item in raw:
        if not isinstance(item, Mapping) or set(item) != {"cell", "value"}:
            raise ToolInputError("each change is an object with exactly 'cell' and 'value'")
        cell, value = item["cell"], item["value"]
        if not isinstance(cell, str) or not cell.strip() or not isinstance(value, str):
            raise ToolInputError("a change's 'cell' and 'value' are strings, the cell non-empty")
        changes[cell.strip()] = value
    return changes


__all__ = [
    "CALCULATE_DEFINITION",
    "MAX_FORMULAS",
    "MAX_MATCHES",
    "MAX_RANGE_CELLS",
    "TOOL_DEFINITIONS",
    "ToolInputError",
    "ToolOutcome",
    "WorkbookSession",
]
