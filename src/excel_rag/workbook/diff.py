"""What changed between two versions of a workbook -- cells, formulas, what they read, and impact.

The index keeps only the active version of a workbook (older ones are garbage-collected), so a diff
is computed from the two files, in memory, by the same reader and model builder ingestion uses.
Nothing is evaluated: a changed input is reported with the formulas that read it, not with the
values those formulas would now compute, and a changed *saved* result is labelled as such.

Sheets are matched by name. A renamed sheet reads as one sheet removed and one added -- the
workbook carries no stable sheet identity to say otherwise.
"""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass, field
from enum import StrEnum
from pathlib import Path
from typing import Any

from ..models import A1Range
from .canonical import (
    CellValue,
    FormulaEntry,
    NamedRange,
    SheetModel,
    WorkbookModel,
    format_value,
)

#: The most cell changes a diff lists; the counts always cover every change.
DEFAULT_MAX_CHANGES = 500
#: The most formulas listed as reading one changed input.
MAX_IMPACT = 20


class ChangeKind(StrEnum):
    ADDED = "added"
    REMOVED = "removed"
    #: An input value changed.
    VALUE = "value"
    #: The formula text changed (its saved result may have changed with it).
    FORMULA = "formula"
    #: The formula is the same but its last-saved result differs: an upstream input moved.
    RESULT = "result"
    #: A value became a formula or the reverse.
    KIND = "kind"


@dataclass(frozen=True, slots=True)
class CellChange:
    sheet: str
    coordinate: str
    kind: ChangeKind
    before: str | None
    after: str | None
    before_formula: str | None = None
    after_formula: str | None = None
    #: For a changed formula: the ranges it reads now and no longer reads (sheet-qualified).
    reads_added: tuple[str, ...] = ()
    reads_removed: tuple[str, ...] = ()
    #: For a changed input: formulas in the new version that read it (sheet-qualified).
    read_by: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class NamedRangeChange:
    name: str
    before: str | None
    after: str | None


@dataclass(frozen=True, slots=True)
class WorkbookDiff:
    before_file: str
    after_file: str
    sheets_added: tuple[str, ...]
    sheets_removed: tuple[str, ...]
    changes: tuple[CellChange, ...]
    #: Every change, counted by sheet and kind -- including the ones past ``max_changes``.
    counts: dict[str, dict[str, int]] = field(default_factory=dict)
    named_ranges: tuple[NamedRangeChange, ...] = ()
    truncated: bool = False

    @property
    def total(self) -> int:
        return sum(sum(by_kind.values()) for by_kind in self.counts.values())

    @property
    def unchanged(self) -> bool:
        return not (self.total or self.sheets_added or self.sheets_removed or self.named_ranges)


def _qualified(sheet: str, a1: str) -> str:
    plain = sheet.replace("_", "").isalnum()
    return f"{sheet}!{a1}" if plain else f"'{sheet.replace(chr(39), chr(39) * 2)}'!{a1}"


def _shown(cell: CellValue | None) -> str | None:
    if cell is None:
        return None
    if cell.computed:
        return format_value(cell.cached_value) if cell.cached_value is not None else None
    return format_value(cell.value)


def _entry_for(sheet: SheetModel, coordinate: str) -> FormulaEntry | None:
    for entry in sheet.formulas:
        if coordinate in entry.member_coordinates:
            return entry
    return None


def _reads(entry: FormulaEntry | None) -> set[str]:
    if entry is None:
        return set()
    return {_qualified(reference.sheet_name, reference.a1_range) for reference in entry.references}


def _named(name: NamedRange) -> str:
    if name.resolved and name.sheet_name and name.a1:
        return _qualified(name.sheet_name, name.a1)
    return f"(unresolved: {name.detail or 'not a static range'})"


def _readers(model: WorkbookModel, sheet: str, coordinate: str) -> tuple[str, ...]:
    """Formulas in ``model`` whose references cover ``sheet!coordinate``."""
    target = A1Range.parse(sheet, coordinate)
    readers: list[str] = []
    for formula_sheet in model.sheets:
        for entry in formula_sheet.formulas:
            for reference in entry.references:
                if reference.sheet_name != sheet:
                    continue
                rows, columns = reference.row_span, reference.column_span
                if rows is None or columns is None:
                    continue
                if (
                    rows["gte"] <= target.min_row <= rows["lte"]
                    and columns["gte"] <= target.min_col <= columns["lte"]
                ):
                    readers.append(_qualified(formula_sheet.name, entry.a1_range.a1))
                    break
    return tuple(dict.fromkeys(readers))


def _same(left: Any, right: Any) -> bool:
    # 1 == 1.0 == True in Python; a cell that turned from TRUE into 1 has changed.
    return type(left) is type(right) and left == right


def _compare_cell(old: CellValue | None, new: CellValue | None) -> ChangeKind | None:
    if old is None and new is None:
        return None
    if old is None:
        return ChangeKind.ADDED
    if new is None:
        return ChangeKind.REMOVED
    if old.computed != new.computed:
        return ChangeKind.KIND
    if old.formula is not None or new.formula is not None:
        if old.formula != new.formula:
            return ChangeKind.FORMULA
        return None if _same(old.cached_value, new.cached_value) else ChangeKind.RESULT
    if old.array_master is not None or new.array_master is not None:
        # An array formula's extent: its master's change says it all.
        return None if _same(old.cached_value, new.cached_value) else ChangeKind.RESULT
    return None if _same(old.value, new.value) else ChangeKind.VALUE


def _row_major(coordinates: Iterable[str], cells: dict[str, CellValue]) -> list[str]:
    def key(coordinate: str) -> tuple[int, int]:
        cell = cells[coordinate]
        return cell.row, cell.column

    return sorted(coordinates, key=key)


def diff_workbooks(
    before: WorkbookModel,
    after: WorkbookModel,
    *,
    max_changes: int = DEFAULT_MAX_CHANGES,
) -> WorkbookDiff:
    """Compare two workbook models sheet by sheet, cell by cell. Never evaluates a formula."""
    old_sheets = {sheet.name: sheet for sheet in before.sheets if not sheet.is_macro_sheet}
    new_sheets = {sheet.name: sheet for sheet in after.sheets if not sheet.is_macro_sheet}
    changes: list[CellChange] = []
    counts: dict[str, dict[str, int]] = {}
    truncated = False

    for name in [sheet.name for sheet in after.sheets if sheet.name in old_sheets]:
        old_sheet, new_sheet = old_sheets[name], new_sheets.get(name)
        if new_sheet is None:
            continue
        old_cells = dict(old_sheet.cells)
        new_cells = dict(new_sheet.cells)
        merged = {**old_cells, **new_cells}
        for coordinate in _row_major(set(old_cells) | set(new_cells), merged):
            old, new = old_cells.get(coordinate), new_cells.get(coordinate)
            kind = _compare_cell(old, new)
            if kind is None:
                continue
            by_kind = counts.setdefault(name, {})
            by_kind[kind.value] = by_kind.get(kind.value, 0) + 1
            if len(changes) >= max_changes:
                truncated = True
                continue
            reads_added: tuple[str, ...] = ()
            reads_removed: tuple[str, ...] = ()
            read_by: tuple[str, ...] = ()
            if kind in (ChangeKind.FORMULA, ChangeKind.KIND, ChangeKind.ADDED):
                old_reads = _reads(_entry_for(old_sheet, coordinate))
                new_reads = _reads(_entry_for(new_sheet, coordinate))
                reads_added = tuple(sorted(new_reads - old_reads))
                reads_removed = tuple(sorted(old_reads - new_reads))
            if kind is ChangeKind.VALUE or (kind is ChangeKind.KIND and new and not new.computed):
                readers = _readers(after, name, coordinate)
                read_by = readers[:MAX_IMPACT]
            changes.append(
                CellChange(
                    sheet=name,
                    coordinate=coordinate,
                    kind=kind,
                    before=_shown(old),
                    after=_shown(new),
                    before_formula=old.formula if old else None,
                    after_formula=new.formula if new else None,
                    reads_added=reads_added,
                    reads_removed=reads_removed,
                    read_by=read_by,
                )
            )

    old_names = {name.label: name for name in before.named_ranges}
    new_names = {name.label: name for name in after.named_ranges}
    named_changes = [
        NamedRangeChange(
            label,
            _named(old_names[label]) if label in old_names else None,
            _named(new_names[label]) if label in new_names else None,
        )
        for label in sorted(set(old_names) | set(new_names))
        if label not in old_names
        or label not in new_names
        or _named(old_names[label]) != _named(new_names[label])
    ]

    return WorkbookDiff(
        before_file=before.source_file,
        after_file=after.source_file,
        sheets_added=tuple(sheet.name for sheet in after.sheets if sheet.name not in old_sheets),
        sheets_removed=tuple(sheet.name for sheet in before.sheets if sheet.name not in new_sheets),
        changes=tuple(changes),
        counts=counts,
        named_ranges=tuple(named_changes),
        truncated=truncated,
    )


def diff_files(
    before: str | Path | bytes,
    after: str | Path | bytes,
    *,
    max_changes: int = DEFAULT_MAX_CHANGES,
) -> WorkbookDiff:
    """Read two workbook files (paths or bytes) and diff them."""
    from . import load_workbook

    return diff_workbooks(load_workbook(before), load_workbook(after), max_changes=max_changes)


def _line(change: CellChange) -> str:
    where = _qualified(change.sheet, change.coordinate)
    if change.kind is ChangeKind.ADDED:
        text = f"{where}: added {change.after_formula or change.after!r}"
    elif change.kind is ChangeKind.REMOVED:
        text = f"{where}: removed (was {change.before_formula or change.before!r})"
    elif change.kind is ChangeKind.FORMULA:
        text = f"{where}: formula {change.before_formula} → {change.after_formula}"
        if change.before != change.after:
            text += f" (saved result {change.before!r} → {change.after!r})"
    elif change.kind is ChangeKind.RESULT:
        text = f"{where}: saved result {change.before!r} → {change.after!r} (formula unchanged)"
    elif change.kind is ChangeKind.KIND:
        before = change.before_formula or repr(change.before)
        after = change.after_formula or repr(change.after)
        text = f"{where}: {before} → {after}"
    else:
        text = f"{where}: {change.before!r} → {change.after!r}"
    if change.reads_added:
        text += f"; now reads {', '.join(change.reads_added)}"
    if change.reads_removed:
        text += f"; no longer reads {', '.join(change.reads_removed)}"
    if change.read_by:
        text += f"; read by {', '.join(change.read_by)}"
    return text


def format_diff(diff: WorkbookDiff) -> str:
    """The diff as markdown for a person or a context window."""
    lines = [f"# Changes from {diff.before_file} to {diff.after_file}"]
    if diff.unchanged:
        lines.append("No differences in cells, formulas, sheets or named ranges.")
        return "\n".join(lines)
    if diff.sheets_added:
        lines.append(f"Sheets added: {', '.join(diff.sheets_added)}.")
    if diff.sheets_removed:
        lines.append(f"Sheets removed: {', '.join(diff.sheets_removed)}.")
    if diff.counts:
        summary = "; ".join(
            f"{sheet}: " + ", ".join(f"{count} {kind}" for kind, count in sorted(by_kind.items()))
            for sheet, by_kind in diff.counts.items()
        )
        lines.append(f"{diff.total} cell change(s) — {summary}.")
    lines.append(
        "Saved results are the values Excel last saved; nothing is recalculated here, so a changed "
        "input's effect is shown as the formulas that read it."
    )
    for change in diff.changes:
        lines.append(f"- {_line(change)}")
    if diff.truncated:
        lines.append(f"({diff.total - len(diff.changes)} more change(s) not listed.)")
    if diff.named_ranges:
        lines.append("Named ranges:")
        lines.extend(
            f"- {change.name}: {change.before or '(absent)'} → {change.after or '(absent)'}"
            for change in diff.named_ranges
        )
    return "\n".join(lines)


__all__ = [
    "DEFAULT_MAX_CHANGES",
    "CellChange",
    "ChangeKind",
    "NamedRangeChange",
    "WorkbookDiff",
    "diff_files",
    "diff_workbooks",
    "format_diff",
]
