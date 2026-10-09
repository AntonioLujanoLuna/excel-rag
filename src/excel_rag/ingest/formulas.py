"""Static formula reference extraction -- never evaluation.

A formula is turned into typed :class:`~excel_rag.models.Reference` edges by text analysis alone.
Two inviolable rules from the design:

* **Never invent an edge.** A precedent that cannot be determined statically (``INDIRECT``, volatile
  ``OFFSET``, an external workbook, an unsupported dynamic array) becomes an
  :class:`~excel_rag.models.UnresolvedReference` carrying the reason, not a guessed target.
* **Edges point at ranges, not cells.** ``SUM(Actuals!D2:D500)`` is *one* range edge; the rectangle
  is indexed once and overlaps are answered by an ``integer_range`` query at retrieval time.

The declared node ids come from :mod:`excel_rag.ingest.canonical`, so an edge target and the
structure node it points at are produced by the same convention.
"""

from __future__ import annotations

import re
from collections.abc import Mapping
from dataclasses import dataclass

from ..models import A1Range, Reference, ReferenceKind, UnresolvedReason, UnresolvedReference
from .canonical import (
    NamedRange,
    cell_node_id,
    column_letter,
    named_range_node_id,
    range_node_id,
    table_column_node_id,
)

_STRING_RE = re.compile(r'"(?:[^"]|"")*"')
_EXTERNAL_RE = re.compile(
    r"'[^']*\[[^\]]+\][^']*'|\[[^\]]+\.(?:xlsx?|xlsm|xlsb|csv|xls)\]|\[\d+\]", re.IGNORECASE
)
_FUNC_RE = re.compile(r"(?<![A-Za-z0-9_.])([A-Za-z_][A-Za-z0-9_.]*)\s*\(")
_TABLE_REF_RE = re.compile(r"(?<![A-Za-z0-9_.])([A-Za-z_][A-Za-z0-9_]*)\s*\[([^\]]*)\]")
_COL_RE = re.compile(
    r"(?:(?:'(?P<qsheet>[^']+)'|(?P<sheet>[A-Za-z_][A-Za-z0-9_.]*))!)?"
    r"(?<![A-Za-z0-9_$])\$?(?P<start>[A-Za-z]{1,3}):\$?(?P<end>[A-Za-z]{1,3})(?![A-Za-z0-9_])"
)
_CELL_RE = re.compile(
    r"(?:(?:'(?P<qsheet>[^']+)'|(?P<sheet>[A-Za-z_][A-Za-z0-9_.]*))!)?"
    r"(?<![A-Za-z0-9_$])(?P<start>\$?[A-Za-z]{1,3}\$?\d{1,7})"
    r"(?::(?P<end>\$?[A-Za-z]{1,3}\$?\d{1,7}))?(?![A-Za-z0-9_])"
)
_IDENT_RE = re.compile(r"(?<![A-Za-z0-9_.$])(?P<name>[A-Za-z_][A-Za-z0-9_.]*)(?![A-Za-z0-9_(])")
_SPILL_RE = re.compile(r"(?:[)\]]|\d)\s*#")
_AT_RE = re.compile(r"(?<![\w.])@")

#: Functions whose result is a dynamic array: they cannot be flattened to static edges.
_DYNAMIC_FUNCTIONS = frozenset(
    {
        "LET",
        "LAMBDA",
        "SEQUENCE",
        "FILTER",
        "SORT",
        "SORTBY",
        "UNIQUE",
        "RANDARRAY",
        "XMATCH",
        "TAKE",
        "DROP",
        "HSTACK",
        "VSTACK",
        "TOROW",
        "TOCOL",
        "WRAPROWS",
        "WRAPCOLS",
        "EXPAND",
        "CHOOSECOLS",
        "CHOOSEROWS",
        "MAP",
        "REDUCE",
        "SCAN",
        "BYROW",
        "BYCOL",
        "MAKEARRAY",
        "TEXTSPLIT",
        "TEXTBEFORE",
        "TEXTAFTER",
        "GROUPBY",
        "PIVOTBY",
        "ANCHORARRAY",
        "SINGLE",
    }
)


@dataclass(frozen=True, slots=True)
class TableInfo:
    """An Excel table resolved to a sheet and rectangle, with its column names."""

    name: str
    sheet_name: str
    a1_range: A1Range
    columns: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class NamedRangeInfo:
    """A defined name resolved to its target, or ``None`` when it cannot be statically resolved."""

    name: str
    sheet_name: str
    a1_range: A1Range


@dataclass(frozen=True, slots=True)
class FormulaContext:
    """Everything needed to resolve a formula's precedents, and nothing that needs evaluation."""

    workbook_id: str
    version: int
    sheet_name: str
    known_sheets: frozenset[str]
    macro_sheets: frozenset[str]
    sheet_max_row: Mapping[str, int]
    tables: Mapping[str, TableInfo]
    defined_names: Mapping[str, NamedRangeInfo]


@dataclass(frozen=True, slots=True)
class ReferenceEdge:
    """A resolved reference plus the geometry a cluster expansion needs."""

    reference: Reference
    absolute: bool
    min_row: int
    max_row: int
    min_col: int
    max_col: int


@dataclass(frozen=True, slots=True)
class ParsedFormula:
    references: tuple[ReferenceEdge, ...]
    unresolved: tuple[UnresolvedReference, ...]


def _blank(text: str, spans: list[tuple[int, int]]) -> str:
    chars = list(text)
    for start, end in spans:
        for position in range(start, min(end, len(chars))):
            chars[position] = " "
    return "".join(chars)


def _helper_spans(text: str) -> tuple[list[tuple[int, int]], list[UnresolvedReference]]:
    """Spans of string literals, external links and table refs, and the gaps they imply."""
    spans: list[tuple[int, int]] = []
    unresolved: list[UnresolvedReference] = []
    for match in _STRING_RE.finditer(text):
        spans.append(match.span())
    for match in _EXTERNAL_RE.finditer(text):
        spans.append(match.span())
        unresolved.append(
            UnresolvedReference(
                reference_text=match.group(0).strip(),
                reason=UnresolvedReason.EXTERNAL_LINK,
                detail="external workbook or link index is not resolved during ingestion",
            )
        )
    return spans, unresolved


def _scan_functions(text: str) -> list[UnresolvedReference]:
    found: list[UnresolvedReference] = []
    if re.search(r"\bINDIRECT\s*\(", text, re.IGNORECASE):
        found.append(
            UnresolvedReference(
                reference_text="INDIRECT(...)",
                reason=UnresolvedReason.INDIRECT,
                detail="INDIRECT target depends on a runtime string; not statically resolvable",
            )
        )
    if re.search(r"\bOFFSET\s*\(", text, re.IGNORECASE):
        found.append(
            UnresolvedReference(
                reference_text="OFFSET(...)",
                reason=UnresolvedReason.VOLATILE_OFFSET,
                detail="OFFSET is volatile; its target is not statically resolvable",
            )
        )
    for match in _FUNC_RE.finditer(text):
        raw = match.group(1)
        base = raw.upper().rsplit(".", 1)[-1]
        if base in _DYNAMIC_FUNCTIONS:
            found.append(
                UnresolvedReference(
                    reference_text=f"{raw}(...)",
                    reason=UnresolvedReason.DYNAMIC_ARRAY,
                    detail="dynamic-array formula; precedents cannot be flattened to static edges",
                )
            )
        elif raw.upper().startswith(("_XLFN.", "_XLUDF.")):
            found.append(
                UnresolvedReference(
                    reference_text=f"{raw}(...)",
                    reason=UnresolvedReason.UNSUPPORTED_FUNCTION,
                    detail="function form is not supported for static reference extraction",
                )
            )
    if _AT_RE.search(text):
        found.append(
            UnresolvedReference(
                reference_text="@",
                reason=UnresolvedReason.DYNAMIC_ARRAY,
                detail="implicit-intersection '@' operator is not statically resolvable",
            )
        )
    if _SPILL_RE.search(text):
        found.append(
            UnresolvedReference(
                reference_text="#",
                reason=UnresolvedReason.DYNAMIC_ARRAY,
                detail="spill-range '#' reference is not statically resolvable",
            )
        )
    return found


def _last_bracket_token(spec: str) -> str:
    inner = spec.strip()
    # [[#All],[Amount]] -> Amount ; [#Totals] -> #Totals ; Amount -> Amount
    tokens = re.findall(r"\[([^\[\]]+)\]|([^,\[\]]+)", inner)
    flat = [str(a or b) for a, b in tokens]
    flat = [token.strip() for token in flat if token and token.strip()]
    for token in reversed(flat):
        if token and not token.startswith("#"):
            return token
    return flat[-1] if flat else ""


def _resolve_sheet(
    explicit: str | None,
    context: FormulaContext,
    unresolved: list[UnresolvedReference],
    reference_text: str,
) -> str | None:
    sheet = explicit or context.sheet_name
    if sheet in context.macro_sheets:
        unresolved.append(
            UnresolvedReference(
                reference_text=reference_text,
                reason=UnresolvedReason.MACRO_SHEET,
                detail=f"{sheet!r} is a macro sheet; its cells are not loaded",
            )
        )
        return None
    if sheet not in context.known_sheets:
        unresolved.append(
            UnresolvedReference(
                reference_text=reference_text,
                reason=UnresolvedReason.OUT_OF_RANGE,
                detail=f"no such sheet: {sheet!r}",
            )
        )
        return None
    return sheet


def parse_formula(formula: str, context: FormulaContext) -> ParsedFormula:
    """Parse one formula into typed edges and explicit gaps. Pure text analysis; no evaluation."""
    text = formula[1:] if formula.startswith("=") else formula
    unresolved: list[UnresolvedReference] = []
    unresolved.extend(_scan_functions(text))
    spans, external_unresolved = _helper_spans(text)
    unresolved.extend(external_unresolved)
    clean = _blank(text, spans)

    edges: list[ReferenceEdge] = []
    seen: set[tuple[str, str, str, str]] = set()
    structure_spans: list[tuple[int, int]] = []

    def add_edge(
        target: str,
        sheet_name: str,
        a1: str,
        kind: ReferenceKind,
        absolute: bool,
        min_row: int,
        max_row: int,
        min_col: int,
        max_col: int,
    ) -> None:
        key = (target, sheet_name, a1, kind.value)
        if key in seen:
            return
        seen.add(key)
        edges.append(
            ReferenceEdge(
                reference=Reference(
                    target_node_id=target, sheet_name=sheet_name, a1_range=a1, kind=kind
                ),
                absolute=absolute,
                min_row=min_row,
                max_row=max_row,
                min_col=min_col,
                max_col=max_col,
            )
        )

    # 1. structured table references: Name[Column]
    for match in _TABLE_REF_RE.finditer(clean):
        structure_spans.append(match.span())
        name = match.group(1)
        info = context.tables.get(name) or context.tables.get(name.lower())
        if info is None:
            unresolved.append(
                UnresolvedReference(
                    reference_text=match.group(0),
                    reason=UnresolvedReason.OUT_OF_RANGE,
                    detail=f"structured reference to unknown table {name!r}",
                )
            )
            continue
        column = _last_bracket_token(match.group(2))
        if not column or column.startswith("#"):
            add_edge(
                range_node_id(
                    context.workbook_id, context.version, info.sheet_name, info.a1_range.a1
                ),
                info.sheet_name,
                info.a1_range.a1,
                ReferenceKind.RANGE,
                True,
                info.a1_range.min_row,
                info.a1_range.max_row,
                info.a1_range.min_col,
                info.a1_range.max_col,
            )
            continue
        matched_column = next(
            (existing for existing in info.columns if existing.lower() == column.lower()), None
        )
        if matched_column is None:
            unresolved.append(
                UnresolvedReference(
                    reference_text=match.group(0),
                    reason=UnresolvedReason.OUT_OF_RANGE,
                    detail=f"table {info.name!r} has no column {column!r}",
                )
            )
            continue
        add_edge(
            table_column_node_id(
                context.workbook_id, context.version, info.sheet_name, info.name, matched_column
            ),
            info.sheet_name,
            f"{column_letter(info.a1_range.min_col)}:{column_letter(info.a1_range.max_col)}",
            ReferenceKind.TABLE_COLUMN,
            True,
            info.a1_range.min_row,
            info.a1_range.max_row,
            info.a1_range.min_col,
            info.a1_range.max_col,
        )

    # 2. whole-column references: A:B
    col_clean = _blank(clean, structure_spans)
    col_spans: list[tuple[int, int]] = []
    for match in _COL_RE.finditer(col_clean):
        col_spans.append(match.span())
        reference_text = match.group(0)
        sheet = _resolve_sheet(
            match.group("qsheet") or match.group("sheet"), context, unresolved, reference_text
        )
        if sheet is None:
            continue
        start_col = A1Range.parse(sheet, f"{match.group('start')}1")
        end_col = A1Range.parse(sheet, f"{match.group('end')}1")
        max_row = context.sheet_max_row.get(sheet, 1)
        a1 = f"{column_letter(start_col.min_col)}1:{column_letter(end_col.max_col)}{max_row}"
        add_edge(
            range_node_id(context.workbook_id, context.version, sheet, a1),
            sheet,
            a1,
            ReferenceKind.RANGE,
            False,
            1,
            max_row,
            start_col.min_col,
            end_col.max_col,
        )

    # 3. A1 cell and range references
    cell_clean = _blank(col_clean, col_spans)
    cell_spans: list[tuple[int, int]] = []
    for match in _CELL_RE.finditer(cell_clean):
        cell_spans.append(match.span())
        reference_text = match.group(0)
        explicit = match.group("qsheet") or match.group("sheet")
        sheet = _resolve_sheet(explicit, context, unresolved, reference_text)
        if sheet is None:
            continue
        start = match.group("start") or ""
        end = match.group("end")
        try:
            parsed = A1Range.parse(sheet, f"{start}:{end}" if end else start)
        except ValueError:
            unresolved.append(
                UnresolvedReference(
                    reference_text=reference_text,
                    reason=UnresolvedReason.MALFORMED,
                    detail="not a parseable A1 reference",
                )
            )
            continue
        absolute = "$" in start or (end is not None and "$" in end)
        is_cell = parsed.min_row == parsed.max_row and parsed.min_col == parsed.max_col
        kind = ReferenceKind.CELL if is_cell else ReferenceKind.RANGE
        target = (
            cell_node_id(context.workbook_id, context.version, sheet, parsed.a1)
            if is_cell
            else range_node_id(context.workbook_id, context.version, sheet, parsed.a1)
        )
        add_edge(
            target,
            sheet,
            parsed.a1,
            kind,
            absolute,
            parsed.min_row,
            parsed.max_row,
            parsed.min_col,
            parsed.max_col,
        )

    # 4. named ranges: bare identifiers that match a defined name
    ident_clean = _blank(cell_clean, cell_spans)
    for match in _IDENT_RE.finditer(ident_clean):
        name = match.group("name")
        defined = context.defined_names.get(name.lower())
        if defined is None:
            continue
        add_edge(
            named_range_node_id(context.workbook_id, context.version, defined.name),
            defined.sheet_name,
            defined.a1_range.a1,
            ReferenceKind.NAMED_RANGE,
            True,
            defined.a1_range.min_row,
            defined.a1_range.max_row,
            defined.a1_range.min_col,
            defined.a1_range.max_col,
        )

    return ParsedFormula(references=tuple(edges), unresolved=tuple(unresolved))


def resolve_named_range(
    raw_name: str,
    attr_text: str | None,
    local_sheet_id: int | None,
    *,
    workbook_id: str,
    version: int,
    known_sheets: frozenset[str],
    sheet_by_index: Mapping[int, str],
) -> NamedRange:
    """Resolve one defined name into a :class:`NamedRange`, or a kept gap if not a rectangle."""
    node = named_range_node_id(workbook_id, version, raw_name)
    scope = "workbook"
    if local_sheet_id is not None:
        scope = sheet_by_index.get(local_sheet_id, "workbook")
    if not attr_text:
        return NamedRange(raw_name, scope, None, None, node, False, None, "empty definition")
    candidate = attr_text.strip()
    if candidate.startswith("="):
        candidate = candidate[1:].strip()
    reference = re.fullmatch(
        r"(?:'(?P<q>[^']+)'|(?P<s>[A-Za-z_][A-Za-z0-9_. ]*))!\$?[A-Za-z]{1,3}\$?\d{1,7}"
        r"(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?",
        candidate,
    )
    if reference is None:
        return NamedRange(
            raw_name, scope, None, None, node, False, None, f"not a static rectangle: {attr_text!r}"
        )
    sheet_name = (reference.group("q") or reference.group("s") or "").strip()
    if sheet_name not in known_sheets:
        return NamedRange(
            raw_name,
            scope,
            sheet_name or None,
            None,
            node,
            False,
            None,
            f"points at unknown sheet {sheet_name!r}",
        )
    body = candidate.split("!", 1)[1]
    try:
        parsed = A1Range.parse(sheet_name, body)
    except ValueError as exc:
        return NamedRange(raw_name, scope, sheet_name, None, node, False, None, str(exc))
    target = (
        cell_node_id(workbook_id, version, sheet_name, parsed.a1)
        if parsed.min_row == parsed.max_row and parsed.min_col == parsed.max_col
        else range_node_id(workbook_id, version, sheet_name, parsed.a1)
    )
    return NamedRange(raw_name, scope, sheet_name, parsed.a1, node, True, target)
