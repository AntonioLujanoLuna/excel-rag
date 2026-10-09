"""Static formula reference extraction -- never evaluation.

A formula is turned into typed :class:`~excel_rag.models.Reference` edges by text analysis alone.
Two inviolable rules from the design:

* **Never invent an edge.** A precedent that cannot be determined statically (``INDIRECT``, volatile
  ``OFFSET``, an external workbook, an unsupported dynamic array) becomes an
  :class:`~excel_rag.models.UnresolvedReference` carrying the reason, not a guessed target.
* **Edges point at ranges, not cells.** ``SUM(Actuals!D2:D500)`` is *one* range edge; the rectangle
  is indexed once and overlaps are answered by an ``integer_range`` query at retrieval time.

Tokenising is openpyxl's (:class:`openpyxl.formula.tokenizer.Tokenizer`): it separates function
calls, string literals and reference operands, so a function whose name spells a cell (``LOG10``)
or a string that spells one (``"B7"``) is never mistaken for a reference. Each ``OPERAND RANGE``
token is then resolved on its own: a cell or rectangle, a whole column or row, a 3-D reference
across a run of sheets, a structured table reference, or a defined name.

The declared node ids come from :mod:`excel_rag.workbook.canonical`, so an edge target and the
structure node it points at are produced by the same convention.
"""

from __future__ import annotations

import re
from collections.abc import Mapping
from dataclasses import dataclass, field

from openpyxl.formula.tokenizer import (  # type: ignore[import-untyped]
    Token,
    Tokenizer,
    TokenizerError,
)

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
    r"'[^']*\[[^\]]+\][^']*'|\[[^\]]+\.(?:xlsx?|xlsm|xlsb|csv|xls)\]|^\[\d+\]", re.IGNORECASE
)
#: A spill operator: ``#`` right after a reference or a call, before an operator or the end. The
#: tokenizer refuses it, so it is recorded as a gap and stripped before tokenising.
_SPILL_RE = re.compile(r"(?<=[A-Za-z0-9_$)\]])\s*#(?=\s*(?:$|[-+*/^&=<>,;)\s:]))")
_AT_RE = re.compile(r"(?<![\w.\[])@")
_SHEET_REF_RE = re.compile(r"(?:'(?P<quoted>(?:[^']|'')+)'|(?P<bare>[^'!]+))!(?P<ref>.+)")
_CELL_REF_RE = re.compile(r"\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?")
_COLUMNS_REF_RE = re.compile(r"\$?(?P<start>[A-Za-z]{1,3}):\$?(?P<end>[A-Za-z]{1,3})")
_ROWS_REF_RE = re.compile(r"\$?(?P<start>\d{1,7}):\$?(?P<end>\d{1,7})")
_TABLE_REF_RE = re.compile(r"(?P<table>[A-Za-z_\\][A-Za-z0-9_.]*)?\s*\[(?P<spec>.*)\]")
_NAME_RE = re.compile(r"[A-Za-z_\\][A-Za-z0-9_.]*")
#: Functions that bind local names; inside them a bare identifier is a variable, not a name.
_BINDING_FUNCTIONS = frozenset({"LET", "LAMBDA"})

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
    #: Workbook sheet order, which a 3-D reference (``Jan:Mar!B2``) spans. Empty means unknown, and
    #: a 3-D reference is then a gap rather than a guess.
    sheet_order: tuple[str, ...] = ()
    #: Used width per sheet, the extent a whole-row reference (``2:2``) is clipped to.
    sheet_max_col: Mapping[str, int] = field(default_factory=dict)


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


def _gap(text: str, reason: UnresolvedReason, detail: str) -> UnresolvedReference:
    return UnresolvedReference(reference_text=text, reason=reason, detail=detail)


def _function_gaps(raw_name: str) -> list[UnresolvedReference]:
    """The gaps a function call implies, by name alone."""
    upper = raw_name.upper()
    base = upper.rsplit(".", 1)[-1]
    if base == "INDIRECT":
        return [
            _gap(
                "INDIRECT(...)",
                UnresolvedReason.INDIRECT,
                "INDIRECT target depends on a runtime string; not statically resolvable",
            )
        ]
    if base == "OFFSET":
        return [
            _gap(
                "OFFSET(...)",
                UnresolvedReason.VOLATILE_OFFSET,
                "OFFSET is volatile; its target is not statically resolvable",
            )
        ]
    if base in _DYNAMIC_FUNCTIONS:
        return [
            _gap(
                f"{raw_name}(...)",
                UnresolvedReason.DYNAMIC_ARRAY,
                "dynamic-array formula; precedents cannot be flattened to static edges",
            )
        ]
    if upper.startswith(("_XLFN.", "_XLUDF.")):
        return [
            _gap(
                f"{raw_name}(...)",
                UnresolvedReason.UNSUPPORTED_FUNCTION,
                "function form is not supported for static reference extraction",
            )
        ]
    return []


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


def _check_sheet(
    sheet: str,
    context: FormulaContext,
    unresolved: list[UnresolvedReference],
    reference_text: str,
) -> bool:
    if sheet in context.macro_sheets:
        unresolved.append(
            _gap(
                reference_text,
                UnresolvedReason.MACRO_SHEET,
                f"{sheet!r} is a macro sheet; its cells are not loaded",
            )
        )
        return False
    if sheet not in context.known_sheets:
        unresolved.append(
            _gap(reference_text, UnresolvedReason.OUT_OF_RANGE, f"no such sheet: {sheet!r}")
        )
        return False
    return True


class _Collector:
    """Accumulates de-duplicated edges and gaps for one formula."""

    def __init__(self, context: FormulaContext) -> None:
        self.context = context
        self.edges: list[ReferenceEdge] = []
        self.unresolved: list[UnresolvedReference] = []
        self._seen_edges: set[tuple[str, str, str, str]] = set()

    def edge(
        self,
        target: str,
        sheet_name: str,
        a1: str,
        kind: ReferenceKind,
        absolute: bool,
        bounds: tuple[int, int, int, int],
    ) -> None:
        key = (target, sheet_name, a1, kind.value)
        if key in self._seen_edges:
            return
        self._seen_edges.add(key)
        min_row, max_row, min_col, max_col = bounds
        self.edges.append(
            ReferenceEdge(
                reference=Reference(
                    target_node_id=target,
                    sheet_name=sheet_name,
                    a1_range=a1,
                    kind=kind,
                    row_span={"gte": min_row, "lte": max_row},
                    column_span={"gte": min_col, "lte": max_col},
                ),
                absolute=absolute,
                min_row=min_row,
                max_row=max_row,
                min_col=min_col,
                max_col=max_col,
            )
        )

    def gap(self, text: str, reason: UnresolvedReason, detail: str) -> None:
        self.unresolved.append(_gap(text, reason, detail))

    def rectangle(self, sheet: str, region: A1Range, absolute: bool) -> None:
        """An edge to a cell node, or to a range node for anything wider."""
        context = self.context
        is_cell = region.cell_count == 1
        target = (
            cell_node_id(context.workbook_id, context.version, sheet, region.a1)
            if is_cell
            else range_node_id(context.workbook_id, context.version, sheet, region.a1)
        )
        self.edge(
            target,
            sheet,
            region.a1,
            ReferenceKind.CELL if is_cell else ReferenceKind.RANGE,
            absolute,
            (region.min_row, region.max_row, region.min_col, region.max_col),
        )


def _resolve_table(collector: _Collector, text: str, table: str | None, spec: str) -> None:
    context = collector.context
    if table is None:
        collector.gap(
            text,
            UnresolvedReason.OUT_OF_RANGE,
            "structured reference without a table name; its table depends on the formula's cell",
        )
        return
    info = context.tables.get(table) or context.tables.get(table.lower())
    if info is None:
        collector.gap(
            text, UnresolvedReason.OUT_OF_RANGE, f"structured reference to unknown table {table!r}"
        )
        return
    bounds = (
        info.a1_range.min_row,
        info.a1_range.max_row,
        info.a1_range.min_col,
        info.a1_range.max_col,
    )
    column = _last_bracket_token(spec)
    if not column or column.startswith("#"):
        collector.edge(
            range_node_id(context.workbook_id, context.version, info.sheet_name, info.a1_range.a1),
            info.sheet_name,
            info.a1_range.a1,
            ReferenceKind.RANGE,
            True,
            bounds,
        )
        return
    matched = next((name for name in info.columns if name.lower() == column.lower()), None)
    if matched is None:
        collector.gap(
            text, UnresolvedReason.OUT_OF_RANGE, f"table {info.name!r} has no column {column!r}"
        )
        return
    collector.edge(
        table_column_node_id(
            context.workbook_id, context.version, info.sheet_name, info.name, matched
        ),
        info.sheet_name,
        f"{column_letter(info.a1_range.min_col)}:{column_letter(info.a1_range.max_col)}",
        ReferenceKind.TABLE_COLUMN,
        True,
        bounds,
    )


def _resolve_on_sheet(collector: _Collector, text: str, sheet: str, ref: str) -> bool:
    """Resolve the part after ``Sheet!`` (or a bare operand) on one sheet.

    Returns ``False`` when ``ref`` is not a coordinate form at all, so the caller can try it as a
    defined name.
    """
    context = collector.context
    absolute = "$" in ref
    if _CELL_REF_RE.fullmatch(ref):
        try:
            region = A1Range.parse(sheet, ref)
        except ValueError:
            collector.gap(text, UnresolvedReason.MALFORMED, "not a parseable A1 reference")
            return True
        collector.rectangle(sheet, region, absolute)
        return True
    columns = _COLUMNS_REF_RE.fullmatch(ref)
    if columns is not None:
        first = A1Range.parse(sheet, f"{columns.group('start')}1").min_col
        last = A1Range.parse(sheet, f"{columns.group('end')}1").min_col
        first, last = min(first, last), max(first, last)
        max_row = context.sheet_max_row.get(sheet, 1)
        a1 = f"{column_letter(first)}1:{column_letter(last)}{max_row}"
        collector.edge(
            range_node_id(context.workbook_id, context.version, sheet, a1),
            sheet,
            a1,
            ReferenceKind.RANGE,
            absolute,
            (1, max_row, first, last),
        )
        return True
    rows = _ROWS_REF_RE.fullmatch(ref)
    if rows is not None:
        first, last = sorted((int(rows.group("start")), int(rows.group("end"))))
        if first < 1:
            collector.gap(text, UnresolvedReason.MALFORMED, "row 0 does not exist")
            return True
        max_col = context.sheet_max_col.get(sheet, 1)
        a1 = f"A{first}:{column_letter(max_col)}{last}"
        collector.edge(
            range_node_id(context.workbook_id, context.version, sheet, a1),
            sheet,
            a1,
            ReferenceKind.RANGE,
            absolute,
            (first, last, 1, max_col),
        )
        return True
    if "#REF!" in ref.upper():
        collector.gap(text, UnresolvedReason.MALFORMED, "broken reference (#REF!)")
        return True
    return False


def _resolve_name(collector: _Collector, text: str, name: str, *, binds_names: bool) -> None:
    context = collector.context
    defined = context.defined_names.get(name.lower())
    if defined is not None:
        collector.edge(
            named_range_node_id(context.workbook_id, context.version, defined.name),
            defined.sheet_name,
            defined.a1_range.a1,
            ReferenceKind.NAMED_RANGE,
            True,
            (
                defined.a1_range.min_row,
                defined.a1_range.max_row,
                defined.a1_range.min_col,
                defined.a1_range.max_col,
            ),
        )
        return
    if binds_names:
        # Inside LET/LAMBDA a bare identifier is a local variable; the formula is already a gap.
        return
    collector.gap(
        text,
        UnresolvedReason.OUT_OF_RANGE,
        f"name {name!r} is undefined or does not resolve to a static rectangle",
    )


def _resolve_operand(collector: _Collector, raw: str, *, binds_names: bool) -> None:
    context = collector.context
    text = raw.strip()
    if text.startswith("@"):
        text = text[1:]
    if not text:
        return
    if _EXTERNAL_RE.search(text):
        collector.gap(
            raw.strip(),
            UnresolvedReason.EXTERNAL_LINK,
            "external workbook or link index is not resolved during ingestion",
        )
        return

    table = _TABLE_REF_RE.fullmatch(text)
    if table is not None:
        _resolve_table(collector, text, table.group("table"), table.group("spec"))
        return

    qualified = _SHEET_REF_RE.fullmatch(text)
    if qualified is None:
        if not _resolve_on_sheet(collector, text, context.sheet_name, text):
            if _NAME_RE.fullmatch(text):
                _resolve_name(collector, text, text, binds_names=binds_names)
            else:
                collector.gap(text, UnresolvedReason.MALFORMED, "unrecognised reference form")
        return

    quoted = qualified.group("quoted")
    sheet_part = quoted.replace("''", "'") if quoted is not None else qualified.group("bare")
    ref = qualified.group("ref")
    # `Sheet1!A1:Sheet1!B2` names the sheet twice; anything else across sheets is not a rectangle.
    if "!" in ref:
        head, _, tail = ref.partition(":")
        tail_match = _SHEET_REF_RE.fullmatch(tail)
        if tail_match is None:
            collector.gap(text, UnresolvedReason.MALFORMED, "unrecognised reference form")
            return
        tail_quoted = tail_match.group("quoted")
        tail_sheet = (
            tail_quoted.replace("''", "'") if tail_quoted is not None else tail_match.group("bare")
        )
        if tail_sheet != sheet_part:
            collector.gap(text, UnresolvedReason.MALFORMED, "a range cannot span two sheets")
            return
        ref = f"{head}:{tail_match.group('ref')}"

    if ":" in sheet_part:
        _resolve_three_d(collector, text, sheet_part, ref)
        return
    if not _check_sheet(sheet_part, context, collector.unresolved, text):
        return
    if not _resolve_on_sheet(collector, text, sheet_part, ref):
        if _NAME_RE.fullmatch(ref):
            # A sheet-scoped defined name: `Inputs!Rate`.
            _resolve_name(collector, text, ref, binds_names=binds_names)
        else:
            collector.gap(text, UnresolvedReason.MALFORMED, "unrecognised reference form")


def _resolve_three_d(collector: _Collector, text: str, sheet_part: str, ref: str) -> None:
    """``Jan:Mar!B2`` reads ``B2`` on every sheet from ``Jan`` to ``Mar``, in workbook order."""
    context = collector.context
    first, _, last = sheet_part.partition(":")
    order = context.sheet_order
    if not order:
        collector.gap(
            text,
            UnresolvedReason.OUT_OF_RANGE,
            "3-D reference, but the workbook's sheet order is unknown",
        )
        return
    for sheet in (first, last):
        if sheet not in order:
            collector.gap(text, UnresolvedReason.OUT_OF_RANGE, f"no such sheet: {sheet!r}")
            return
    start, end = sorted((order.index(first), order.index(last)))
    for sheet in order[start : end + 1]:
        if not _check_sheet(sheet, context, collector.unresolved, text):
            continue
        if not _resolve_on_sheet(collector, text, sheet, ref):
            collector.gap(text, UnresolvedReason.MALFORMED, "unrecognised reference form")
            return


def parse_formula(formula: str, context: FormulaContext) -> ParsedFormula:
    """Parse one formula into typed edges and explicit gaps. Pure text analysis; no evaluation."""
    text = formula[1:] if formula.startswith("=") else formula
    collector = _Collector(context)

    # Spill and implicit-intersection operators are judged on the text outside string literals.
    outside_strings = _blank(text, [match.span() for match in _STRING_RE.finditer(text)])
    if _AT_RE.search(outside_strings):
        collector.gap(
            "@",
            UnresolvedReason.DYNAMIC_ARRAY,
            "implicit-intersection '@' operator is not statically resolvable",
        )
    spills = [match.span() for match in _SPILL_RE.finditer(outside_strings)]
    if spills:
        collector.gap(
            "#",
            UnresolvedReason.DYNAMIC_ARRAY,
            "spill-range '#' reference is not statically resolvable",
        )

    try:
        tokens: list[Token] = Tokenizer("=" + _blank_out(text, spills)).items
    except TokenizerError as exc:
        collector.gap(
            text[:200], UnresolvedReason.MALFORMED, f"formula could not be tokenised: {exc}"
        )
        return ParsedFormula(references=(), unresolved=tuple(collector.unresolved))

    binds_names = any(
        token.type == Token.FUNC
        and token.subtype == Token.OPEN
        and token.value[:-1].strip().upper().rsplit(".", 1)[-1] in _BINDING_FUNCTIONS
        for token in tokens
    )
    for token in tokens:
        if token.type == Token.FUNC and token.subtype == Token.OPEN:
            name = token.value[:-1].strip()
            if ":" in name:
                # `A1:INDEX(...)`: one end of the range is a function result.
                anchor, _, name = name.rpartition(":")
                collector.gap(
                    f"{anchor}:{name}(...)",
                    UnresolvedReason.UNSUPPORTED_FUNCTION,
                    "a range bounded by a function result is not statically resolvable",
                )
            collector.unresolved.extend(_function_gaps(name))
        elif token.type == Token.OPERAND and token.subtype == Token.RANGE:
            _resolve_operand(collector, token.value, binds_names=binds_names)
        elif token.type == Token.OPERAND and token.subtype == Token.ERROR:
            if token.value.upper() == "#REF!":
                collector.gap(token.value, UnresolvedReason.MALFORMED, "broken reference (#REF!)")

    return ParsedFormula(references=tuple(collector.edges), unresolved=tuple(collector.unresolved))


def _blank_out(text: str, spans: list[tuple[int, int]]) -> str:
    """Delete the given spans (the spill operators the tokenizer refuses)."""
    if not spans:
        return text
    pieces: list[str] = []
    cursor = 0
    for start, end in spans:
        pieces.append(text[cursor:start])
        cursor = end
    pieces.append(text[cursor:])
    return "".join(pieces)


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
