"""Charts, pivot tables and data validation: sheet objects that read ranges without being formulas.

A chart plots ranges, a pivot table summarises a source range, and a list validation offers the
values of a range. None of them is a cell formula, yet each depends on cells exactly as a formula
does: change ``Data!B7`` and the chart that plots ``Data!B2:B13`` changes. They are read here as
the reference texts they store (``'Data'!$B$2:$B$13``, a pivot cache's worksheet source, a
validation's ``formula1``), and :mod:`.build` resolves those texts with the formula parser, so an
object's edges and a formula's edges are the same kind of thing and ``dependents`` finds both.

Nothing is rendered, refreshed or recalculated: a pivot's cached records and a chart's cached
points are not read, only where they come from.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

from .canonical import column_letter

#: openpyxl chart class name -> how a person names it.
_CHART_KINDS = {
    "AreaChart": "area chart",
    "AreaChart3D": "3-D area chart",
    "BarChart": "bar chart",
    "BarChart3D": "3-D bar chart",
    "BubbleChart": "bubble chart",
    "DoughnutChart": "doughnut chart",
    "LineChart": "line chart",
    "LineChart3D": "3-D line chart",
    "PieChart": "pie chart",
    "PieChart3D": "3-D pie chart",
    "ProjectedPieChart": "pie-of-pie chart",
    "RadarChart": "radar chart",
    "ScatterChart": "scatter chart",
    "StockChart": "stock chart",
    "SurfaceChart": "surface chart",
    "SurfaceChart3D": "3-D surface chart",
}

#: At most this many objects of one kind are read per sheet; the rest are counted in a warning.
MAX_OBJECTS_PER_SHEET = 500


@dataclass(frozen=True, slots=True)
class RawObject:
    """A chart, pivot table or data validation as read: where it is and what it reads.

    ``sources`` are reference texts in formula syntax, resolved later on ``anchor``'s sheet.
    ``unreadable`` names a source that is not a worksheet range (a pivot over an external
    connection), kept so it becomes a gap rather than vanishing.
    """

    kind: str  # "chart" | "pivot_table" | "data_validation"
    key: str  # unique per sheet and kind: a chart's ordinal, a pivot's name, a validation's cells
    name: str
    detail: str
    anchor: str  # the A1 cell or range the object occupies or applies to
    sources: tuple[str, ...]
    unreadable: tuple[str, ...] = ()


def read_sheet_objects(ws: Any) -> tuple[tuple[RawObject, ...], tuple[str, ...]]:
    """The charts, pivot tables and data validations of one openpyxl worksheet, and warnings.

    A malformed object is skipped with a warning naming it; it never fails the workbook.
    """
    objects: list[RawObject] = []
    warnings: list[str] = []
    for kind, reader, items in (
        ("chart", _chart, list(getattr(ws, "_charts", ()))),
        ("pivot table", _pivot, list(getattr(ws, "_pivots", ()))),
        ("data validation", _validation, list(_validations(ws))),
    ):
        if len(items) > MAX_OBJECTS_PER_SHEET:
            warnings.append(
                f"sheet {ws.title!r}: {len(items)} {kind}s; the first {MAX_OBJECTS_PER_SHEET} read"
            )
        for ordinal, item in enumerate(items[:MAX_OBJECTS_PER_SHEET], start=1):
            try:
                found = reader(item, ordinal)
            except (AttributeError, TypeError, ValueError, IndexError) as exc:
                warnings.append(f"sheet {ws.title!r}: {kind} {ordinal} not read ({exc})")
                continue
            if found is not None:
                objects.append(found)
    return tuple(objects), tuple(warnings)


def _chart(chart: Any, ordinal: int) -> RawObject:
    sources: list[str] = []
    for series in getattr(chart, "series", ()) or ():
        for part in ("tx", "cat", "val", "xVal", "yVal", "bubbleSize"):
            sources.extend(_data_source_refs(getattr(series, part, None)))
    title = _chart_title(chart)
    kind = _CHART_KINDS.get(type(chart).__name__, "chart")
    return RawObject(
        kind="chart",
        key=str(ordinal),
        name=title or f"Chart {ordinal}",
        detail=kind,
        anchor=_anchor_cell(getattr(chart, "anchor", None)),
        sources=tuple(dict.fromkeys(sources)),
    )


def _data_source_refs(source: Any) -> list[str]:
    """The ``<c:f>`` texts of a series' title, categories or values (``strRef``/``numRef``/...)."""
    if source is None:
        return []
    found: list[str] = []
    for attribute in ("strRef", "numRef", "multiLvlStrRef"):
        reference = getattr(source, attribute, None)
        text = getattr(reference, "f", None)
        if text:
            found.append(str(text))
    return found


def _chart_title(chart: Any) -> str | None:
    title = getattr(chart, "title", None)
    if title is None:
        return None
    if isinstance(title, str):
        return title.strip() or None
    text = getattr(title, "tx", None)
    rich = getattr(text, "rich", None)
    if rich is not None:
        runs = [
            run.t
            for paragraph in getattr(rich, "p", ()) or ()
            for run in (getattr(paragraph, "r", ()) or ())
            if getattr(run, "t", None)
        ]
        joined = " ".join("".join(runs).split())
        return joined or None
    reference = getattr(getattr(text, "strRef", None), "f", None)
    return f"title from {reference}" if reference else None


def _anchor_cell(anchor: Any) -> str:
    """The top-left cell a drawing is anchored at (``_from`` is zero-based)."""
    if isinstance(anchor, str):
        return anchor.replace("$", "")
    origin = getattr(anchor, "_from", None)
    if origin is None:
        return "A1"
    return f"{column_letter(int(origin.col) + 1)}{int(origin.row) + 1}"


def _pivot(pivot: Any, ordinal: int) -> RawObject:
    name = str(getattr(pivot, "name", None) or f"PivotTable{ordinal}")
    location = getattr(getattr(pivot, "location", None), "ref", None) or "A1"
    cache_source = getattr(getattr(pivot, "cache", None), "cacheSource", None)
    worksheet = getattr(cache_source, "worksheetSource", None)
    sources: list[str] = []
    unreadable: list[str] = []
    if worksheet is not None and getattr(worksheet, "name", None):
        sources.append(str(worksheet.name))  # a table or defined name
    elif worksheet is not None and getattr(worksheet, "ref", None):
        sheet = getattr(worksheet, "sheet", None)
        if getattr(worksheet, "id", None):
            unreadable.append(f"[{worksheet.id}]{sheet or ''}!{worksheet.ref}")
        elif sheet:
            sources.append(f"'{str(sheet).replace(chr(39), chr(39) * 2)}'!{worksheet.ref}")
        else:
            unreadable.append(str(worksheet.ref))
    else:
        source_type = getattr(cache_source, "type", None) or "unknown"
        unreadable.append(f"{source_type} source")
    return RawObject(
        kind="pivot_table",
        key=name,
        name=name,
        detail="pivot table",
        anchor=str(location),
        sources=tuple(sources),
        unreadable=tuple(unreadable),
    )


def _validations(ws: Any) -> list[Any]:
    container = getattr(ws, "data_validations", None)
    return list(getattr(container, "dataValidation", ()) or ())


#: A validation formula that is only a constant reads nothing: ``"a,b,c"``, ``5``, ``TRUE``.
_CONSTANT_RE = re.compile(
    r'=?\s*(?:"(?:[^"]|"")*"|[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?|TRUE|FALSE)\s*'
)


def _validation(validation: Any, ordinal: int) -> RawObject | None:
    cells = " ".join(str(getattr(validation, "sqref", "") or "").split())
    if not cells:
        return None
    sources = [
        str(text)
        for text in (getattr(validation, "formula1", None), getattr(validation, "formula2", None))
        if text and not _CONSTANT_RE.fullmatch(str(text))
    ]
    if not sources:
        return None  # a constant list or bound reads no cell: nothing to link
    kind = str(getattr(validation, "type", None) or "any")
    return RawObject(
        kind="data_validation",
        key=cells,
        name=f"{kind} validation on {cells}",
        detail=f"{kind} validation",
        anchor=cells.split(" ", 1)[0],
        sources=tuple(sources),
    )


__all__ = ["MAX_OBJECTS_PER_SHEET", "RawObject", "read_sheet_objects"]
