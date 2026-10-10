"""Read an ``.xlsx`` / ``.xlsm`` into raw sheets -- no macros, no link refresh, no evaluation.

A CSV file and a binary ``.xlsb`` are read too, into the same raw sheets, by
:mod:`.formats`; what those formats cannot carry is stated in the workbook's warnings.

The rules that govern this module:

* ``data_only=False`` gives the formula text; the last value Excel saved for each formula cell comes
  from one streaming ``expat`` pass over the sheet part (:mod:`.sheetscan`), which also reads the
  declared dimension. openpyxl exposes the text and the saved value under one flag or the other,
  never together, and a second ``data_only=True`` load -- a whole worksheet model to read one
  ``<v>`` per formula -- is now only the fallback. The cached value is kept separately and
  labelled as cached, never recomputed.
* Macros are never loaded (``keep_vba=False``) and never executed -- openpyxl does not run VBA, and
  a workbook carrying a ``vbaProject.bin`` or a macro sheet is only *noticed*, so a formula pointing
  at a macro sheet resolves to an explicit ``macro_sheet`` gap rather than a fabricated edge.
* An array formula (a legacy CSE formula, or a dynamic-array formula Excel 365 saved) is read as
  its text, and the other cells of its saved extent are *computed* cells: their value is the one
  Excel last saved, labelled as cached and pointing at the formula that produced it -- never
  mistaken for input data.
* Package XML is parsed with ``defusedxml`` (no entity expansion), and openpyxl uses it too once it
  is installed, so a billion-laughs part is refused rather than expanded.
* A declared dimension is not trusted. openpyxl already ignores ``<dimension ref="A1:XFD1048576">``,
  so nothing is allocated for it; we read the string from the package only to *report* that a
  hostile dimension was declared and ignored.
"""

from __future__ import annotations

import hashlib
import io
import re
import zipfile
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Any
from xml.etree.ElementTree import Element, ParseError

import openpyxl  # type: ignore[import-untyped]
from defusedxml import DefusedXmlException  # type: ignore[import-untyped]
from defusedxml.ElementTree import fromstring as _fromstring  # type: ignore[import-untyped]
from openpyxl.utils import get_column_letter  # type: ignore[import-untyped]
from openpyxl.utils.exceptions import InvalidFileException  # type: ignore[import-untyped]
from openpyxl.worksheet.formula import (  # type: ignore[import-untyped]
    ArrayFormula,
    DataTableFormula,
)

from .canonical import CellValue
from .errors import WorkbookError
from .formats import looks_like_xlsb, read_csv_cells, read_xlsb_sheets, sheet_name_for
from .sheetscan import SheetScan, cached_value, scan_sheet

_XLNS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_RELNS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CTNS = "http://schemas.openxmlformats.org/package/2006/content-types"
_SHEET_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
_MACRO_REL_SUFFIX = "/xlMacrosheet"

#: A declared dimension larger than this many cells is reported as ignored, not honoured.
MAX_DECLARED_CELLS = 5_000_000
#: File suffixes read as CSV (when the bytes are not a zip package).
CSV_SUFFIXES = (".csv", ".tsv", ".txt")
#: Refuse a package whose declared uncompressed size is absurd, before openpyxl reads it.
MAX_UNCOMPRESSED_BYTES = 512 * 1024 * 1024

_CURRENCY = re.compile(r"[$€£¥₹]")
_DIMENSION = re.compile(r"\$?([A-Za-z]{1,3})\$?(\d{1,9})(?::\$?([A-Za-z]{1,3})\$?(\d{1,9}))?")


@dataclass(frozen=True, slots=True)
class RawTable:
    """An Excel table object (``ws.tables``) as read, before it is matched to a region."""

    name: str
    ref: str
    columns: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class RawDefinedName:
    """A defined name as read. ``scope_sheet`` is the sheet it is local to, ``None`` for
    workbook-wide."""

    name: str
    attr_text: str | None
    scope_sheet: str | None = None


@dataclass(frozen=True, slots=True)
class RawSheet:
    """A worksheet as read: populated cells, merges, tables, and package-level flags."""

    name: str
    visibility: str
    cells: dict[str, CellValue]
    merged_ranges: tuple[str, ...]
    tables: tuple[RawTable, ...]
    declared_dimension: str | None = None
    declared_dimension_flagged: bool = False
    is_macro_sheet: bool = False


@dataclass(frozen=True, slots=True)
class RawWorkbook:
    """Everything the reader extracted, before region detection."""

    path: str
    source_file: str
    source_sha256: str
    sheets: tuple[RawSheet, ...]
    defined_names: tuple[RawDefinedName, ...]
    macro_sheet_names: tuple[str, ...] = ()
    has_vba: bool = False
    warnings: tuple[str, ...] = field(default_factory=tuple)


@dataclass(frozen=True, slots=True)
class _SheetFact:
    name: str
    state: str
    part: str
    is_macro: bool


@dataclass(frozen=True, slots=True)
class _PackageFacts:
    sheets: tuple[_SheetFact, ...]
    declared_dimensions: dict[str, str]
    has_vba: bool
    #: One streaming pass per worksheet part (see :mod:`.sheetscan`), keyed by part name. A part
    #: that could not be scanned is absent, and the reader falls back to openpyxl's cached load.
    scans: dict[str, SheetScan] = field(default_factory=dict)

    @property
    def macro_names(self) -> frozenset[str]:
        return frozenset(sheet.name for sheet in self.sheets if sheet.is_macro)

    @property
    def by_name(self) -> dict[str, _SheetFact]:
        return {sheet.name: sheet for sheet in self.sheets}


def column_index(letters: str) -> int:
    """``A`` -> 1, ``Z`` -> 26, ``AA`` -> 27."""
    index = 0
    for char in letters.upper():
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index


def _dimension_area(ref: str) -> int:
    match = _DIMENSION.fullmatch(ref.strip())
    if match is None:
        return 0
    start_col, start_row, end_col, end_row = match.groups()
    min_col, min_row = column_index(start_col), int(start_row)
    max_col = column_index(end_col) if end_col else min_col
    max_row = int(end_row) if end_row else min_row
    if max_row < min_row or max_col < min_col:
        return 0
    return (max_row - min_row + 1) * (max_col - min_col + 1)


def merged_top_left(ref: str) -> str:
    """``"A1:B1"`` -> ``"A1"``: the only cell of a merged range that carries a value."""
    start = ref.split(":", 1)[0]
    return start.replace("$", "").upper()


def _formula_text(value: Any) -> tuple[str, str | None]:
    """A formula cell's text, and the extent an array formula covers (``None`` for a plain one).

    openpyxl gives an array formula as an :class:`ArrayFormula` and a what-if data table as a
    :class:`DataTableFormula`, not as a string; ``str()`` of either is an object repr.
    """
    if isinstance(value, ArrayFormula):
        text = value.text or ""
        ref = str(value.ref) if value.ref else None
        return (text if text.startswith("=") else f"={text}"), ref
    if isinstance(value, DataTableFormula):
        # Excel shows a data table as {=TABLE(row_input, column_input)}; the inputs are cells.
        first = value.r1 or ""
        second = value.r2 or ""
        if value.dt2D:
            arguments = f"{first},{second}"
        elif value.dtr:
            arguments = f"{first},"
        else:
            arguments = f",{first}"
        return f"=TABLE({arguments})", str(value.ref) if value.ref else None
    return str(value), None


def _extent_bounds(ref: str) -> tuple[int, int, int, int] | None:
    """``B2:B40`` -> ``(min_row, max_row, min_col, max_col)``, or ``None`` if unparseable."""
    match = _DIMENSION.fullmatch(ref.strip())
    if match is None:
        return None
    start_col, start_row, end_col, end_row = match.groups()
    min_col, min_row = column_index(start_col), int(start_row)
    max_col = column_index(end_col) if end_col else min_col
    max_row = int(end_row) if end_row else min_row
    if max_row < min_row or max_col < min_col:
        return None
    return min_row, max_row, min_col, max_col


def _mark_array_extents(cells: dict[str, CellValue], extents: dict[str, str]) -> None:
    """Turn the bare values inside each array formula's extent into cached formula results.

    Excel writes an extent cell as a bare ``<v>``, which openpyxl reads as a constant. The work per
    extent is bounded by the cells that exist, not by the declared extent, so a hostile
    ``ref="A1:XFD1048576"`` costs one pass over the populated cells.
    """
    for master, ref in extents.items():
        bounds = _extent_bounds(ref)
        if bounds is None:
            continue
        min_row, max_row, min_col, max_col = bounds
        area = (max_row - min_row + 1) * (max_col - min_col + 1)
        if area <= len(cells):
            candidates = [
                f"{get_column_letter(col)}{row}"
                for row in range(min_row, max_row + 1)
                for col in range(min_col, max_col + 1)
            ]
        else:
            candidates = [
                coordinate
                for coordinate, cell in cells.items()
                if min_row <= cell.row <= max_row and min_col <= cell.column <= max_col
            ]
        for coordinate in candidates:
            cell = cells.get(coordinate)
            if cell is None or coordinate == master or cell.computed:
                continue
            cells[coordinate] = replace(
                cell, value=None, cached_value=cell.value, array_master=master
            )


def _parse_xml(data: bytes, part: str) -> Element:
    """Parse one package part with ``defusedxml``: no entity expansion, no external entities.

    A malformed or hostile part is a :class:`WorkbookError` like any other unreadable package.
    """
    try:
        return _fromstring(data)  # type: ignore[no-any-return]
    except DefusedXmlException as exc:
        raise WorkbookError(f"refusing workbook: {part} declares XML entities ({exc})") from exc
    except ParseError as exc:
        raise WorkbookError(f"not a valid workbook package: {part} is not XML ({exc})") from exc


def _open_archive(data: bytes) -> zipfile.ZipFile:
    """The package as a zip, refused when it is not one or declares an absurd unpacked size."""
    try:
        archive = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile as exc:
        raise WorkbookError(
            "not a valid .xlsx/.xlsm package (not a zip archive; the file may be corrupt, "
            "truncated, or password-protected)"
        ) from exc
    uncompressed = sum(item.file_size for item in archive.infolist())
    if uncompressed > MAX_UNCOMPRESSED_BYTES:
        raise WorkbookError(
            f"refusing workbook: declared uncompressed size {uncompressed} bytes exceeds "
            f"the {MAX_UNCOMPRESSED_BYTES}-byte ceiling"
        )
    return archive


def _read_package(data: bytes) -> _PackageFacts:
    archive = _open_archive(data)
    names = archive.namelist()
    if "xl/workbook.xml" not in names:
        raise WorkbookError("not an Excel workbook package: xl/workbook.xml is missing")

    relationships: dict[str, tuple[str, str]] = {}
    if "xl/_rels/workbook.xml.rels" in names:
        rels = _parse_xml(archive.read("xl/_rels/workbook.xml.rels"), "xl/_rels/workbook.xml.rels")
        for node in rels:
            resolved = node.attrib.get("Target", "")
            target = resolved.lstrip("/")
            if not resolved.startswith("/") and not target.startswith("xl/"):
                target = f"xl/{target}"
            relationships[node.attrib.get("Id", "")] = (node.attrib.get("Type", ""), target)

    content_types: dict[str, str] = {}
    if "[Content_Types].xml" in names:
        types = _parse_xml(archive.read("[Content_Types].xml"), "[Content_Types].xml")
        for node in types:
            part = node.attrib.get("PartName")
            if part:
                content_types[part.lstrip("/")] = node.attrib.get("ContentType", "")

    sheets: list[_SheetFact] = []
    declared: dict[str, str] = {}
    scans: dict[str, SheetScan] = {}
    workbook = _parse_xml(archive.read("xl/workbook.xml"), "xl/workbook.xml")
    sheet_container = workbook.find(f"{{{_XLNS}}}sheets")
    if sheet_container is not None:
        for node in sheet_container:
            name = node.attrib.get("name", "")
            rel_id = node.attrib.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", ""
            )
            rel_type, target = relationships.get(rel_id, ("", ""))
            is_macro = rel_type.endswith(_MACRO_REL_SUFFIX) or content_types.get(
                target, ""
            ).endswith("macrosheet+xml")
            sheets.append(
                _SheetFact(
                    name=name,
                    state=node.attrib.get("state", "visible"),
                    part=target,
                    is_macro=is_macro,
                )
            )
            if target in names and not is_macro:
                try:
                    with archive.open(target) as stream:
                        scan = scan_sheet(stream, target)
                except WorkbookError:
                    continue  # openpyxl's own load decides whether this part is readable
                scans[target] = scan
                if scan.dimension:
                    declared[target] = scan.dimension

    has_vba = any(name.endswith("vbaProject.bin") for name in names)
    return _PackageFacts(
        sheets=tuple(sheets), declared_dimensions=declared, has_vba=has_vba, scans=scans
    )


def _load(data: bytes, *, data_only: bool) -> Any:
    try:
        return openpyxl.load_workbook(
            io.BytesIO(data), data_only=data_only, keep_vba=False, read_only=False
        )
    except DefusedXmlException as exc:
        raise WorkbookError(f"refusing workbook: a part declares XML entities ({exc})") from exc
    except (
        InvalidFileException,
        zipfile.BadZipFile,
        KeyError,
        ValueError,
        OSError,
        ParseError,
    ) as exc:
        raise WorkbookError(f"could not read the workbook: {exc}") from exc


def _read_csv(data: bytes, *, label: str, location: str, digest: str) -> RawWorkbook:
    if len(data) > MAX_UNCOMPRESSED_BYTES:
        raise WorkbookError(
            f"refusing CSV: {len(data)} bytes exceeds the {MAX_UNCOMPRESSED_BYTES}-byte ceiling"
        )
    sheet_name = sheet_name_for(label)
    sheet = RawSheet(
        name=sheet_name,
        visibility="visible",
        cells=read_csv_cells(data, sheet_name),
        merged_ranges=(),
        tables=(),
    )
    return RawWorkbook(
        path=location,
        source_file=label,
        source_sha256=digest,
        sheets=(sheet,),
        defined_names=(),
        warnings=("CSV: one sheet of values; numbers and TRUE/FALSE are typed, the rest is text",),
    )


def _read_xlsb(data: bytes, *, label: str, location: str, digest: str) -> RawWorkbook:
    archive = _open_archive(data)
    sheets, warnings = read_xlsb_sheets(archive)
    has_vba = any(name.endswith("vbaProject.bin") for name in archive.namelist())
    return RawWorkbook(
        path=location,
        source_file=label,
        source_sha256=digest,
        sheets=tuple(
            RawSheet(name=name, visibility="visible", cells=cells, merged_ranges=(), tables=())
            for name, cells in sheets
        ),
        defined_names=(),
        has_vba=has_vba,
        warnings=tuple(warnings),
    )


def read_workbook(path: str | Path | bytes, *, name: str | None = None) -> RawWorkbook:
    """Read a workbook into :class:`RawWorkbook`, or raise :class:`WorkbookError`.

    ``path`` is a file path, or the workbook's bytes (a chat upload never touches disk); ``name``
    then labels it in messages and in ``source_file``. The same package guards apply either way.
    Never loads or executes macros, never refreshes external links, never evaluates a formula.
    """
    if isinstance(path, bytes):
        data = path
        label = name or "workbook.xlsx"
        location = label
    else:
        source = Path(path)
        if not source.is_file():
            raise WorkbookError(f"no such workbook: {source}")
        data = source.read_bytes()
        label = name or source.name
        location = str(source)
    digest = hashlib.sha256(data).hexdigest()
    if not data.startswith(b"PK") and Path(label).suffix.lower() in CSV_SUFFIXES:
        return _read_csv(data, label=label, location=location, digest=digest)
    if looks_like_xlsb(data):
        return _read_xlsb(data, label=label, location=location, digest=digest)
    facts = _read_package(data)

    formula_wb = _load(data, data_only=False)
    # The saved values of formula cells come from the streaming scan when every worksheet part was
    # scanned and none needs the shared-string table; otherwise from openpyxl's data_only load.
    scanned = all(
        fact.part in facts.scans and not facts.scans[fact.part].needs_shared_strings
        for fact in facts.sheets
        if not fact.is_macro and fact.name in formula_wb.sheetnames
    )
    cached_wb: Any = None
    if not scanned:
        try:
            cached_wb = _load(data, data_only=True)
        except WorkbookError:
            cached_wb = None

    macro_names = facts.macro_names
    sheets: list[RawSheet] = []
    for ws in formula_wb.worksheets:
        name = ws.title
        fact = facts.by_name.get(name)
        declared_ref = facts.declared_dimensions.get(fact.part) if fact else None
        if name in macro_names:
            sheets.append(
                RawSheet(
                    name=name,
                    visibility=fact.state if fact else ws.sheet_state,
                    cells={},
                    merged_ranges=(),
                    tables=(),
                    declared_dimension=declared_ref,
                    declared_dimension_flagged=False,
                    is_macro_sheet=True,
                )
            )
            continue

        merged_ranges = tuple(str(item) for item in ws.merged_cells.ranges)
        master_ranges = {merged_top_left(item): item for item in merged_ranges}
        cached_ws = (
            cached_wb[name] if cached_wb is not None and name in cached_wb.sheetnames else None
        )
        scan = facts.scans.get(fact.part) if scanned and fact is not None else None
        cells: dict[str, CellValue] = {}
        #: Each array formula's master coordinate -> the extent its result was saved over.
        array_extents: dict[str, str] = {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                data_type = cell.data_type or "n"
                formula: str | None = None
                extent: str | None = None
                if data_type == "f":
                    formula, extent = _formula_text(cell.value)
                    if extent is not None:
                        array_extents[cell.coordinate] = extent
                number_format = cell.number_format or "General"
                saved: Any = None
                if formula is not None and scan is not None:
                    saved = cached_value(
                        scan.cached.get(cell.coordinate), number_format, formula_wb.epoch
                    )
                elif formula is not None and cached_ws is not None:
                    saved = cached_ws[cell.coordinate].value
                cells[cell.coordinate] = CellValue(
                    sheet_name=name,
                    coordinate=cell.coordinate,
                    row=cell.row,
                    column=cell.column,
                    value=cell.value if formula is None else None,
                    formula=formula,
                    cached_value=saved,
                    data_type=data_type,
                    number_format=number_format,
                    is_date=bool(cell.is_date),
                    is_percentage="%" in number_format,
                    is_currency=bool(_CURRENCY.search(number_format)),
                    merged_range=master_ranges.get(cell.coordinate),
                    array_range=extent,
                )
        _mark_array_extents(cells, array_extents)
        tables: list[RawTable] = []
        for table_name in list(ws.tables):
            table = ws.tables[table_name]
            columns = tuple(column.name or "" for column in (table.tableColumns or ()))
            tables.append(
                RawTable(
                    name=table.displayName or table_name,
                    ref=table.ref or "",
                    columns=columns,
                )
            )

        flagged = (
            declared_ref is not None
            and _dimension_area(declared_ref) > MAX_DECLARED_CELLS
            and _dimension_area(declared_ref) > 4 * max(1, len(cells))
        )
        sheets.append(
            RawSheet(
                name=name,
                visibility=fact.state if fact else ws.sheet_state,
                cells=cells,
                merged_ranges=merged_ranges,
                tables=tuple(tables),
                declared_dimension=declared_ref,
                declared_dimension_flagged=flagged,
                is_macro_sheet=False,
            )
        )

    defined_names: list[RawDefinedName] = [
        RawDefinedName(name=name, attr_text=dn.attr_text)
        for name, dn in formula_wb.defined_names.items()
    ]
    # openpyxl >= 3.1 keeps a sheet-scoped name on its worksheet, not in ``wb.defined_names``.
    for ws in formula_wb.worksheets:
        defined_names.extend(
            RawDefinedName(name=name, attr_text=dn.attr_text, scope_sheet=ws.title)
            for name, dn in ws.defined_names.items()
        )

    return RawWorkbook(
        path=location,
        source_file=label,
        source_sha256=digest,
        sheets=tuple(sheets),
        defined_names=tuple(defined_names),
        macro_sheet_names=tuple(sorted(macro_names)),
        has_vba=facts.has_vba,
        warnings=tuple(
            f"sheet {sheet.name!r} declared dimension {sheet.declared_dimension!r}; ignored"
            for sheet in sheets
            if sheet.declared_dimension_flagged
        ),
    )
