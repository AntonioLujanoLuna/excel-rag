"""Workbook formats other than ``.xlsx``/``.xlsm``: CSV and binary ``.xlsb``.

Both are read into the same :class:`~excel_rag.workbook.reader.RawSheet` the openpyxl reader
produces, so region detection, documents, rendering and the tools work on them unchanged. What
each format cannot carry is said, not invented:

* **CSV** is one sheet of values. Numbers and ``TRUE``/``FALSE`` are typed; everything else,
  including text that looks like a date or a formula (``=SUM(A1:A3)``), stays text -- a CSV has
  no formulas, and guessing a date format is how ``3/4`` becomes the wrong month. The delimiter is
  sniffed (comma, semicolon, tab, pipe), and in a semicolon-delimited file ``1200,5`` is the
  decimal it means there; the encoding is UTF-8 (with or without a BOM), falling back to
  Windows-1252. A number with a leading zero (``007``) is an identifier and stays text.
* **.xlsb** is read with ``pyxlsb`` (the ``xlsb`` extra), record by record. It yields every cell's
  value, and which cells are formulas -- but not the formula text, nor number formats, merges,
  tables or defined names. A formula cell therefore carries only its last-saved value, labelled as
  computed, and the workbook's warnings say what was not read. The declared dimension is never
  used, so a hostile one allocates nothing.
"""

from __future__ import annotations

import csv
import io
import re
import zipfile
from pathlib import PurePosixPath
from typing import IO, Any

from defusedxml.ElementTree import fromstring  # type: ignore[import-untyped]

from .canonical import CellValue, column_letter
from .errors import WorkbookError

#: The most rows and columns a CSV may have before it is refused (Excel's own sheet limits).
MAX_CSV_ROWS = 1_048_576
MAX_CSV_COLUMNS = 16_384

_NUMBER = re.compile(r"[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?")
_INVALID_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")

#: BIFF12 error codes, as Excel displays them.
_ERRORS = {
    0x00: "#NULL!",
    0x07: "#DIV/0!",
    0x0F: "#VALUE!",
    0x17: "#REF!",
    0x1D: "#NAME?",
    0x24: "#NUM!",
    0x2A: "#N/A",
    0x2B: "#GETTING_DATA",
}


def sheet_name_for(label: str) -> str:
    """A valid worksheet name from a file name: its stem, without ``[]:*?/\\``, at most 31 chars."""
    stem = PurePosixPath(label.replace("\\", "/")).stem or "Sheet1"
    cleaned = _INVALID_SHEET_CHARS.sub("_", stem).strip("'") or "Sheet1"
    return cleaned[:31]


def _decode(data: bytes) -> str:
    try:
        return data.decode("utf-8-sig")
    except UnicodeDecodeError:
        return data.decode("cp1252", errors="replace")


_COMMA_DECIMAL = re.compile(r"[+-]?\d+,\d+")


def _typed(text: str, *, comma_decimal: bool = False) -> tuple[Any, str]:
    """A CSV field as a value and openpyxl data type: number, boolean, or text.

    ``comma_decimal`` (a semicolon-delimited file, the European convention) reads ``1200,5`` as
    1200.5; elsewhere a comma is a thousands separator or text, and the field stays text.
    """
    stripped = text.strip()
    if comma_decimal and _COMMA_DECIMAL.fullmatch(stripped):
        return float(stripped.replace(",", ".")), "n"
    if _NUMBER.fullmatch(stripped):
        is_integer = stripped.lstrip("+-").isdigit()
        # Leading zeros are an identifier (a ZIP code, an account), not a number.
        if is_integer and len(stripped.lstrip("+-")) > 1 and stripped.lstrip("+-")[0] == "0":
            return text, "s"
        return (int(stripped) if is_integer else float(stripped)), "n"
    upper = stripped.upper()
    if upper in ("TRUE", "FALSE"):
        return upper == "TRUE", "b"
    return text, "s"


def read_csv_cells(data: bytes, sheet_name: str) -> dict[str, CellValue]:
    """Parse CSV bytes into populated cells on one sheet."""
    text = _decode(data)
    sample = text[:65_536]
    try:
        dialect: Any = csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    comma_decimal = getattr(dialect, "delimiter", ",") == ";"
    cells: dict[str, CellValue] = {}
    try:
        for row_index, row in enumerate(csv.reader(io.StringIO(text), dialect), start=1):
            if row_index > MAX_CSV_ROWS:
                raise WorkbookError(f"refusing CSV: more than {MAX_CSV_ROWS} rows")
            if len(row) > MAX_CSV_COLUMNS:
                raise WorkbookError(f"refusing CSV: row {row_index} has more than 16,384 fields")
            for column_index, field in enumerate(row, start=1):
                if field == "":
                    continue
                value, data_type = _typed(field, comma_decimal=comma_decimal)
                coordinate = f"{column_letter(column_index)}{row_index}"
                cells[coordinate] = CellValue(
                    sheet_name=sheet_name,
                    coordinate=coordinate,
                    row=row_index,
                    column=column_index,
                    value=value,
                    data_type=data_type,
                )
    except csv.Error as exc:
        raise WorkbookError(f"not a readable CSV: {exc}") from exc
    return cells


def looks_like_xlsb(data: bytes) -> bool:
    if not data.startswith(b"PK"):
        return False
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            return "xl/workbook.bin" in archive.namelist()
    except zipfile.BadZipFile:
        return False


def _count_records(stream: IO[bytes], record_id: int) -> int:
    """How many records of one id a BIFF12 part holds (pyxlsb's reader skips unhandled ids)."""
    data = stream.read()
    position, count = 0, 0
    while position < len(data):
        ident = 0
        for shift in range(4):
            if position >= len(data):
                return count
            byte = data[position]
            position += 1
            ident += byte << (8 * shift)
            if not byte & 0x80:
                break
        length = 0
        for shift in range(4):
            if position >= len(data):
                return count
            byte = data[position]
            position += 1
            length += (byte & 0x7F) << (7 * shift)
            if not byte & 0x80:
                break
        position += length
        count += ident == record_id
    return count


def read_xlsb_sheets(
    archive: zipfile.ZipFile,
) -> tuple[list[tuple[str, dict[str, CellValue]]], list[str]]:
    """Each sheet's name and populated cells, and the warnings for what ``.xlsb`` did not give."""
    try:
        from pyxlsb import biff12  # type: ignore[import-untyped]
        from pyxlsb.reader import BIFF12Reader  # type: ignore[import-untyped]
        from pyxlsb.stringtable import StringTable  # type: ignore[import-untyped]
    except ModuleNotFoundError as exc:
        raise WorkbookError(
            "reading .xlsb needs the 'xlsb' extra: install excel-rag[xlsb]"
        ) from exc

    names = set(archive.namelist())
    relationships: dict[str, str] = {}
    if "xl/_rels/workbook.bin.rels" in names:
        for node in fromstring(archive.read("xl/_rels/workbook.bin.rels")):
            relationships[node.attrib.get("Id", "")] = node.attrib.get("Target", "")

    sheets: list[tuple[str, str]] = []
    with archive.open("xl/workbook.bin") as stream:
        workbook_part = stream.read()
    for recid, record in BIFF12Reader(fp=io.BytesIO(workbook_part)):
        if recid == biff12.SHEET and record is not None:
            target = relationships.get(record.rId, "")
            part = target.lstrip("/") if target.startswith("/") else f"xl/{target}"
            sheets.append((record.name, part))
        elif recid == biff12.SHEETS_END:
            break
    defined_names = _count_records(io.BytesIO(workbook_part), biff12.DEFINEDNAME)

    strings: Any = None
    if "xl/sharedStrings.bin" in names:
        with archive.open("xl/sharedStrings.bin") as stream:
            strings = StringTable(fp=io.BytesIO(stream.read()))

    formula_records = {
        biff12.FORMULA_STRING,
        biff12.FORMULA_FLOAT,
        biff12.FORMULA_BOOL,
        biff12.FORMULA_BOOLERR,
    }
    value_records = formula_records | {
        biff12.NUM,
        biff12.BOOLERR,
        biff12.BOOL,
        biff12.FLOAT,
        biff12.STRING,
    }
    result: list[tuple[str, dict[str, CellValue]]] = []
    formula_cells = 0
    for name, part in sheets:
        cells: dict[str, CellValue] = {}
        if part not in names:
            result.append((name, cells))
            continue
        row_number = 0
        in_data = False
        with archive.open(part) as stream:
            for recid, record in BIFF12Reader(fp=io.BytesIO(stream.read())):
                if recid == biff12.SHEETDATA:
                    in_data = True
                elif recid == biff12.SHEETDATA_END:
                    break
                elif not in_data:
                    continue
                elif recid == biff12.ROW and record is not None:
                    row_number = record.r + 1
                elif recid in value_records and record is not None and record.v is not None:
                    value: Any = record.v
                    data_type = "n"
                    if recid == biff12.STRING:
                        value = strings[value] if strings is not None else None
                        data_type = "s"
                    elif recid == biff12.FORMULA_STRING:
                        data_type = "s"
                    elif recid in (biff12.BOOL, biff12.FORMULA_BOOL):
                        data_type = "b"
                    elif recid in (biff12.BOOLERR, biff12.FORMULA_BOOLERR):
                        value = _ERRORS.get(int(str(value), 16), "#VALUE!")
                        data_type = "e"
                    if isinstance(value, float) and value.is_integer():
                        value = int(value)
                    if value is None or row_number < 1:
                        continue
                    column = record.c + 1
                    coordinate = f"{column_letter(column)}{row_number}"
                    computed = recid in formula_records
                    formula_cells += computed
                    cells[coordinate] = CellValue(
                        sheet_name=name,
                        coordinate=coordinate,
                        row=row_number,
                        column=column,
                        value=None if computed else value,
                        cached_value=value if computed else None,
                        data_type=data_type,
                        formula_unavailable=computed,
                    )
        result.append((name, cells))

    warnings = ["xlsb: number formats are not read, so dates appear as Excel serial numbers"]
    if formula_cells:
        warnings.append(
            f"xlsb: formula text is not read; {formula_cells} formula cell(s) carry their "
            "last-saved values only, with no precedents"
        )
    if defined_names:
        warnings.append(f"xlsb: {defined_names} defined name(s) not read")
    return result, warnings


__all__ = [
    "MAX_CSV_COLUMNS",
    "MAX_CSV_ROWS",
    "looks_like_xlsb",
    "read_csv_cells",
    "read_xlsb_sheets",
    "sheet_name_for",
]
