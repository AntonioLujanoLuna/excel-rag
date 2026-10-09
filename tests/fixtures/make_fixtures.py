"""Generate the acceptance-test workbooks at test time -- **no binary .xlsx is committed**.

Every function writes an ``.xlsx``/``.xlsm`` into the given directory and returns its path. The few
that need an encrypted/macro/hostile package build it by rewriting the zip openpyxl produced, which
is exactly the low-level shape the reader must survive.
"""

from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table

_MACRO_SHEET_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>'
)


def _rewrite_zip(
    source: Path,
    target: Path,
    *,
    edits: dict[str, bytes | object] | None = None,
    adds: dict[str, bytes] | None = None,
) -> Path:
    """Copy a workbook zip, replacing named parts (bytes or ``bytes -> bytes`` callable)."""
    edits = edits or {}
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.namelist():
            data = zin.read(item)
            if item in edits:
                editor = edits[item]
                data = editor(data) if callable(editor) else bytes(editor)  # type: ignore[arg-type]
            zout.writestr(item, data)
        for name, data in (adds or {}).items():
            zout.writestr(name, data)
    return target


def two_tables_one_sheet(directory: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report"
    ws["A1"] = "Quarterly Report"
    ws.merge_cells("A1:C1")
    ws["A3"], ws["B3"], ws["C3"] = "Region", "Q1", "Q2"
    ws["A4"], ws["B4"], ws["C4"] = "North", 100, 120
    ws["A5"], ws["B5"], ws["C5"] = "South", 90, 110
    ws["A7"], ws["B7"] = "Product", "Units"
    ws["A8"], ws["B8"] = "Widget", 5
    ws["A9"], ws["B9"] = "Gadget", 7
    path = directory / "two_tables.xlsx"
    wb.save(path)
    return path


def merged_multilevel_header(directory: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Forecast"
    ws["A1"] = "Sales"
    ws.merge_cells("A1:B1")
    ws["C1"] = "Costs"
    ws.merge_cells("C1:D1")
    for col, label in zip("ABCD", ("Q1", "Q2", "Q1", "Q2"), strict=True):
        ws[f"{col}2"] = label
    for row, values in ((3, (10, 20, 5, 8)), (4, (11, 21, 6, 9))):
        for col, value in zip("ABCD", values, strict=True):
            ws[f"{col}{row}"] = value
    path = directory / "merged_header.xlsx"
    wb.save(path)
    return path


def units_and_notes(directory: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws["A1"], ws["B1"] = "Revenue", "Cost"
    ws["A2"], ws["B2"] = "USD", "USD"
    ws["A3"], ws["B3"] = 100, 50
    ws["A4"], ws["B4"] = 120, 60
    ws["A7"] = "Notes:"
    ws["A8"] = "Revenue is reported in constant currency."
    path = directory / "units_notes.xlsx"
    wb.save(path)
    return path


def cross_sheet_formula(directory: Path) -> Path:
    wb = openpyxl.Workbook()
    forecast = wb.active
    forecast.title = "Forecast"
    forecast["A1"], forecast["B1"] = "Metric", "Value"
    forecast["A2"] = "Revenue"
    forecast["B2"] = "=SUM(Actuals!D2:D500)*(1+Assumptions!C7)"
    actuals = wb.create_sheet("Actuals")
    actuals["A1"], actuals["D1"] = "id", "amount"
    for row, value in enumerate((10, 20, 30, 40, 50), start=2):
        actuals[f"A{row}"] = f"a{row}"
        actuals[f"D{row}"] = value
    assumptions = wb.create_sheet("Assumptions")
    assumptions["C7"] = 0.05
    path = directory / "cross_sheet.xlsx"
    wb.save(path)
    return path


def named_range(directory: Path) -> Path:
    wb = openpyxl.Workbook()
    assumptions = wb.active
    assumptions.title = "Assumptions"
    assumptions["C7"] = 0.05
    wb.defined_names.add(DefinedName("GrowthRate", attr_text="Assumptions!$C$7"))
    calc = wb.create_sheet("Calc")
    calc["A1"], calc["B1"] = "Base", "Growth"
    calc["A2"], calc["B2"] = 100, "=100*GrowthRate"
    path = directory / "named_range.xlsx"
    wb.save(path)
    return path


def indirect_offset(directory: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Calc"
    ws["A1"], ws["B1"] = "x", "y"
    ws["A2"], ws["B2"] = 1, 2
    ws["A3"] = '=INDIRECT("A2")'
    ws["A4"] = "=OFFSET(A3,1,0)"
    ws["A5"] = "=SUM(A2:A4)"
    path = directory / "indirect_offset.xlsx"
    wb.save(path)
    return path


def large_region(directory: Path, rows: int = 500, columns: int = 6) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    headers = [f"col_{index}" for index in range(columns)]
    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col, value=header)
    for row in range(2, rows + 2):
        for col in range(1, columns + 1):
            ws.cell(row=row, column=col, value=(row - 2) * columns + col)
    path = directory / f"large_{rows}x{columns}.xlsx"
    wb.save(path)
    return path


def table_object(directory: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws.append(["Region", "Revenue"])
    ws.append(["North", 100])
    ws.append(["South", 120])
    ws.append(["East", 90])
    table = Table(displayName="SalesTable", ref="A1:B4")
    ws.add_table(table)
    path = directory / "table_object.xlsx"
    wb.save(path)
    return path


def _inject_cached_value(path: Path, coordinate: str, value: str, target: Path) -> Path:
    """Rewrite a formula cell's ``<v/>`` into a cached ``<v>value</v>`` in the sheet XML."""
    with zipfile.ZipFile(path) as archive:
        name = next(item for item in archive.namelist() if item.startswith("xl/worksheets/"))
        sheet_xml = archive.read(name)

    pattern = re.compile(
        (rb'<c r="' + coordinate.encode() + rb'"[^>]*>.*?<f>[^<]*</f>\s*<v\s*/>'),
        re.DOTALL,
    )
    replaced, count = pattern.subn(
        lambda match: match.group(0)[: -len(b"<v />")].rstrip() + b"<v>" + value.encode() + b"</v>",
        sheet_xml,
    )
    if count != 1:
        raise AssertionError(f"could not inject a cached value into {coordinate}: {count} matches")
    return _rewrite_zip(path, target, edits={name: replaced})


def cached_value(directory: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Calc"
    ws["A1"], ws["B1"], ws["C1"] = 2, 3, "=A1+B1"
    path = directory / "cached.xlsx"
    wb.save(path)
    return _inject_cached_value(path, "C1", "999", directory / "cached_value.xlsx")


def hostile_dimension(directory: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"], ws["B1"] = "Region", "Amount"
    ws["A2"], ws["B2"] = "North", 100
    path = directory / "hostile_base.xlsx"
    wb.save(path)
    with zipfile.ZipFile(path) as archive:
        name = next(item for item in archive.namelist() if item.startswith("xl/worksheets/"))
    return _rewrite_zip(
        path,
        directory / "hostile_dimension.xlsx",
        edits={
            name: lambda data: re.sub(
                rb'ref="[A-Z0-9]+:[A-Z]+[0-9]+"', b'ref="A1:XFD1048576"', data
            )
        },
    )


def macro_workbook(directory: Path) -> Path:
    """A workbook carrying a vbaProject part and a macro sheet, plus a formula that reads it."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"], ws["B1"] = "Metric", "Value"
    ws["A2"], ws["B2"] = "x", "=Macro1!A1"
    base = directory / "macro_base.xlsx"
    wb.save(base)
    return _rewrite_zip(
        base,
        directory / "macro.xlsm",
        edits={
            "xl/workbook.xml": lambda data: data.replace(
                b"</sheets>",
                b'<sheet name="Macro1" sheetId="77" r:id="rIdMacro"/></sheets>',
            ),
            "xl/_rels/workbook.xml.rels": lambda data: data.replace(
                b"</Relationships>",
                b'<Relationship Id="rIdMacro" '
                b'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                b'relationships/xlMacrosheet" '
                b'Target="macrosheets/sheet1.xml"/></Relationships>',
            ),
            "[Content_Types].xml": lambda data: data.replace(
                b"</Types>",
                b'<Override PartName="/xl/macrosheets/sheet1.xml" '
                b'ContentType="application/vnd.ms-excel.macrosheet+xml"/>'
                b'<Override PartName="/xl/vbaProject.bin" '
                b'ContentType="application/vnd.ms-office.vbaProject"/></Types>',
            ),
        },
        adds={
            "xl/macrosheets/sheet1.xml": _MACRO_SHEET_XML,
            "xl/vbaProject.bin": b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1FAKE-VBA-PROJECT",
        },
    )


def corrupt_file(directory: Path) -> Path:
    path = directory / "not_a_workbook.xlsx"
    path.write_bytes(b"this is definitely not a zip archive")
    return path


def copy_as(source: Path, directory: Path, name: str) -> Path:
    target = directory / name
    shutil.copyfile(source, target)
    return target
