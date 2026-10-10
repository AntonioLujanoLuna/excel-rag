"""One streaming pass over a worksheet part: its declared dimension and its formulas' saved values.

openpyxl exposes a formula's text and the value Excel last saved for it under one load flag or the
other, never together, so the obvious reader loads every workbook twice. The second load builds a
whole worksheet model only to read one ``<v>`` per formula cell. This module reads exactly that,
with ``expat`` and no element tree, in a pass that also picks up ``<dimension ref>`` -- which the
reader would otherwise parse the whole sheet into a tree to find.

What it returns is *raw*: the cell type attribute and the ``<v>`` (or inline-string) text.
Turning it into a Python value needs the cell's number format, which the formula load already has,
so :func:`cached_value` does that conversion the way openpyxl's ``data_only`` load would. A
shared-string result (``t="s"``, which Excel does not write for formula cells but another producer might) is
reported as unsupported, and the reader falls back to openpyxl's load for that workbook.

Entity declarations are refused, as ``defusedxml`` would: a part that declares one raises
:class:`~excel_rag.workbook.errors.WorkbookError`.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import IO, Any
from xml.parsers import expat

from openpyxl.styles.numbers import (  # type: ignore[import-untyped]
    is_date_format,
    is_timedelta_format,
)
from openpyxl.utils.cell import coordinate_from_string  # type: ignore[import-untyped]
from openpyxl.utils.datetime import from_excel, from_ISO8601  # type: ignore[import-untyped]

from .errors import WorkbookError

_READ_CHUNK = 1 << 16


@dataclass(frozen=True, slots=True)
class RawCached:
    """A formula cell's saved result as written: its ``t`` attribute and its text."""

    data_type: str
    text: str | None


@dataclass(slots=True)
class SheetScan:
    """What one pass over a worksheet part found."""

    dimension: str | None = None
    cached: dict[str, RawCached] = field(default_factory=dict)
    #: A formula cell carried a shared-string result, which this scan does not resolve.
    needs_shared_strings: bool = False


def _column_letters(index: int) -> str:
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


class _Scanner:
    def __init__(self) -> None:
        self.scan = SheetScan()
        self._row = 0
        self._column = 0
        self._coordinate: str | None = None
        self._type = "n"
        self._has_formula = False
        self._text: list[str] | None = None
        self._value: str | None = None
        self._inline: list[str] | None = None

    @staticmethod
    def _local(name: str) -> str:
        return name.rsplit(" ", 1)[-1]

    def start(self, name: str, attrs: dict[str, str]) -> None:
        tag = self._local(name)
        if tag == "c":
            reference = attrs.get("r")
            if reference:
                letters, row = coordinate_from_string(reference)
                self._coordinate = f"{letters.upper()}{row}"
                self._row = row
                self._column = _column_index(letters)
            else:
                self._column += 1
                self._coordinate = f"{_column_letters(self._column)}{self._row}"
            self._type = attrs.get("t", "n")
            self._has_formula = False
            self._value = None
            self._inline = None
        elif tag == "row":
            reference = attrs.get("r")
            self._row = int(reference) if reference and reference.isdigit() else self._row + 1
            self._column = 0
        elif tag == "f" and self._coordinate is not None:
            self._has_formula = True
        elif tag == "v" and self._coordinate is not None:
            self._text = []
        elif tag == "is" and self._coordinate is not None:
            self._inline = []
        elif tag == "t" and self._inline is not None:
            self._text = []
        elif tag == "dimension" and self.scan.dimension is None:
            ref = attrs.get("ref")
            if ref:
                self.scan.dimension = ref

    def end(self, name: str) -> None:
        tag = self._local(name)
        if tag == "v" and self._text is not None:
            self._value = "".join(self._text)
            self._text = None
        elif tag == "t" and self._text is not None and self._inline is not None:
            self._inline.append("".join(self._text))
            self._text = None
        elif tag == "c" and self._coordinate is not None:
            if self._has_formula:
                if self._type == "inlineStr":
                    text = "".join(self._inline) if self._inline is not None else None
                else:
                    text = self._value or None
                if self._type == "s" and text is not None:
                    self.scan.needs_shared_strings = True
                self.scan.cached[self._coordinate] = RawCached(self._type, text)
            self._coordinate = None

    def characters(self, data: str) -> None:
        if self._text is not None:
            self._text.append(data)


def _column_index(letters: str) -> int:
    index = 0
    for char in letters.upper():
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index


def _forbid_entities(*_args: Any) -> None:
    raise WorkbookError("refusing workbook: a worksheet part declares XML entities")


def scan_sheet(stream: IO[bytes], part: str) -> SheetScan:
    """Stream one worksheet part.

    Raises :class:`WorkbookError` on malformed or entity-bearing XML.
    """
    scanner = _Scanner()
    parser = expat.ParserCreate(namespace_separator=" ")
    parser.SetParamEntityParsing(expat.XML_PARAM_ENTITY_PARSING_NEVER)
    parser.EntityDeclHandler = _forbid_entities
    parser.UnparsedEntityDeclHandler = _forbid_entities
    parser.ExternalEntityRefHandler = _forbid_entities  # type: ignore[assignment]
    parser.StartElementHandler = scanner.start
    parser.EndElementHandler = scanner.end
    parser.CharacterDataHandler = scanner.characters
    parser.buffer_text = True
    try:
        while chunk := stream.read(_READ_CHUNK):
            parser.Parse(chunk, False)
        parser.Parse(b"", True)
    except expat.ExpatError as exc:
        raise WorkbookError(f"not a valid workbook package: {part} is not XML ({exc})") from exc
    return scanner.scan


def cached_value(raw: RawCached | None, number_format: str, epoch: Any) -> Any:
    """The saved value as openpyxl's ``data_only`` load would give it; ``None`` if none was saved.

    A date-formatted number becomes a ``datetime`` (or a ``timedelta`` for a duration format), as
    openpyxl converts it; a serial outside the date range is ``#VALUE!``, as openpyxl reports it.
    """
    if raw is None or raw.text is None:
        return None
    text = raw.text
    data_type = raw.data_type
    if data_type == "n":
        number: int | float = (
            float(text) if ("." in text or "E" in text or "e" in text) else int(text)
        )
        if is_date_format(number_format):
            try:
                return from_excel(number, epoch, timedelta=is_timedelta_format(number_format))
            except (OverflowError, ValueError):
                return "#VALUE!"
        return number
    if data_type == "b":
        return bool(int(text))
    if data_type == "d":
        return from_ISO8601(text)
    # "str", "inlineStr" and "e" carry their text as is.
    return text


__all__ = ["RawCached", "SheetScan", "cached_value", "scan_sheet"]
