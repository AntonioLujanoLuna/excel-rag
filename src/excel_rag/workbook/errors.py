"""The one exception reading or modelling a workbook raises to a caller.

A malformed, encrypted or unreadable workbook must surface as a clear :class:`WorkbookError`, never
as a bare ``zipfile.BadZipFile`` or an openpyxl internal trace: the caller is a CLI, an ingest
worker or a chat attachment handler, and it needs one failure type it can catch and report.
"""

from __future__ import annotations


class WorkbookError(RuntimeError):
    """A workbook could not be read, modelled or indexed.

    Raised for an unreadable package, a hostile structure, or a workbook that would exceed the
    configured document budget. The message always names the file it concerns.
    """


#: The name ingestion has always raised; the same class.
IngestError = WorkbookError
