"""The one exception ingestion raises to a caller.

A malformed, encrypted or unreadable workbook must surface as a clear :class:`IngestError`, never as
a bare ``zipfile.BadZipFile`` or an openpyxl internal trace: the caller is a CLI or an ingest
worker, and it needs one failure type it can catch and report.
"""

from __future__ import annotations


class IngestError(RuntimeError):
    """A workbook could not be read, modelled or indexed.

    Raised for an unreadable package, a hostile structure, or a workbook that would exceed the
    configured document budget. The message always names the file it concerns.
    """
