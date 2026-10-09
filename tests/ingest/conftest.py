"""Shared helpers: build a fixture workbook into ``tmp_path`` and ingest it."""

from __future__ import annotations

from pathlib import Path

import pytest
from fixtures import make_fixtures as mk

from excel_rag.ingest import IngestedWorkbook, ingest_workbook


@pytest.fixture
def build(tmp_path: Path) -> Builder:
    return Builder(tmp_path)


class Builder:
    """A tiny facade so a test reads ``build(load, "wb", 1)`` instead of repeating paths."""

    def __init__(self, directory: Path) -> None:
        self.directory = directory

    def path(self, name: str) -> Path:
        return getattr(mk, name)(self.directory)

    def ingest(self, name: str, workbook_id: str = "wb", version: int = 1) -> IngestedWorkbook:
        return ingest_workbook(self.path(name), workbook_id=workbook_id, version=version)
