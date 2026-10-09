"""Structure-aware RAG over Excel workbooks.

An Excel workbook is treated like a codebase: worksheets resemble source files, tables and named
ranges resemble symbols, and formulas form a dependency graph. Both the semantics and that
structure are indexed in Elasticsearch, and one stateless endpoint answers a semantic search with
exact coordinates, values and statically resolved references.

Elasticsearch is the only persistent search/vector/metadata store. There is no graph database, no
SQL engine, no agent layer and no answer synthesis: the response is retrieval evidence, and
reasoning belongs to the caller.
"""

from __future__ import annotations

from importlib import import_module
from typing import Any

from .settings import Settings

__all__ = ["Settings", "__version__", "create_app"]

__version__ = "0.1.0"


def __getattr__(name: str) -> Any:
    """Expose ``create_app`` lazily so importing the package never imports the service.

    Resolved through ``importlib`` rather than a static import so ``import excel_rag`` stays cheap
    and mypy does not require the service modules to exist for a pure-models import. A missing
    ``excel_rag.app`` surfaces as ``AttributeError``, which is what ``getattr`` callers expect.
    """
    if name == "create_app":
        try:
            module = import_module(f"{__name__}.app")
        except ModuleNotFoundError as exc:
            raise AttributeError(f"module {__name__!r} has no attribute {name!r}") from exc
        return module.create_app
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


def __dir__() -> list[str]:
    return sorted(__all__)
