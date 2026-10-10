"""An MCP server: the workbook tools for any MCP client (Claude Desktop, Claude Code, an agent).

``excel-rag mcp`` serves, over stdio, the same operations an attached workbook gets in a
conversation -- render it within a token budget, read a range, find a value, trace precedents and
dependents -- plus a diff of two versions, all on files under the directories it is given
(``--root``, default the working directory). With ``--search`` it also serves the indexed retrieval
endpoint's search as a tool, against the configured Elasticsearch.

Every tool is read-only: nothing is written, no macro runs, no formula is evaluated. A path outside
the roots is refused, so a model cannot be talked into reading files it was not given. Loaded
workbooks are kept in a small cache keyed by path, modification time and size, so a conversation
of tool calls parses each file once and an edited file is read again.

Needs the ``mcp`` extra (``pip install excel-rag[mcp]``).
"""

from __future__ import annotations

import threading
from collections import OrderedDict
from collections.abc import Callable, Sequence
from pathlib import Path
from typing import TYPE_CHECKING, Any, TypeVar

from .context import ToolInputError, WorkbookSession, render_workbook
from .workbook.calc import calculation_available
from .workbook.errors import WorkbookError

if TYPE_CHECKING:
    from .retrieval import RetrievalService

#: The suffixes a workbook path may have.
WORKBOOK_SUFFIXES = (".xlsx", ".xlsm", ".xlsb", ".csv")
#: How many parsed workbooks the server keeps.
CACHE_SIZE = 8
#: The most files ``list_workbooks`` names.
MAX_LISTED = 200

_T = TypeVar("_T")


class PathRefused(ValueError):
    """A path the server will not open: outside every root, missing, or not a workbook."""


class WorkbookCatalog:
    """Resolves tool paths against the allowed roots and caches parsed workbooks."""

    def __init__(self, roots: Sequence[str | Path], *, cache_size: int = CACHE_SIZE) -> None:
        resolved = [Path(root).expanduser().resolve() for root in roots]
        if not resolved:
            raise ValueError("at least one root directory is required")
        for root in resolved:
            if not root.is_dir():
                raise ValueError(f"not a directory: {root}")
        self.roots = tuple(resolved)
        self._cache: OrderedDict[tuple[Path, int, int], WorkbookSession] = OrderedDict()
        self._cache_size = cache_size
        self._lock = threading.Lock()

    def resolve(self, path: str) -> Path:
        """An existing workbook file under one of the roots, or :class:`PathRefused`.

        A relative path is tried against each root in order; an absolute one must lie inside a
        root after symlinks are resolved.
        """
        if not path or not path.strip():
            raise PathRefused("path is empty")
        given = Path(path).expanduser()
        candidates = [given] if given.is_absolute() else [root / given for root in self.roots]
        for candidate in candidates:
            resolved = candidate.resolve()
            if not any(resolved.is_relative_to(root) for root in self.roots):
                continue
            if resolved.is_file():
                if resolved.suffix.lower() not in WORKBOOK_SUFFIXES:
                    raise PathRefused(f"{path}: not an .xlsx/.xlsm/.xlsb/.csv workbook")
                return resolved
        if given.is_absolute() and not any(
            given.resolve().is_relative_to(root) for root in self.roots
        ):
            raise PathRefused(f"{path}: outside the directories this server may read")
        raise PathRefused(f"{path}: no such workbook under {', '.join(map(str, self.roots))}")

    def session(self, path: str) -> WorkbookSession:
        resolved = self.resolve(path)
        stat = resolved.stat()
        key = (resolved, stat.st_mtime_ns, stat.st_size)
        with self._lock:
            cached = self._cache.get(key)
            if cached is not None:
                self._cache.move_to_end(key)
                return cached
        session = WorkbookSession.load(resolved)
        with self._lock:
            for stale in [item for item in self._cache if item[0] == resolved]:
                del self._cache[stale]
            self._cache[key] = session
            while len(self._cache) > self._cache_size:
                self._cache.popitem(last=False)
        return session

    def list_workbooks(self) -> list[str]:
        """Workbook files under the roots, as paths relative to their root."""
        found: list[str] = []
        for root in self.roots:
            for candidate in sorted(root.rglob("*")):
                if candidate.suffix.lower() in WORKBOOK_SUFFIXES and candidate.is_file():
                    name = candidate.name
                    if name.startswith("~$"):  # Excel's lock file for an open workbook
                        continue
                    label = str(candidate.relative_to(root))
                    found.append(label if len(self.roots) == 1 else f"{root.name}/{label}")
                    if len(found) >= MAX_LISTED:
                        return found
        return found


def build_server(
    roots: Sequence[str | Path],
    *,
    search_service: Callable[[], RetrievalService] | None = None,
    calculate: bool | None = None,
) -> Any:
    """Build the MCP server over ``roots``; with ``search_service``, also the ``search_index``
    tool (the factory is called per search, so the service can be built lazily). ``calculate``
    (formula evaluation) is served when the ``calc`` extra is installed, unless it is ``False``."""
    try:
        from mcp.server.mcpserver import MCPServer
        from mcp.server.mcpserver.exceptions import ToolError
        from mcp_types import ToolAnnotations
    except ModuleNotFoundError as exc:  # pragma: no cover - exercised only without the extra
        raise RuntimeError("the MCP server needs the 'mcp' extra: install excel-rag[mcp]") from exc

    catalog = WorkbookCatalog(roots)
    read_only = ToolAnnotations(readOnlyHint=True, destructiveHint=False, openWorldHint=False)
    offer_calculate = calculation_available() if calculate is None else calculate
    recalculation = (
        "Formula cells show the value Excel last saved; calculate computes cells under changed "
        "inputs (a what-if), marking what it recomputed. Files are never modified"
        if offer_calculate
        else "Formula cells show the value Excel last saved; nothing is recalculated"
    )
    server = MCPServer(
        "excel-rag",
        instructions=(
            "Tools over Excel workbooks in the directories this server was given. Start with "
            "list_workbooks or render_workbook for an overview, then read_range, find, "
            f"precedents and dependents for detail. {recalculation}, and no macro runs. Cite "
            "cells as Sheet!A1."
        ),
    )

    def answer(call: Callable[[], _T]) -> _T:
        try:
            return call()
        except (PathRefused, ToolInputError, WorkbookError) as error:
            raise ToolError(str(error)) from error

    @server.tool(annotations=read_only)
    def list_workbooks() -> str:
        """List the .xlsx/.xlsm/.xlsb/.csv workbooks this server may read, as paths to pass to
        the other tools."""
        names = catalog.list_workbooks()
        if not names:
            return "No workbooks under " + ", ".join(map(str, catalog.roots)) + "."
        suffix = f"\n(First {MAX_LISTED} shown.)" if len(names) >= MAX_LISTED else ""
        return "\n".join(names) + suffix

    @server.tool(name="render_workbook", annotations=read_only)
    def render(path: str, token_budget: int = 8000) -> str:
        """Render a workbook as markdown within a token budget: sheets, regions with headers and
        units, grids with row numbers and column letters, formulas with their saved values and
        what they read. Omitted rows and columns are marked; read them with read_range."""

        def run() -> str:
            if token_budget < 500:
                raise ToolInputError("token_budget must be at least 500")
            session = catalog.session(path)
            return render_workbook(session.model, token_budget=token_budget, tools_hint=True).text

        return answer(run)

    @server.tool(annotations=read_only)
    def read_range(path: str, sheet: str, range: str) -> str:
        """Read a rectangle (e.g. A1:F40) of one sheet as a grid, with the formulas inside it.
        At most 2,000 cells; a larger range is cut by rows and the reply names the rest."""
        return answer(lambda: catalog.session(path).read_range(sheet, range))

    @server.tool(annotations=read_only)
    def find(path: str, query: str) -> str:
        """Find cells (by value or formula text) and named ranges containing the text, across
        every sheet, case-insensitively. At most 50 matches."""
        return answer(lambda: catalog.session(path).find(query))

    @server.tool(annotations=read_only)
    def precedents(path: str, sheet: str, range: str) -> str:
        """List the formulas inside a range with their saved values and the cells and ranges
        they read, resolved statically (INDIRECT and external links are reported, not
        guessed), and any chart, pivot table or validation in the range with what it reads."""
        return answer(lambda: catalog.session(path).precedents(sheet, range))

    @server.tool(annotations=read_only)
    def dependents(path: str, sheet: str, range: str) -> str:
        """List the formulas, charts, pivot tables and data validations anywhere in the workbook
        that read any cell of the range: what changes if this input changes."""
        return answer(lambda: catalog.session(path).dependents(sheet, range))

    if offer_calculate:

        @server.tool(name="calculate", annotations=read_only)
        def calculate_cells(
            path: str, sheet: str, range: str, changes: dict[str, str] | None = None
        ) -> str:
            """Compute a range's cells, optionally after changing inputs: changes maps a cell
            (Assumptions!B4) to its value as typed (0.07, 7%, TRUE, text). Formulas no change
            reaches keep Excel's saved value; recomputed ones are marked, and what cannot be
            computed (INDIRECT, OFFSET, circular references) is reported unknown with the
            reason. Nothing is written to the file; every call starts from it as saved."""
            return answer(lambda: catalog.session(path).calculate(sheet, range, changes or {}))

    @server.tool(name="diff_workbooks", annotations=read_only)
    def diff_versions(before: str, after: str, max_changes: int = 200) -> str:
        """Compare two versions of a workbook: sheets added or removed, each changed cell (input
        value, formula, or saved result), what a changed formula now reads, and which formulas
        read a changed input."""
        from .workbook.diff import diff_workbooks, format_diff

        def run() -> str:
            if max_changes < 1:
                raise ToolInputError("max_changes must be at least 1")
            old, new = catalog.session(before), catalog.session(after)
            return format_diff(diff_workbooks(old.model, new.model, max_changes=max_changes))

        return answer(run)

    if search_service is not None:
        factory = search_service

        @server.tool(annotations=read_only)
        def search_index(
            query: str,
            workbook_ids: list[str] | None = None,
            top_k: int = 10,
            expand_references: bool = True,
        ) -> str:
            """Search the indexed workbooks (the retrieval service's hybrid search) and return
            the hits with their workbook, sheet and A1 range, plus the structural nodes they
            reference. Evidence for an answer, not an answer."""
            from .models import SearchFilters, SearchRequest
            from .retrieval import UnknownWorkbook

            try:
                request = SearchRequest(
                    query=query,
                    filters=SearchFilters(workbook_ids=tuple(workbook_ids or ())),
                    top_k=top_k,
                    expand_references=expand_references,
                )
            except ValueError as error:
                raise ToolError(f"invalid search: {error}") from error
            try:
                response = factory().search(request)
            except UnknownWorkbook as error:
                raise ToolError(f"no active version for workbook {error.workbook_id!r}") from error
            return response.model_dump_json(indent=1, exclude_none=True)

    return server


def default_search_service() -> RetrievalService:
    """The retrieval service the settings configure, as ``excel-rag serve`` would build it."""
    from .app import create_client
    from .embedding import build_embedder
    from .rerank import build_reranker
    from .retrieval import Repository, RetrievalService
    from .settings import Settings

    settings = Settings()
    repository = Repository(create_client(settings), settings)
    repository.ensure_indices()
    return RetrievalService(
        repository, settings, build_embedder(settings.embedding), build_reranker(settings.rerank)
    )


def run(roots: Sequence[str | Path], *, search: bool = False) -> None:  # pragma: no cover
    """Serve over stdio until the client disconnects."""
    service: RetrievalService | None = None
    lock = threading.Lock()

    def lazy() -> RetrievalService:
        nonlocal service
        with lock:
            if service is None:
                service = default_search_service()
            return service

    build_server(roots, search_service=lazy if search else None).run("stdio")


__all__ = [
    "CACHE_SIZE",
    "MAX_LISTED",
    "WORKBOOK_SUFFIXES",
    "PathRefused",
    "WorkbookCatalog",
    "build_server",
    "default_search_service",
    "run",
]
