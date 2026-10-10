"""An attached workbook, for a conversation rather than an index.

:func:`render_workbook` turns an ``.xlsx``/``.xlsm`` (a path or an upload's bytes) into text for the
context window within a token budget; :class:`WorkbookSession` answers the tool calls -- read a
range, find a value, trace precedents and dependents -- that reach what the budget left out. Both
work on the in-memory :class:`~excel_rag.workbook.WorkbookModel`: no Elasticsearch, no embeddings,
nothing persisted.
"""

from __future__ import annotations

from .render import (
    DETAIL_LADDER,
    FORMULA_MARK,
    Detail,
    RenderedWorkbook,
    TokenCounter,
    approx_tokens,
    formula_line,
    render_workbook,
)
from .tools import (
    CALCULATE_DEFINITION,
    MAX_FORMULAS,
    MAX_MATCHES,
    MAX_RANGE_CELLS,
    TOOL_DEFINITIONS,
    ToolInputError,
    ToolOutcome,
    WorkbookSession,
)

__all__ = [
    "CALCULATE_DEFINITION",
    "DETAIL_LADDER",
    "FORMULA_MARK",
    "MAX_FORMULAS",
    "MAX_MATCHES",
    "MAX_RANGE_CELLS",
    "TOOL_DEFINITIONS",
    "Detail",
    "RenderedWorkbook",
    "TokenCounter",
    "ToolInputError",
    "ToolOutcome",
    "WorkbookSession",
    "approx_tokens",
    "formula_line",
    "render_workbook",
]
