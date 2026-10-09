"""Ask Claude a question about an attached workbook: rendered into context, explored with tools.

    pip install anthropic excel-rag
    python examples/ask_workbook.py budget.xlsx "How is projected revenue calculated?"

The workbook is rendered to fit a token budget (``render_workbook``) and sent once, at the start
of the conversation, behind a prompt-cache breakpoint -- every later turn of the loop re-sends it,
so caching it is what keeps tool calls cheap. Whatever the budget left out, Claude reads with the
workbook tools (``WorkbookSession``): the definitions go in ``tools``, and each ``tool_use`` block
is answered with ``session.tool_result``.

This is a manual loop rather than the SDK's tool runner because the tools are plain JSON
definitions shared with any client, and their error results carry ``is_error``.

The request opts into server-side refusal fallbacks (``fallbacks="default"``): if a safety
classifier declines, the API re-runs the request on a fallback model instead of returning the
refusal. Drop the two fallback arguments to turn that off.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

from excel_rag.context import WorkbookSession, render_workbook

MODEL = "claude-opus-5-5"
FALLBACK_BETA = "server-side-fallback-2026-07-01"

SYSTEM = (
    "A spreadsheet is attached as text. Cite the cells you rely on as Sheet!A1. Values marked "
    "with ƒ are the results Excel last saved for formula cells; nothing was recalculated, so do "
    "not present them as freshly computed. Where the rendering marks rows, columns or formulas "
    "as omitted, read them with the workbook tools before answering from them, and use the "
    "precedents and dependents tools to explain how a value is calculated."
)


def ask(
    client: Any,
    source: str | Path | bytes,
    question: str,
    *,
    name: str | None = None,
    token_budget: int = 30_000,
    max_turns: int = 20,
) -> str:
    """Answer ``question`` about workbook ``source``; ``client`` is an ``anthropic.Anthropic``."""
    rendered = render_workbook(source, name=name, token_budget=token_budget, tools_hint=True)
    session = WorkbookSession(rendered.model)
    messages: list[dict[str, Any]] = [
        {
            "role": "user",
            "content": [
                {
                    "type": "text",
                    "text": f"<workbook>\n{rendered.text}</workbook>",
                    "cache_control": {"type": "ephemeral"},
                },
                {"type": "text", "text": question},
            ],
        }
    ]
    for _ in range(max_turns):
        response = client.beta.messages.create(
            model=MODEL,
            max_tokens=16_000,
            betas=[FALLBACK_BETA],
            fallbacks="default",
            system=SYSTEM,
            tools=session.tool_definitions(),
            messages=messages,
        )
        if response.stop_reason == "refusal":
            raise RuntimeError("the request was declined, and so was its fallback")
        # Append the whole content, not just its text: tool_use blocks must be echoed back.
        messages.append({"role": "assistant", "content": response.content})
        if response.stop_reason == "end_turn":
            return "".join(block.text for block in response.content if block.type == "text")
        if response.stop_reason != "tool_use":
            raise RuntimeError(f"stopped before answering: {response.stop_reason}")
        # Every tool_result answering this turn goes back in one user message.
        messages.append(
            {
                "role": "user",
                "content": [
                    session.tool_result(block.id, block.name, block.input)
                    for block in response.content
                    if block.type == "tool_use"
                ],
            }
        )
    raise RuntimeError(f"no answer within {max_turns} turns")


def main(argv: list[str]) -> int:
    if len(argv) != 2:
        print("usage: ask_workbook.py WORKBOOK QUESTION", file=sys.stderr)
        return 2
    import anthropic

    print(ask(anthropic.Anthropic(), Path(argv[0]), argv[1]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
