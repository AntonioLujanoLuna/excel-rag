"""The ``examples/ask_workbook.py`` loop, against a scripted client (no API call is made).

What is checked is the plumbing a real conversation depends on: the workbook is sent once behind a
cache breakpoint with the tool definitions, every ``tool_use`` in a turn is answered in a single
user message with matching ids, a bad call comes back as an error result, and the loop ends on the
final answer, a refusal, or the turn cap.
"""

from __future__ import annotations

import importlib.util
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest

from excel_rag.workbook.calc import calculation_available

_SPEC = importlib.util.spec_from_file_location(
    "ask_workbook", Path(__file__).parents[2] / "examples" / "ask_workbook.py"
)
assert _SPEC is not None and _SPEC.loader is not None
example = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(example)


def _tool_use(identifier: str, name: str, tool_input: dict[str, Any]) -> SimpleNamespace:
    return SimpleNamespace(type="tool_use", id=identifier, name=name, input=tool_input)


def _text(text: str) -> SimpleNamespace:
    return SimpleNamespace(type="text", text=text)


class _ScriptedClient:
    """Plays back responses and records every request, shaped like ``client.beta.messages``."""

    def __init__(self, responses: list[SimpleNamespace]) -> None:
        self.responses = responses
        self.requests: list[dict[str, Any]] = []
        self.beta = SimpleNamespace(messages=SimpleNamespace(create=self._create))

    def _create(self, **request: Any) -> SimpleNamespace:
        # Snapshot the message list: the loop keeps appending to the same object.
        self.requests.append({**request, "messages": list(request["messages"])})
        return self.responses.pop(0)


def test_tool_calls_are_answered_then_the_answer_is_returned(fixtures) -> None:
    client = _ScriptedClient(
        [
            SimpleNamespace(
                stop_reason="tool_use",
                content=[
                    _text("Tracing the formula."),
                    _tool_use("toolu_a", "precedents", {"sheet": "Forecast", "range": "B2"}),
                    _tool_use("toolu_b", "read_range", {"sheet": "Nope", "range": "A1"}),
                ],
            ),
            SimpleNamespace(stop_reason="end_turn", content=[_text("Forecast!B2 sums D2:D500.")]),
        ]
    )
    answer = example.ask(client, fixtures.cross_sheet_formula(), "How is revenue calculated?")
    assert answer == "Forecast!B2 sums D2:D500."

    first, second = client.requests
    assert first["model"] == "claude-opus-5-5"
    assert first["fallbacks"] == "default"
    assert [tool["name"] for tool in first["tools"]] == [
        "read_range",
        "find",
        "precedents",
        "dependents",
        *(["calculate"] if calculation_available() else []),
    ]
    workbook_block = first["messages"][0]["content"][0]
    assert workbook_block["text"].startswith("<workbook>\n# Workbook: ")
    assert workbook_block["cache_control"] == {"type": "ephemeral"}

    results = second["messages"][-1]
    assert results["role"] == "user"
    assert [block["tool_use_id"] for block in results["content"]] == ["toolu_a", "toolu_b"]
    assert "Assumptions!C7 = 0.05" in results["content"][0]["content"]
    assert results["content"][1]["is_error"] is True


def test_a_refusal_is_raised_not_returned_as_an_answer(fixtures) -> None:
    client = _ScriptedClient([SimpleNamespace(stop_reason="refusal", content=[])])
    with pytest.raises(RuntimeError, match="declined"):
        example.ask(client, fixtures.units_and_notes(), "q")


def test_running_out_of_tokens_is_not_an_answer(fixtures) -> None:
    client = _ScriptedClient([SimpleNamespace(stop_reason="max_tokens", content=[_text("Rev")])])
    with pytest.raises(RuntimeError, match="max_tokens"):
        example.ask(client, fixtures.units_and_notes(), "q")


def test_the_loop_is_capped(fixtures) -> None:
    looping = SimpleNamespace(
        stop_reason="tool_use", content=[_tool_use("toolu_x", "find", {"query": "Revenue"})]
    )
    client = _ScriptedClient([looping, looping, looping])
    with pytest.raises(RuntimeError, match="3 turns"):
        example.ask(client, fixtures.units_and_notes(), "q", max_turns=3)
