"""Evaluation cases: a question, the workbook it is about, and the rectangles that answer it.

A case file is JSON Lines, one case per line::

    {"id": "growth", "question": "What growth rate is assumed?", "workbook_id": "wb42",
     "expected": [{"sheet": "Assumptions", "a1": "A3:B7"}]}

``expected`` names where a correct answer's citation is: a hit is relevant when it is on that sheet
and its A1 range intersects the rectangle. A whole-workbook or whole-sheet summary is never counted
as relevant -- it intersects everything on its sheet and cites nothing.
"""

from __future__ import annotations

import json
from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path

from ..models import A1Range


@dataclass(frozen=True, slots=True)
class Expected:
    """A rectangle a correct answer cites."""

    sheet: str
    a1: str

    def bounds(self) -> A1Range:
        return A1Range.parse(self.sheet, self.a1)


@dataclass(frozen=True, slots=True)
class EvalCase:
    id: str
    question: str
    workbook_id: str
    expected: tuple[Expected, ...]


class CaseFileError(ValueError):
    """A case file line that is not a valid case."""


def parse_cases(lines: Iterable[str], *, source: str = "<cases>") -> tuple[EvalCase, ...]:
    """Parse JSON Lines into cases; blank lines and ``#`` comments are skipped."""
    cases: list[EvalCase] = []
    for number, line in enumerate(lines, start=1):
        text = line.strip()
        if not text or text.startswith("#"):
            continue
        try:
            raw = json.loads(text)
            expected = tuple(
                Expected(str(item["sheet"]), str(item["a1"])) for item in raw["expected"]
            )
            for item in expected:
                item.bounds()
            case = EvalCase(
                id=str(raw.get("id") or f"line-{number}"),
                question=str(raw["question"]),
                workbook_id=str(raw["workbook_id"]),
                expected=expected,
            )
        except (json.JSONDecodeError, KeyError, TypeError, ValueError) as exc:
            raise CaseFileError(f"{source}:{number}: not a valid case ({exc})") from exc
        if not case.expected:
            raise CaseFileError(f"{source}:{number}: a case needs at least one expected range")
        cases.append(case)
    return tuple(cases)


def load_cases(path: str | Path) -> tuple[EvalCase, ...]:
    source = Path(path)
    return parse_cases(source.read_text(encoding="utf-8").splitlines(), source=str(source))


__all__ = ["CaseFileError", "EvalCase", "Expected", "load_cases", "parse_cases"]
