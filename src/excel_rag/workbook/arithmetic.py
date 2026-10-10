"""A fast path for the most common formula: arithmetic over cells and numbers.

``=B1+A2*(1+$D$1)`` filled down 50,000 rows is the shape of most large models, and the general
evaluator spends most of half a millisecond per cell in dispatch. A formula made only of numbers,
cell references, ``+ - * / ^``, unary signs, ``%`` and parentheses is compiled here instead into a
small stack program over its tokens -- nothing is passed to ``eval`` -- with Excel's precedence
(negation, then ``%``, then ``^`` left to right, then ``* /``, then ``+ -``) and Excel's results
where they differ from Python's: ``0^0`` and a negative number to a fractional power are
``#NUM!``, ``0`` to a negative power and ``x/0`` are ``#DIV/0!``, an overflow is ``#NUM!``.

It takes numbers, booleans (1 and 0), empty cells (0) and error values (which propagate, the
first met first). Anything else -- text, which Excel may coerce by locale rules, or an array --
raises :class:`Fallback`, and the general evaluator answers instead.
"""

from __future__ import annotations

import math
from collections.abc import Callable, Sequence
from typing import Any

from openpyxl.formula.tokenizer import (  # type: ignore[import-untyped]
    Token,
    Tokenizer,
    TokenizerError,
)

#: Binding strength of each operator; negation binds tightest, as in Excel (``-2^2`` is 4).
_PRECEDENCE = {"neg": 5, "pos": 5, "%": 4, "^": 3, "*": 2, "/": 2, "+": 1, "-": 1}


class Fallback(Exception):
    """An input the fast path does not decide; the general evaluator does."""


class ExcelError(Exception):
    """An error value the arithmetic produced or met, by its Excel spelling."""

    def __init__(self, code: str) -> None:
        super().__init__(code)
        self.code = code


Program = list[tuple[str, Any]]


def compile_arithmetic(template: str, names: Sequence[str]) -> Callable[..., float] | None:
    """A function computing ``template`` from the values of ``names`` (its placeholders, in
    argument order), or ``None`` when the formula is more than arithmetic."""
    try:
        tokens = Tokenizer(template).items
    except TokenizerError:
        return None
    positions = {name.upper(): index for index, name in enumerate(names)}
    program: Program = []
    operators: list[str] = []
    expect_operand = True
    for token in tokens:
        kind, subtype, value = token.type, token.subtype, token.value
        if kind == Token.WSPACE:
            continue
        if kind == Token.OPERAND and subtype == Token.NUMBER and expect_operand:
            program.append(("num", float(value)))
            expect_operand = False
        elif kind == Token.OPERAND and subtype == Token.RANGE and expect_operand:
            index = positions.get(value.strip().upper())
            if index is None:
                return None
            program.append(("ref", index))
            expect_operand = False
        elif kind == Token.OP_PRE and value in "+-" and expect_operand:
            operators.append("neg" if value == "-" else "pos")
        elif kind == Token.OP_POST and value == "%" and not expect_operand:
            _flush(program, operators, "%")
            program.append(("op", "%"))
        elif kind == Token.OP_IN and value in ("+", "-", "*", "/", "^") and not expect_operand:
            _flush(program, operators, value)
            operators.append(value)
            expect_operand = True
        elif kind == Token.PAREN and subtype == Token.OPEN and expect_operand:
            operators.append("(")
        elif kind == Token.PAREN and subtype == Token.CLOSE and not expect_operand:
            while operators and operators[-1] != "(":
                program.append(("op", operators.pop()))
            if not operators:
                return None
            operators.pop()
        else:
            return None
    if expect_operand:
        return None
    while operators:
        operator = operators.pop()
        if operator == "(":
            return None
        program.append(("op", operator))
    return lambda *values: _run(program, values)


def _flush(program: Program, operators: list[str], incoming: str) -> None:
    """Emit the stacked operators that bind at least as tightly as ``incoming`` (all binary
    operators are left-associative in Excel, ``^`` included)."""
    strength = _PRECEDENCE[incoming]
    while operators and operators[-1] != "(" and _PRECEDENCE[operators[-1]] >= strength:
        program.append(("op", operators.pop()))


def _run(program: Program, values: Sequence[Any]) -> float:
    """Run the program left to right, so the first error met is the one Excel reports."""
    stack: list[float] = []
    for kind, operand in program:
        if kind == "num":
            stack.append(operand)
        elif kind == "ref":
            stack.append(_number(values[operand]))
        elif operand == "neg":
            stack.append(-stack.pop())
        elif operand == "pos":
            pass
        elif operand == "%":
            stack.append(stack.pop() / 100)
        else:
            right, left = stack.pop(), stack.pop()
            stack.append(_binary(operand, left, right))
    result = stack.pop()
    if not math.isfinite(result):
        raise ExcelError("#NUM!")
    return result


def _binary(operator: str, left: float, right: float) -> float:
    if operator == "+":
        return left + right
    if operator == "-":
        return left - right
    if operator == "*":
        return left * right
    if operator == "/":
        if right == 0:
            raise ExcelError("#DIV/0!")
        return left / right
    # "^"
    if left == 0 and right == 0:
        raise ExcelError("#NUM!")
    if left == 0 and right < 0:
        raise ExcelError("#DIV/0!")
    if left < 0 and not float(right).is_integer():
        raise ExcelError("#NUM!")
    try:
        return math.pow(left, right)
    except (OverflowError, ValueError) as exc:
        raise ExcelError("#NUM!") from exc


def _number(value: Any) -> float:
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, int | float):
        return float(value)
    name = type(value).__name__
    if name == "XlError":
        raise ExcelError(str(value))
    if name == "Token" and str(value) == "empty":
        return 0.0
    if type(value).__module__ == "numpy" and getattr(value, "shape", None) == ():
        return _number(value.item())
    raise Fallback(name)


__all__ = ["ExcelError", "Fallback", "compile_arithmetic"]
