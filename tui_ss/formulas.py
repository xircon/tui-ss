#!/usr/bin/env python3
"""Formula parsing and evaluation."""

from __future__ import annotations

import ast
import math
import re
from collections.abc import Callable

from .model import Spreadsheet, parse_cell_reference, shift_cell_reference

CELL_RE = re.compile(r'(?<!["A-Za-z0-9_$])(\$?[A-Za-z]+\$?[0-9]+)\b(?!")')
RANGE_RE = re.compile(r"(\$?[A-Za-z]+\$?[0-9]+):(\$?[A-Za-z]+\$?[0-9]+)\b")
COMPARE_EQ_RE = re.compile(r"(?<![<>=!])=(?![=])")
FUNCTION_RE = re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)\s*(?=\()")


class FormulaError(Exception):
    """Raised when a formula cannot be evaluated."""


def shift_formula_references(raw: str, row_delta: int, col_delta: int) -> str:
    if not raw.startswith("="):
        return raw

    def replace_ref(match: re.Match[str]) -> str:
        return shift_cell_reference(match.group(1), row_delta, col_delta)

    return CELL_RE.sub(replace_ref, raw)


def coerce_number(value: object) -> float:
    if value in ("", None):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value))
    except ValueError as exc:
        raise FormulaError(f"not a number: {value}") from exc


class Evaluator:
    def __init__(self, sheet: Spreadsheet) -> None:
        self.sheet = sheet

    def display_value(self, row: int, col: int) -> str:
        raw = self.sheet.get_raw(row, col)
        if not raw:
            return ""
        if not raw.startswith("="):
            return raw
        value = self.evaluate_cell(row, col, set())
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)

    def evaluate_cell(self, row: int, col: int, seen: set[tuple[int, int]]) -> object:
        key = (row, col)
        if key in seen:
            raise FormulaError("circular reference")
        raw = self.sheet.get_raw(row, col)
        if not raw:
            return ""
        if not raw.startswith("="):
            return self._coerce_literal(raw)
        seen = set(seen)
        seen.add(key)
        expression = raw[1:].strip()
        return self.evaluate_expression(expression, seen)

    def evaluate_expression(self, expression: str, seen: set[tuple[int, int]]) -> object:
        expanded = RANGE_RE.sub(lambda match: f'RANGE("{match.group(1)}","{match.group(2)}")', expression)
        expanded = CELL_RE.sub(lambda match: f'CELL("{match.group(1)}")', expanded)
        expanded = FUNCTION_RE.sub(lambda match: match.group(1).upper(), expanded)
        expanded = expanded.replace("<>", "!=")
        expanded = COMPARE_EQ_RE.sub("==", expanded)
        try:
            node = ast.parse(expanded, mode="eval")
        except SyntaxError as exc:
            raise FormulaError("invalid expression") from exc
        return self._eval_node(node.body, seen)

    def _eval_node(self, node: ast.AST, seen: set[tuple[int, int]]) -> object:
        if isinstance(node, ast.Constant):
            return node.value
        if isinstance(node, ast.Compare):
            left = self._eval_node(node.left, seen)
            for op, comparator in zip(node.ops, node.comparators):
                right = self._eval_node(comparator, seen)
                if isinstance(op, ast.Eq):
                    passed = self._values_equal(left, right)
                elif isinstance(op, ast.NotEq):
                    passed = not self._values_equal(left, right)
                elif isinstance(op, ast.Lt):
                    passed = coerce_number(left) < coerce_number(right)
                elif isinstance(op, ast.LtE):
                    passed = coerce_number(left) <= coerce_number(right)
                elif isinstance(op, ast.Gt):
                    passed = coerce_number(left) > coerce_number(right)
                elif isinstance(op, ast.GtE):
                    passed = coerce_number(left) >= coerce_number(right)
                else:
                    raise FormulaError("unsupported comparison")
                if not passed:
                    return False
                left = right
            return True
        if isinstance(node, ast.BoolOp):
            values = [self._truthy(self._eval_node(value, seen)) for value in node.values]
            if isinstance(node.op, ast.And):
                return all(values)
            if isinstance(node.op, ast.Or):
                return any(values)
            raise FormulaError("unsupported logical operator")
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
            value = coerce_number(self._eval_node(node.operand, seen))
            return value if isinstance(node.op, ast.UAdd) else -value
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.Not):
            return not self._truthy(self._eval_node(node.operand, seen))
        if isinstance(node, ast.BinOp) and isinstance(node.op, (ast.Add, ast.Sub, ast.Mult, ast.Div, ast.Pow)):
            left = coerce_number(self._eval_node(node.left, seen))
            right = coerce_number(self._eval_node(node.right, seen))
            if isinstance(node.op, ast.Add):
                return left + right
            if isinstance(node.op, ast.Sub):
                return left - right
            if isinstance(node.op, ast.Mult):
                return left * right
            if right == 0:
                raise FormulaError("division by zero")
            if isinstance(node.op, ast.Div):
                return left / right
            return left**right
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
            name = node.func.id.upper()
            args = [self._eval_node(arg, seen) for arg in node.args]
            return self._call(name, args, seen)
        raise FormulaError("unsupported formula")

    def _call(self, name: str, args: list[object], seen: set[tuple[int, int]]) -> object:
        if name == "CELL":
            if len(args) != 1:
                raise FormulaError("CELL needs one argument")
            return self._lookup_cell(str(args[0]), seen)
        if name == "RANGE":
            if len(args) != 2:
                raise FormulaError("RANGE needs two arguments")
            return self._range_values(str(args[0]), str(args[1]), seen)
        if name == "IF":
            if len(args) != 3:
                raise FormulaError("IF needs three arguments")
            return args[1] if self._truthy(args[0]) else args[2]
        if name == "LOOKUP":
            return self._lookup(args)

        flattened = self._flatten(args)
        if name == "COUNT":
            return sum(1 for value in flattened if value not in ("", None))
        operations: dict[str, Callable[[list[float]], object]] = {
            "SUM": lambda values: sum(values),
            "AVG": lambda values: sum(values) / len(values) if values else 0.0,
            "AVERAGE": lambda values: sum(values) / len(values) if values else 0.0,
            "MIN": lambda values: min(values) if values else 0.0,
            "MAX": lambda values: max(values) if values else 0.0,
            "ROUND": lambda values: round(values[0], int(values[1])) if len(values) > 1 else round(values[0]),
            "ABS": lambda values: abs(values[0]),
            "INT": lambda values: math.trunc(values[0]),
            "COS": lambda values: math.cos(values[0]),
        }
        if name not in operations:
            raise FormulaError(f"unknown function: {name}")
        numeric = [coerce_number(value) for value in flattened]
        return operations[name](numeric)

    def _flatten(self, args: list[object]) -> list[object]:
        values: list[object] = []
        for item in args:
            if isinstance(item, list):
                values.extend(item)
            else:
                values.append(item)
        return values

    def _lookup_cell(self, ref: str, seen: set[tuple[int, int]]) -> object:
        row, col = parse_cell_reference(ref)
        return self.evaluate_cell(row, col, seen)

    def _lookup(self, args: list[object]) -> object:
        if len(args) not in {2, 3}:
            raise FormulaError("LOOKUP needs two or three arguments")
        needle = args[0]
        lookup_values = self._flatten([args[1]])
        result_values = lookup_values if len(args) == 2 else self._flatten([args[2]])
        if len(result_values) < len(lookup_values):
            raise FormulaError("LOOKUP result range is too small")
        for index, candidate in enumerate(lookup_values):
            if self._values_equal(candidate, needle):
                return result_values[index]
        raise FormulaError("lookup value not found")

    def _range_values(self, start_ref: str, end_ref: str, seen: set[tuple[int, int]]) -> list[object]:
        start_row, start_col = parse_cell_reference(start_ref)
        end_row, end_col = parse_cell_reference(end_ref)
        row_lo, row_hi = sorted((start_row, end_row))
        col_lo, col_hi = sorted((start_col, end_col))
        values: list[object] = []
        for row in range(row_lo, row_hi + 1):
            for col in range(col_lo, col_hi + 1):
                values.append(self.evaluate_cell(row, col, seen))
        return values

    def _coerce_literal(self, raw: str) -> object:
        try:
            if "." in raw:
                return float(raw)
            return int(raw)
        except ValueError:
            return raw

    def _truthy(self, value: object) -> bool:
        if isinstance(value, str):
            return value.strip().lower() not in {"", "0", "false", "no"}
        return bool(value)

    def _values_equal(self, left: object, right: object) -> bool:
        try:
            return coerce_number(left) == coerce_number(right)
        except FormulaError:
            return str(left) == str(right)
