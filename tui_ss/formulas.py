#!/usr/bin/env python3
"""Formula parsing and evaluation."""

from __future__ import annotations

import ast
import math
import re
from collections.abc import Callable
from datetime import date, datetime

from .model import Spreadsheet, column_label, parse_cell_reference, parse_reference_parts, shift_cell_reference

CELL_RE = re.compile(r'(?<!["A-Za-z0-9_$])(\$?[A-Za-z]+\$?[0-9]+)\b(?!")')
RANGE_RE = re.compile(r"(\$?[A-Za-z]+\$?[0-9]+):(\$?[A-Za-z]+\$?[0-9]+)\b")
COMPARE_EQ_RE = re.compile(r"(?<![<>=!])=(?![=])")
FUNCTION_RE = re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)\s*(?=\()")


class FormulaError(Exception):
    """Raised when a formula cannot be evaluated."""


DATE_STYLE_ALIASES = {
    "date": "date:ansi",
    "date:ansi": "date:ansi",
    "date:us": "date:us",
    "date:european": "date:european",
    "date:uk": "date:uk",
}
TIME_STYLE_ALIASES = {
    "time": "time:24h",
    "time:24h": "time:24h",
    "time:24h-seconds": "time:24h-seconds",
    "time:12h": "time:12h",
    "time:12h-seconds": "time:12h-seconds",
}


def is_formula_text(raw: str) -> bool:
    return bool(raw) and raw.startswith("=")


def unescape_literal_text(raw: str) -> str:
    if raw.startswith("'"):
        return raw[1:]
    return raw


def shift_formula_references(raw: str, row_delta: int, col_delta: int) -> str:
    if not is_formula_text(raw):
        return raw

    def replace_ref(match: re.Match[str]) -> str:
        return shift_cell_reference(match.group(1), row_delta, col_delta)

    return CELL_RE.sub(replace_ref, raw)


def shift_formula_references_for_structure(
    raw: str,
    *,
    row_index: int | None = None,
    row_delta: int = 0,
    col_index: int | None = None,
    col_delta: int = 0,
) -> str:
    if not is_formula_text(raw):
        return raw

    def shift_position(position: int, index: int | None, delta: int) -> int:
        if index is None or delta == 0:
            return position
        if delta > 0:
            return position + delta if position >= index else position
        deleted_count = -delta
        if position >= index + deleted_count:
            return max(0, position + delta)
        if position >= index:
            return max(0, index)
        return position

    def replace_ref(match: re.Match[str]) -> str:
        ref = match.group(1)
        row, col, row_absolute, col_absolute = parse_reference_parts(ref)
        new_row = shift_position(row, row_index, row_delta)
        new_col = shift_position(col, col_index, col_delta)
        col_prefix = "$" if col_absolute else ""
        row_prefix = "$" if row_absolute else ""
        return f"{col_prefix}{column_label(new_col)}{row_prefix}{new_row + 1}"

    return CELL_RE.sub(replace_ref, raw)


def coerce_number(value: object) -> float:
    if value in ("", None):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value))
    except ValueError as exc:
        parsed_date = parse_date_text(str(value))
        if parsed_date is not None:
            return float(parsed_date.toordinal())
        parsed_time = parse_time_text(str(value))
        if parsed_time is not None:
            return float(parsed_time) / 86400.0
        raise FormulaError(f"not a number: {value}") from exc


def parse_date_text(text: str, style: str = "") -> date | None:
    cleaned = text.strip()
    if not cleaned:
        return None
    if " " in cleaned:
        cleaned = cleaned.split(" ", 1)[0]
    normalized_style = DATE_STYLE_ALIASES.get(style.lower(), style.lower())

    def parse_ymd(value: str) -> date | None:
        parts = value.split("-")
        if len(parts) != 3:
            return None
        try:
            year, month, day = (int(part) for part in parts)
            return date(year, month, day)
        except ValueError:
            return None

    def parse_mdy(value: str) -> date | None:
        parts = value.split("/")
        if len(parts) != 3:
            return None
        try:
            month, day, year = (int(part) for part in parts)
            return date(year, month, day)
        except ValueError:
            return None

    def parse_dmy(value: str) -> date | None:
        parts = value.split("/")
        if len(parts) != 3:
            return None
        try:
            day, month, year = (int(part) for part in parts)
            return date(year, month, day)
        except ValueError:
            return None

    def parse_dmy_short(value: str) -> date | None:
        parts = value.split("/")
        if len(parts) != 3:
            return None
        try:
            day, month, year = (int(part) for part in parts)
            if year < 100:
                year += 2000 if year <= 68 else 1900
            return date(year, month, day)
        except ValueError:
            return None

    parsers: list[Callable[[str], date | None]]
    if normalized_style == "date:us":
        parsers = [parse_mdy, parse_ymd]
    elif normalized_style == "date:uk":
        parsers = [parse_dmy_short, parse_dmy, parse_ymd]
    elif normalized_style == "date:european":
        parsers = [parse_dmy, parse_ymd]
    else:
        parsers = [parse_ymd, parse_mdy, parse_dmy]
    for parser in parsers:
        parsed = parser(cleaned)
        if parsed is not None:
            return parsed
    return None


def parse_time_text(text: str) -> int | None:
    cleaned = text.strip().lower()
    if not cleaned:
        return None
    if " " in cleaned:
        parts = cleaned.split()
        cleaned = parts[-1] if ":" in parts[-1] else cleaned
    match = re.match(r"^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)?$", cleaned)
    if not match:
        return None
    hour = int(match.group(1))
    minute = int(match.group(2))
    second = int(match.group(3) or 0)
    meridiem = match.group(4)
    if meridiem:
        if hour == 12:
            hour = 0
        if meridiem == "pm":
            hour += 12
    if hour < 0 or hour > 23 or minute > 59 or second > 59:
        return None
    return hour * 3600 + minute * 60 + second


def normalize_date_text(text: str, style: str) -> str:
    parsed = parse_date_text(text, style)
    if parsed is None:
        raise ValueError(f"invalid date: {text}")
    return parsed.strftime("%Y-%m-%d")


def format_date_text(text: str, style: str) -> str:
    parsed = parse_date_text(text, "date:ansi")
    if parsed is None:
        return text
    normalized_style = DATE_STYLE_ALIASES.get(style.lower(), style.lower())
    if normalized_style == "date:us":
        return parsed.strftime("%m/%d/%Y")
    if normalized_style == "date:uk":
        return parsed.strftime("%d/%m/%y")
    if normalized_style == "date:european":
        return parsed.strftime("%d/%m/%Y")
    return parsed.strftime("%Y-%m-%d")


def normalize_time_text(text: str, style: str) -> str:
    parsed = parse_time_text(text)
    if parsed is None:
        raise ValueError(f"invalid time: {text}")
    return format_time_text(str(parsed / 86400.0), style or "time:24h-seconds")


def format_time_text(text: str, style: str) -> str:
    normalized_style = TIME_STYLE_ALIASES.get(style.lower(), style.lower())
    parsed = parse_time_text(text)
    if parsed is None:
        try:
            number = float(text)
            parsed = int(round((number % 1) * 86400))
        except ValueError:
            return text
    hours = (parsed // 3600) % 24
    minutes = (parsed % 3600) // 60
    seconds = parsed % 60
    if normalized_style in {"time:12h", "time:12h-seconds"}:
        suffix = "AM" if hours < 12 else "PM"
        display_hour = hours % 12
        if display_hour == 0:
            display_hour = 12
        if normalized_style.endswith("seconds"):
            return f"{display_hour:02d}:{minutes:02d}:{seconds:02d} {suffix}"
        return f"{display_hour:02d}:{minutes:02d} {suffix}"
    if normalized_style.endswith("seconds"):
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    return f"{hours:02d}:{minutes:02d}"


class Evaluator:
    def __init__(self, sheet: Spreadsheet) -> None:
        self.sheet = sheet

    def display_value(self, row: int, col: int) -> str:
        raw = self.sheet.get_raw(row, col)
        if not raw:
            return ""
        if not is_formula_text(raw):
            return unescape_literal_text(raw)
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
        if not is_formula_text(raw):
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
        if isinstance(node, ast.Name):
            return self._lookup_name(node.id, seen)
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
            if name == "IF":
                if len(node.args) != 3:
                    raise FormulaError("IF needs three arguments")
                condition = self._eval_node(node.args[0], seen)
                branch = node.args[1] if self._truthy(condition) else node.args[2]
                return self._eval_node(branch, seen)
            if name == "IFERROR":
                if len(node.args) != 2:
                    raise FormulaError("IFERROR needs two arguments")
                try:
                    return self._eval_node(node.args[0], seen)
                except FormulaError:
                    return self._eval_node(node.args[1], seen)
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
        if name == "LOOKUP":
            return self._lookup(args)
        if name == "VLOOKUP":
            return self._vlookup(args)
        if name == "HLOOKUP":
            return self._hlookup(args)
        if name == "MATCH":
            return self._match(args)
        if name == "COUNTIF":
            return self._countif(args)
        if name == "SUMIF":
            return self._sumif(args)
        if name == "INDEX":
            return self._index(args)
        if name == "DATE":
            if len(args) != 3:
                raise FormulaError("DATE needs three arguments")
            return date(int(coerce_number(args[0])), int(coerce_number(args[1])), int(coerce_number(args[2]))).strftime("%Y-%m-%d")
        if name == "TIME":
            if len(args) != 3:
                raise FormulaError("TIME needs three arguments")
            hour = int(coerce_number(args[0]))
            minute = int(coerce_number(args[1]))
            second = int(coerce_number(args[2]))
            total_seconds = (hour % 24) * 3600 + (minute % 60) * 60 + (second % 60)
            return format_time_text(str(total_seconds / 86400.0), "time:24h-seconds")
        if name == "NOW":
            if args:
                raise FormulaError("NOW needs no arguments")
            return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if name == "TODAY":
            if args:
                raise FormulaError("TODAY needs no arguments")
            return date.today().strftime("%Y-%m-%d")
        if name == "HOUR":
            if len(args) != 1:
                raise FormulaError("HOUR needs one argument")
            return self._coerce_time(args[0]) // 3600
        if name == "MINUTE":
            if len(args) != 1:
                raise FormulaError("MINUTE needs one argument")
            return (self._coerce_time(args[0]) % 3600) // 60
        if name == "SECOND":
            if len(args) != 1:
                raise FormulaError("SECOND needs one argument")
            return self._coerce_time(args[0]) % 60
        if name == "TIMEVALUE":
            if len(args) != 1:
                raise FormulaError("TIMEVALUE needs one argument")
            return float(self._coerce_time(args[0])) / 86400.0
        if name == "YEAR":
            return self._date_part(args, "year")
        if name == "MONTH":
            return self._date_part(args, "month")
        if name == "DAY":
            return self._date_part(args, "day")
        if name == "DATEDIFF":
            if len(args) != 2:
                raise FormulaError("DATEDIFF needs two arguments")
            start_date = self._coerce_date(args[0])
            end_date = self._coerce_date(args[1])
            return (end_date - start_date).days
        if name == "WEEKDAY":
            if len(args) != 1:
                raise FormulaError("WEEKDAY needs one argument")
            return self._coerce_date(args[0]).isoweekday()
        if name == "AND":
            return all(self._truthy(value) for value in self._flatten(args))
        if name == "OR":
            return any(self._truthy(value) for value in self._flatten(args))
        if name == "NOT":
            if len(args) != 1:
                raise FormulaError("NOT needs one argument")
            return not self._truthy(args[0])
        if name == "LEN":
            if len(args) != 1:
                raise FormulaError("LEN needs one argument")
            return len(self._stringify(args[0]))
        if name == "LEFT":
            if not args:
                raise FormulaError("LEFT needs at least one argument")
            text = self._stringify(args[0])
            count = int(coerce_number(args[1])) if len(args) > 1 else 1
            return text[: max(0, count)]
        if name == "RIGHT":
            if not args:
                raise FormulaError("RIGHT needs at least one argument")
            text = self._stringify(args[0])
            count = int(coerce_number(args[1])) if len(args) > 1 else 1
            return text[-max(0, count) :] if count else ""
        if name == "MID":
            if len(args) != 3:
                raise FormulaError("MID needs three arguments")
            text = self._stringify(args[0])
            start = max(1, int(coerce_number(args[1])))
            length = max(0, int(coerce_number(args[2])))
            return text[start - 1 : start - 1 + length]
        if name == "CONCAT":
            return "".join(self._stringify(value) for value in self._flatten(args))
        if name == "VALUE":
            if len(args) != 1:
                raise FormulaError("VALUE needs one argument")
            numeric = coerce_number(args[0])
            return int(numeric) if float(numeric).is_integer() else numeric
        if name == "TEXT":
            if len(args) != 2:
                raise FormulaError("TEXT needs two arguments")
            return self._format_text(args[0], self._stringify(args[1]))

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
            "SIN": lambda values: math.sin(values[0]),
            "TAN": lambda values: math.tan(values[0]),
            "MOD": lambda values: values[0] % values[1],
            "SQRT": lambda values: math.sqrt(values[0]),
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

    def _lookup_name(self, name: str, seen: set[tuple[int, int]]) -> object:
        spec = self.sheet.get_named_range(name)
        if not spec:
            raise FormulaError(f"unknown name: {name}")
        if ":" in spec:
            start_ref, end_ref = spec.split(":", 1)
            return self._range_values(start_ref, end_ref, seen)
        return self._lookup_cell(spec, seen)

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

    def _vlookup(self, args: list[object]) -> object:
        if len(args) != 3:
            raise FormulaError("VLOOKUP needs three arguments")
        needle = args[0]
        table = self._table_values(args[1])
        result_col = int(coerce_number(args[2])) - 1
        if not table or result_col < 0:
            raise FormulaError("VLOOKUP column is invalid")
        width = len(table[0])
        if result_col >= width:
            raise FormulaError("VLOOKUP column is out of range")
        for row in table:
            if row and self._values_equal(row[0], needle):
                return row[result_col]
        raise FormulaError("lookup value not found")

    def _hlookup(self, args: list[object]) -> object:
        if len(args) != 3:
            raise FormulaError("HLOOKUP needs three arguments")
        needle = args[0]
        table = self._table_values(args[1])
        result_row = int(coerce_number(args[2])) - 1
        if not table or result_row < 0:
            raise FormulaError("HLOOKUP row is invalid")
        height = len(table)
        if result_row >= height:
            raise FormulaError("HLOOKUP row is out of range")
        for col_index, candidate in enumerate(table[0]):
            if self._values_equal(candidate, needle):
                return table[result_row][col_index]
        raise FormulaError("lookup value not found")

    def _match(self, args: list[object]) -> int:
        if len(args) != 2:
            raise FormulaError("MATCH needs two arguments")
        needle = args[0]
        values = self._flatten([args[1]])
        for index, candidate in enumerate(values, start=1):
            if self._values_equal(candidate, needle):
                return index
        raise FormulaError("lookup value not found")

    def _index(self, args: list[object]) -> object:
        if len(args) not in {2, 3}:
            raise FormulaError("INDEX needs two or three arguments")
        target = args[0]
        row_index = max(1, int(coerce_number(args[1]))) - 1
        if isinstance(target, list) and target and isinstance(target[0], list):
            col_index = max(1, int(coerce_number(args[2]))) - 1 if len(args) > 2 else 0
            table = self._table_values(target)
            if row_index >= len(table) or col_index >= len(table[0]):
                raise FormulaError("INDEX is out of range")
            return table[row_index][col_index]
        values = self._flatten([target])
        if row_index >= len(values):
            raise FormulaError("INDEX is out of range")
        return values[row_index]

    def _sumif(self, args: list[object]) -> object:
        if len(args) != 3:
            raise FormulaError("SUMIF needs three arguments")
        criteria_values = self._flatten([args[0]])
        sum_values = self._flatten([args[2]])
        if len(criteria_values) != len(sum_values):
            raise FormulaError("SUMIF ranges must be the same size")
        total = 0.0
        for candidate, sum_value in zip(criteria_values, sum_values):
            if self._criteria_matches(candidate, args[1]):
                total += coerce_number(sum_value)
        return total

    def _countif(self, args: list[object]) -> int:
        if len(args) != 2:
            raise FormulaError("COUNTIF needs two arguments")
        values = self._flatten([args[0]])
        return sum(1 for candidate in values if self._criteria_matches(candidate, args[1]))

    def _table_values(self, value: object) -> list[list[object]]:
        if not isinstance(value, list):
            raise FormulaError("lookup table must be a range")
        if not value:
            return []
        if value and not isinstance(value[0], list):
            raise FormulaError("lookup table must be rectangular")
        width = len(value[0])
        for row in value:
            if not isinstance(row, list) or len(row) != width:
                raise FormulaError("lookup table must be rectangular")
        return value

    def _coerce_date(self, value: object) -> date:
        if isinstance(value, date) and not isinstance(value, datetime):
            return value
        parsed = parse_date_text(str(value))
        if parsed is None:
            raise FormulaError(f"not a date: {value}")
        return parsed

    def _coerce_time(self, value: object) -> int:
        if isinstance(value, datetime):
            return value.hour * 3600 + value.minute * 60 + value.second
        parsed = parse_time_text(str(value))
        if parsed is None:
            raise FormulaError(f"not a time: {value}")
        return parsed

    def _date_part(self, args: list[object], part: str) -> int:
        if len(args) != 1:
            raise FormulaError(f"{part.upper()} needs one argument")
        parsed = self._coerce_date(args[0])
        if part == "year":
            return parsed.year
        if part == "month":
            return parsed.month
        return parsed.day

    def _range_values(self, start_ref: str, end_ref: str, seen: set[tuple[int, int]]) -> list[object]:
        start_row, start_col = parse_cell_reference(start_ref)
        end_row, end_col = parse_cell_reference(end_ref)
        row_lo, row_hi = sorted((start_row, end_row))
        col_lo, col_hi = sorted((start_col, end_col))
        values: list[object] | list[list[object]]
        if row_lo != row_hi and col_lo != col_hi:
            values = []
            for row in range(row_lo, row_hi + 1):
                row_values: list[object] = []
                for col in range(col_lo, col_hi + 1):
                    row_values.append(self.evaluate_cell(row, col, seen))
                values.append(row_values)
            return values
        values = []
        for row in range(row_lo, row_hi + 1):
            for col in range(col_lo, col_hi + 1):
                values.append(self.evaluate_cell(row, col, seen))
        return values

    def _coerce_literal(self, raw: str) -> object:
        raw = unescape_literal_text(raw)
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

    def _criteria_matches(self, candidate: object, criterion: object) -> bool:
        if not isinstance(criterion, str):
            return self._values_equal(candidate, criterion)
        text = criterion.strip()
        operator = "="
        operand = text
        for token in ("<=", ">=", "<>", "!=", "<", ">", "="):
            if text.startswith(token):
                operator = token
                operand = text[len(token) :].strip()
                break
        if operator == "=" and operand == text:
            return self._values_equal(candidate, criterion)
        return self._compare_values(candidate, operand, operator)

    def _compare_values(self, left: object, right: object, operator: str) -> bool:
        try:
            left_value = coerce_number(left)
            right_value = coerce_number(right)
        except FormulaError:
            left_value = str(left)
            right_value = str(right)
        if operator == "=":
            return left_value == right_value
        if operator in {"<>", "!="}:
            return left_value != right_value
        if operator == "<":
            return left_value < right_value
        if operator == "<=":
            return left_value <= right_value
        if operator == ">":
            return left_value > right_value
        if operator == ">=":
            return left_value >= right_value
        raise FormulaError(f"unsupported criteria operator: {operator}")

    def _stringify(self, value: object) -> str:
        if value in (None, ""):
            return ""
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)

    def _format_text(self, value: object, pattern: str) -> str:
        normalized = pattern.strip().lower()
        if normalized in {"yyyy-mm-dd", "ansi"}:
            parsed = self._coerce_date(value)
            return parsed.strftime("%Y-%m-%d")
        if normalized in {"dd/mm/yyyy", "european", "eu"}:
            parsed = self._coerce_date(value)
            return parsed.strftime("%d/%m/%Y")
        if normalized in {"mm/dd/yyyy", "us"}:
            parsed = self._coerce_date(value)
            return parsed.strftime("%m/%d/%Y")
        if normalized in {"hh:mm", "h:mm"}:
            return format_time_text(str(coerce_number(value)), "time:24h")
        if normalized in {"hh:mm:ss", "h:mm:ss"}:
            return format_time_text(str(coerce_number(value)), "time:24h-seconds")
        if normalized in {"hh:mm am/pm", "h:mm am/pm"}:
            return format_time_text(str(coerce_number(value)), "time:12h")
        if normalized in {"hh:mm:ss am/pm", "h:mm:ss am/pm"}:
            return format_time_text(str(coerce_number(value)), "time:12h-seconds")
        numeric = coerce_number(value)
        if normalized.endswith("%"):
            decimals = 0
            if "." in normalized:
                decimals = len(normalized.split(".", 1)[1].rstrip("%"))
            return f"{numeric * 100:.{decimals}f}%"
        if "." in normalized and all(char in "#0.," for char in normalized):
            decimals = len(normalized.split(".", 1)[1])
            use_grouping = "," in normalized.split(".", 1)[0]
            return f"{numeric:,.{decimals}f}" if use_grouping else f"{numeric:.{decimals}f}"
        if normalized in {"0", "#"}:
            return str(int(round(numeric)))
        return self._stringify(value)
