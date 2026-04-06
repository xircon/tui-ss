#!/usr/bin/env python3
"""Spreadsheet data model."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterable


def column_label(index: int) -> str:
    if index < 0:
        raise ValueError("column index must be non-negative")
    label = ""
    current = index
    while True:
        current, remainder = divmod(current, 26)
        label = chr(ord("A") + remainder) + label
        if current == 0:
            return label
        current -= 1


def parse_reference_parts(ref: str) -> tuple[int, int, bool, bool]:
    token = ref.strip().upper()
    letters = []
    digits = []
    col_absolute = False
    row_absolute = False
    index = 0
    if index < len(token) and token[index] == "$":
        col_absolute = True
        index += 1
    while index < len(token) and token[index].isalpha():
        letters.append(token[index])
        index += 1
    if index < len(token) and token[index] == "$":
        row_absolute = True
        index += 1
    while index < len(token) and token[index].isdigit():
        digits.append(token[index])
        index += 1
    if index != len(token) or not letters or not digits:
        raise ValueError(f"invalid cell reference: {ref}")
    column = 0
    for char in letters:
        column = (column * 26) + (ord(char) - ord("A") + 1)
    return int("".join(digits)) - 1, column - 1, row_absolute, col_absolute


def parse_cell_reference(ref: str) -> tuple[int, int]:
    row, col, _row_absolute, _col_absolute = parse_reference_parts(ref)
    return row, col


def shift_cell_reference(ref: str, row_delta: int, col_delta: int) -> str:
    row, col, row_absolute, col_absolute = parse_reference_parts(ref)
    new_row = row if row_absolute else max(0, row + row_delta)
    new_col = col if col_absolute else max(0, col + col_delta)
    col_prefix = "$" if col_absolute else ""
    row_prefix = "$" if row_absolute else ""
    return f"{col_prefix}{column_label(new_col)}{row_prefix}{new_row + 1}"


@dataclass(slots=True)
class Cell:
    raw: str = ""


@dataclass(slots=True)
class Spreadsheet:
    rows: int = 50
    cols: int = 26
    cells: dict[str, Cell] = field(default_factory=dict)
    formats: dict[str, str] = field(default_factory=dict)
    backgrounds: dict[str, str] = field(default_factory=dict)
    alignments: dict[str, str] = field(default_factory=dict)
    manual_alignments: set[str] = field(default_factory=set)
    protected: set[str] = field(default_factory=set)
    column_width: int = 12
    column_widths: dict[str, int] = field(default_factory=dict)
    title_rows: int = 0
    title_cols: int = 0
    theme_name: str = "white"
    date_format: str = "date:european"
    active_cell_color: str = "orange"
    language: str = "en"

    def key(self, row: int, col: int) -> str:
        return f"{row}:{col}"

    def in_bounds(self, row: int, col: int) -> bool:
        return 0 <= row < self.rows and 0 <= col < self.cols

    def ensure_size(self, row: int, col: int) -> None:
        if row >= self.rows:
            self.rows = row + 1
        if col >= self.cols:
            self.cols = col + 1

    def set_raw(self, row: int, col: int, raw: str) -> None:
        self.ensure_size(row, col)
        key = self.key(row, col)
        if raw == "":
            self.cells.pop(key, None)
            return
        self.cells[key] = Cell(raw=raw)

    def get_raw(self, row: int, col: int) -> str:
        cell = self.cells.get(self.key(row, col))
        return "" if cell is None else cell.raw

    def clear(self, row: int, col: int) -> None:
        self.cells.pop(self.key(row, col), None)

    def is_protected(self, row: int, col: int) -> bool:
        return self.key(row, col) in self.protected

    def set_format(self, row: int, col: int, style: str) -> None:
        key = self.key(row, col)
        if style:
            self.formats[key] = style
        else:
            self.formats.pop(key, None)

    def get_format(self, row: int, col: int) -> str:
        return self.formats.get(self.key(row, col), "")

    def set_background(self, row: int, col: int, color: str) -> None:
        key = self.key(row, col)
        if color:
            self.backgrounds[key] = color
        else:
            self.backgrounds.pop(key, None)

    def get_background(self, row: int, col: int) -> str:
        return self.backgrounds.get(self.key(row, col), "")

    def set_alignment(self, row: int, col: int, align: str, manual: bool = True) -> None:
        key = self.key(row, col)
        if align:
            self.alignments[key] = align
            if manual:
                self.manual_alignments.add(key)
            else:
                self.manual_alignments.discard(key)
        else:
            self.alignments.pop(key, None)
            self.manual_alignments.discard(key)

    def get_alignment(self, row: int, col: int) -> str:
        return self.alignments.get(self.key(row, col), "")

    def is_alignment_manual(self, row: int, col: int) -> bool:
        return self.key(row, col) in self.manual_alignments

    def protect(self, row: int, col: int) -> None:
        self.protected.add(self.key(row, col))

    def unprotect(self, row: int, col: int) -> None:
        self.protected.discard(self.key(row, col))

    def get_column_width(self, col: int) -> int:
        return max(8, int(self.column_widths.get(str(col), self.column_width)))

    def set_column_width(self, col: int, width: int) -> None:
        self.column_widths[str(col)] = max(8, int(width))

    def iter_cells(self) -> Iterable[tuple[int, int, str]]:
        for key, cell in self.cells.items():
            row_text, col_text = key.split(":", 1)
            yield int(row_text), int(col_text), cell.raw

    def to_dict(self) -> dict[str, object]:
        return {
            "rows": self.rows,
            "cols": self.cols,
            "column_width": self.column_width,
            "column_widths": self.column_widths,
            "title_rows": self.title_rows,
            "title_cols": self.title_cols,
            "theme_name": self.theme_name,
            "date_format": self.date_format,
            "active_cell_color": self.active_cell_color,
            "language": self.language,
            "manual_alignments": [
                {"row": row, "col": col}
                for row, col in sorted(self.iter_manual_alignments())
            ],
            "alignments": [
                {"row": row, "col": col, "align": align}
                for row, col, align in sorted(self.iter_alignments())
            ],
            "formats": [
                {"row": row, "col": col, "style": style}
                for row, col, style in sorted(self.iter_formats())
            ],
            "backgrounds": [
                {"row": row, "col": col, "color": color}
                for row, col, color in sorted(self.iter_backgrounds())
            ],
            "protected": [
                {"row": row, "col": col}
                for row, col in sorted(self.iter_protected())
            ],
            "cells": [
                {"row": row, "col": col, "raw": raw}
                for row, col, raw in sorted(self.iter_cells())
            ],
        }

    def iter_formats(self) -> Iterable[tuple[int, int, str]]:
        for key, style in self.formats.items():
            row_text, col_text = key.split(":", 1)
            yield int(row_text), int(col_text), style

    def iter_backgrounds(self) -> Iterable[tuple[int, int, str]]:
        for key, color in self.backgrounds.items():
            row_text, col_text = key.split(":", 1)
            yield int(row_text), int(col_text), color

    def iter_alignments(self) -> Iterable[tuple[int, int, str]]:
        for key, align in self.alignments.items():
            row_text, col_text = key.split(":", 1)
            yield int(row_text), int(col_text), align

    def iter_manual_alignments(self) -> Iterable[tuple[int, int]]:
        for key in self.manual_alignments:
            row_text, col_text = key.split(":", 1)
            yield int(row_text), int(col_text)

    def iter_protected(self) -> Iterable[tuple[int, int]]:
        for key in self.protected:
            row_text, col_text = key.split(":", 1)
            yield int(row_text), int(col_text)

    @classmethod
    def from_dict(cls, payload: dict[str, object]) -> "Spreadsheet":
        sheet = cls(
            rows=max(1, int(payload.get("rows", 50))),
            cols=max(1, int(payload.get("cols", 26))),
        )
        sheet.column_width = max(8, int(payload.get("column_width", 12)))
        raw_widths = payload.get("column_widths", {})
        if isinstance(raw_widths, dict):
            sheet.column_widths = {str(key): max(8, int(value)) for key, value in raw_widths.items()}
        sheet.title_rows = max(0, int(payload.get("title_rows", 0)))
        sheet.title_cols = max(0, int(payload.get("title_cols", 0)))
        sheet.theme_name = str(payload.get("theme_name", "white")) or "white"
        raw_date_format = str(payload.get("date_format", "date:european")) or "date:european"
        sheet.date_format = raw_date_format if raw_date_format.startswith("date:") else "date:european"
        sheet.active_cell_color = str(payload.get("active_cell_color", "orange")) or "orange"
        sheet.language = str(payload.get("language", "en")) or "en"
        for item in payload.get("cells", []):
            if not isinstance(item, dict):
                continue
            row = int(item.get("row", 0))
            col = int(item.get("col", 0))
            raw = str(item.get("raw", ""))
            if raw:
                sheet.set_raw(row, col, raw)
        for item in payload.get("formats", []):
            if not isinstance(item, dict):
                continue
            sheet.set_format(int(item.get("row", 0)), int(item.get("col", 0)), str(item.get("style", "")))
        for item in payload.get("backgrounds", []):
            if not isinstance(item, dict):
                continue
            sheet.set_background(int(item.get("row", 0)), int(item.get("col", 0)), str(item.get("color", "")))
        for item in payload.get("alignments", []):
            if not isinstance(item, dict):
                continue
            sheet.set_alignment(int(item.get("row", 0)), int(item.get("col", 0)), str(item.get("align", "")), manual=False)
        for item in payload.get("manual_alignments", []):
            if not isinstance(item, dict):
                continue
            row = int(item.get("row", 0))
            col = int(item.get("col", 0))
            if sheet.get_alignment(row, col):
                sheet.manual_alignments.add(sheet.key(row, col))
        for item in payload.get("protected", []):
            if not isinstance(item, dict):
                continue
            sheet.protect(int(item.get("row", 0)), int(item.get("col", 0)))
        return sheet
