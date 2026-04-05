#!/usr/bin/env python3
"""Persistence helpers."""

from __future__ import annotations

import csv
import json
from pathlib import Path

from .model import Spreadsheet


def save_sheet(sheet: Spreadsheet, path: Path) -> None:
    path = path.expanduser()
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.suffix.lower() == ".csv":
        _save_csv(sheet, path)
        return
    if path.suffix.lower() == ".tsv":
        _save_delimited(sheet, path, "\t")
        return
    path.write_text(json.dumps(sheet.to_dict(), indent=2), encoding="utf-8")


def load_sheet(path: Path) -> Spreadsheet:
    path = path.expanduser()
    if path.suffix.lower() == ".csv":
        return _load_csv(path)
    if path.suffix.lower() == ".tsv":
        return _load_delimited(path, "\t")
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError("sheet file must contain an object")
    return Spreadsheet.from_dict(payload)


def _save_csv(sheet: Spreadsheet, path: Path) -> None:
    _save_delimited(sheet, path, ",")


def _save_delimited(sheet: Spreadsheet, path: Path, delimiter: str) -> None:
    rows: dict[int, dict[int, str]] = {}
    max_row = 0
    max_col = 0
    for row, col, raw in sheet.iter_cells():
        rows.setdefault(row, {})[col] = raw
        max_row = max(max_row, row)
        max_col = max(max_col, col)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle, delimiter=delimiter)
        for row in range(max_row + 1):
            writer.writerow([rows.get(row, {}).get(col, "") for col in range(max_col + 1)])


def _load_csv(path: Path) -> Spreadsheet:
    return _load_delimited(path, ",")


def _load_delimited(path: Path, delimiter: str) -> Spreadsheet:
    sheet = Spreadsheet()
    with path.open("r", encoding="utf-8", newline="") as handle:
        reader = csv.reader(handle, delimiter=delimiter)
        for row_index, row in enumerate(reader):
            for col_index, value in enumerate(row):
                if value:
                    sheet.set_raw(row_index, col_index, value)
    return sheet
