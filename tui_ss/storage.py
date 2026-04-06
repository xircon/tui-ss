#!/usr/bin/env python3
"""Persistence helpers."""

from __future__ import annotations

import csv
import json
from pathlib import Path

from .model import Spreadsheet


def save_pdf_text(lines: list[str], path: Path, title: str = "tui-ss") -> None:
    path = path.expanduser()
    path.parent.mkdir(parents=True, exist_ok=True)
    page_width = 595
    page_height = 842
    left_margin = 40
    top_margin = 40
    line_height = 14
    font_size = 11
    lines_per_page = max(1, (page_height - (top_margin * 2)) // line_height)
    pages = [lines[index:index + lines_per_page] for index in range(0, max(1, len(lines)), lines_per_page)]
    if not pages:
        pages = [[""]]

    def pdf_escape(text: str) -> str:
        return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")

    def stream_for_page(page_lines: list[str]) -> bytes:
        commands = ["BT", f"/F1 {font_size} Tf", f"{left_margin} {page_height - top_margin} Td", f"({pdf_escape(title)}) Tj"]
        for line in page_lines:
            commands.append(f"0 -{line_height} Td")
            commands.append(f"({pdf_escape(line)}) Tj")
        commands.append("ET")
        return "\n".join(commands).encode("latin-1", errors="replace")

    objects: list[bytes] = []

    def add_object(data: bytes) -> int:
        objects.append(data)
        return len(objects)

    font_id = add_object(b"<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>")
    page_ids: list[int] = []
    content_ids: list[int] = []
    pages_root_id = 0

    for page_lines in pages:
        stream = stream_for_page(page_lines)
        content_id = add_object(b"<< /Length " + str(len(stream)).encode("ascii") + b" >>\nstream\n" + stream + b"\nendstream")
        content_ids.append(content_id)
        page_obj = (
            b"<< /Type /Page /Parent PAGES_REF 0 R /MediaBox [0 0 "
            + str(page_width).encode("ascii")
            + b" "
            + str(page_height).encode("ascii")
            + b"] /Contents "
            + str(content_id).encode("ascii")
            + b" 0 R /Resources << /Font << /F1 "
            + str(font_id).encode("ascii")
            + b" 0 R >> >> >>"
        )
        page_ids.append(add_object(page_obj))

    kids = b"[" + b" ".join(f"{page_id} 0 R".encode("ascii") for page_id in page_ids) + b"]"
    pages_root_id = add_object(b"<< /Type /Pages /Kids " + kids + b" /Count " + str(len(page_ids)).encode("ascii") + b" >>")
    catalog_id = add_object(b"<< /Type /Catalog /Pages " + str(pages_root_id).encode("ascii") + b" 0 R >>")

    objects = [obj.replace(b"PAGES_REF", str(pages_root_id).encode("ascii")) for obj in objects]

    pdf = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = [0]
    for index, obj in enumerate(objects, start=1):
        offsets.append(len(pdf))
        pdf.extend(f"{index} 0 obj\n".encode("ascii"))
        pdf.extend(obj)
        pdf.extend(b"\nendobj\n")
    xref_offset = len(pdf)
    pdf.extend(f"xref\n0 {len(objects) + 1}\n".encode("ascii"))
    pdf.extend(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        pdf.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
    pdf.extend(
        b"trailer\n<< /Size "
        + str(len(objects) + 1).encode("ascii")
        + b" /Root "
        + str(catalog_id).encode("ascii")
        + b" 0 R >>\nstartxref\n"
        + str(xref_offset).encode("ascii")
        + b"\n%%EOF\n"
    )
    path.write_bytes(bytes(pdf))


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
