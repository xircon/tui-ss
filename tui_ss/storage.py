#!/usr/bin/env python3
"""Persistence helpers."""

from __future__ import annotations

import csv
import json
import tomllib
import zipfile
from datetime import UTC, datetime
from pathlib import Path
from xml.sax.saxutils import escape

from .formulas import is_formula_text, unescape_literal_text
from .model import Spreadsheet, column_label


def load_app_settings(path: Path) -> dict[str, str]:
    path = path.expanduser()
    if not path.exists():
        return {}
    payload = tomllib.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        return {}
    settings = payload.get("settings", payload)
    if not isinstance(settings, dict):
        return {}
    return {str(key): str(value) for key, value in settings.items()}


def save_app_settings(path: Path, settings: dict[str, str]) -> None:
    path = path.expanduser()
    path.parent.mkdir(parents=True, exist_ok=True)
    lines = ["[settings]"]
    for key in sorted(settings):
        value = settings[key].replace("\\", "\\\\").replace('"', '\\"')
        lines.append(f'{key} = "{value}"')
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


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
    if path.suffix.lower() == ".xlsx":
        save_xlsx(sheet, path)
        return
    if path.suffix.lower() == ".ods":
        save_ods(sheet, path)
        return
    path.write_text(json.dumps(sheet.to_dict(), indent=2), encoding="utf-8")


def load_sheet(path: Path, defaults: dict[str, str] | None = None) -> Spreadsheet:
    path = path.expanduser()
    if path.suffix.lower() == ".csv":
        sheet = _load_csv(path)
        _apply_sheet_defaults(sheet, defaults)
        return sheet
    if path.suffix.lower() == ".tsv":
        sheet = _load_delimited(path, "\t")
        _apply_sheet_defaults(sheet, defaults)
        return sheet
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError("sheet file must contain an object")
    sheet = Spreadsheet.from_dict(payload)
    _apply_sheet_defaults(sheet, defaults, payload)
    return sheet


def _apply_sheet_defaults(
    sheet: Spreadsheet,
    defaults: dict[str, str] | None,
    payload: dict[str, object] | None = None,
) -> None:
    if not defaults:
        return
    if defaults.get("theme_name"):
        sheet.theme_name = defaults["theme_name"]
    if defaults.get("date_format"):
        raw_date = defaults["date_format"]
        sheet.date_format = raw_date if raw_date.startswith("date:") else f"date:{raw_date}"
    if defaults.get("time_format"):
        raw_time = defaults["time_format"]
        sheet.time_format = raw_time if raw_time.startswith("time:") else f"time:{raw_time}"
    if defaults.get("active_cell_color"):
        sheet.active_cell_color = defaults["active_cell_color"]
    if defaults.get("tui_foreground_color"):
        sheet.tui_foreground_color = defaults["tui_foreground_color"]
    if defaults.get("tui_background_color"):
        sheet.tui_background_color = defaults["tui_background_color"]
    if defaults.get("row_header_foreground_color"):
        sheet.row_header_foreground_color = defaults["row_header_foreground_color"]
    if defaults.get("row_header_background_color"):
        sheet.row_header_background_color = defaults["row_header_background_color"]
    if defaults.get("column_header_foreground_color"):
        sheet.column_header_foreground_color = defaults["column_header_foreground_color"]
    if defaults.get("column_header_background_color"):
        sheet.column_header_background_color = defaults["column_header_background_color"]
    if defaults.get("sheet_foreground_color"):
        sheet.sheet_foreground_color = defaults["sheet_foreground_color"]
    if defaults.get("sheet_background_color"):
        sheet.sheet_background_color = defaults["sheet_background_color"]
    if defaults.get("formula_coloration"):
        sheet.formula_coloration = defaults["formula_coloration"].lower() in {"1", "true", "yes", "on"}
    if defaults.get("formula_foreground_color"):
        sheet.formula_foreground_color = defaults["formula_foreground_color"]
    if defaults.get("language"):
        sheet.language = defaults["language"]
    if defaults.get("protected_foreground_color"):
        sheet.protected_foreground_color = defaults["protected_foreground_color"]
    if defaults.get("protected_background_color"):
        sheet.protected_background_color = defaults["protected_background_color"]


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


def save_xlsx(sheet: Spreadsheet, path: Path) -> None:
    path = path.expanduser()
    path.parent.mkdir(parents=True, exist_ok=True)
    max_row, max_col = _used_bounds(sheet)
    sheet_xml = _xlsx_sheet_xml(sheet, max_row, max_col)
    created = _iso_timestamp()
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
""",
        )
        archive.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
""",
        )
        archive.writestr(
            "xl/workbook.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
""",
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
""",
        )
        archive.writestr(
            "xl/styles.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><name val="Courier New"/><sz val="11"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
""",
        )
        archive.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        archive.writestr(
            "docProps/core.xml",
            f"""<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>tui-ss export</dc:title>
  <dc:creator>tui-ss</dc:creator>
  <cp:lastModifiedBy>tui-ss</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{created}</dcterms:modified>
</cp:coreProperties>
""",
        )
        archive.writestr(
            "docProps/app.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>tui-ss</Application>
</Properties>
""",
        )


def save_ods(sheet: Spreadsheet, path: Path) -> None:
    path = path.expanduser()
    path.parent.mkdir(parents=True, exist_ok=True)
    max_row, max_col = _used_bounds(sheet)
    created = _iso_timestamp()
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr(
            "mimetype",
            "application/vnd.oasis.opendocument.spreadsheet",
            compress_type=zipfile.ZIP_STORED,
        )
        archive.writestr(
            "content.xml",
            _ods_content_xml(sheet, max_row, max_col),
            compress_type=zipfile.ZIP_DEFLATED,
        )
        archive.writestr(
            "meta.xml",
            f"""<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
 xmlns:dc="http://purl.org/dc/elements/1.1/">
  <office:meta>
    <meta:generator>tui-ss</meta:generator>
    <dc:creator>tui-ss</dc:creator>
    <meta:creation-date>{created}</meta:creation-date>
  </office:meta>
</office:document-meta>
""",
            compress_type=zipfile.ZIP_DEFLATED,
        )
        archive.writestr(
            "styles.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 office:version="1.3">
  <office:styles/>
</office:document-styles>
""",
            compress_type=zipfile.ZIP_DEFLATED,
        )
        archive.writestr(
            "settings.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<office:document-settings
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 office:version="1.3">
  <office:settings/>
</office:document-settings>
""",
            compress_type=zipfile.ZIP_DEFLATED,
        )
        archive.writestr(
            "META-INF/manifest.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest
 xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
 manifest:version="1.3">
  <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.spreadsheet" manifest:full-path="/"/>
  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>
  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="styles.xml"/>
  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="meta.xml"/>
  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="settings.xml"/>
</manifest:manifest>
""",
            compress_type=zipfile.ZIP_DEFLATED,
        )


def _used_bounds(sheet: Spreadsheet) -> tuple[int, int]:
    max_row = 0
    max_col = 0
    saw_any = False
    for row, col, raw in sheet.iter_cells():
        if raw:
            max_row = max(max_row, row)
            max_col = max(max_col, col)
            saw_any = True
    return (max_row, max_col) if saw_any else (0, 0)


def _xlsx_sheet_xml(sheet: Spreadsheet, max_row: int, max_col: int) -> str:
    rows: list[str] = []
    for row in range(max_row + 1):
        cells: list[str] = []
        for col in range(max_col + 1):
            raw = sheet.get_raw(row, col)
            if not raw:
                continue
            ref = f"{column_label(col)}{row + 1}"
            cells.append(_xlsx_cell_xml(ref, raw))
        if cells:
            rows.append(f'<row r="{row + 1}">{"".join(cells)}</row>')
    dimension = f"A1:{column_label(max_col)}{max_row + 1}"
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="{dimension}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>{"".join(rows)}</sheetData>
</worksheet>
"""


def _xlsx_cell_xml(ref: str, raw: str) -> str:
    value = _export_raw_value(raw)
    if is_formula_text(raw):
        formula = escape(raw[1:])
        return f'<c r="{ref}"><f>{formula}</f></c>'
    if _looks_numeric(value):
        return f'<c r="{ref}"><v>{escape(value)}</v></c>'
    return f'<c r="{ref}" t="inlineStr"><is><t>{escape(value)}</t></is></c>'


def _ods_content_xml(sheet: Spreadsheet, max_row: int, max_col: int) -> str:
    row_xml: list[str] = []
    for row in range(max_row + 1):
        cells: list[str] = []
        for col in range(max_col + 1):
            raw = sheet.get_raw(row, col)
            cells.append(_ods_cell_xml(raw))
        row_xml.append(f'<table:table-row>{"".join(cells)}</table:table-row>')
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 office:version="1.3">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Sheet1">{"".join(row_xml)}</table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>
"""


def _ods_cell_xml(raw: str) -> str:
    if not raw:
        return "<table:table-cell/>"
    value = _export_raw_value(raw)
    if is_formula_text(raw):
        formula_text = raw[1:]
        formula_attr = escape(_ods_formula(formula_text))
        if _looks_numeric(value):
            return (
                f'<table:table-cell table:formula="{formula_attr}" office:value-type="float" office:value="0">'
                f"<text:p>{escape(formula_text)}</text:p></table:table-cell>"
            )
        return (
            f'<table:table-cell table:formula="{formula_attr}" office:value-type="string">'
            f"<text:p>{escape(formula_text)}</text:p></table:table-cell>"
        )
    if _looks_numeric(value):
        return f'<table:table-cell office:value-type="float" office:value="{escape(value)}"><text:p>{escape(value)}</text:p></table:table-cell>'
    return f'<table:table-cell office:value-type="string"><text:p>{escape(value)}</text:p></table:table-cell>'


def _ods_formula(formula_text: str) -> str:
    converted = formula_text.replace(",", ";")
    tokens = []
    index = 0
    while index < len(converted):
        char = converted[index]
        if char.isalpha() or char == "$":
            end = index
            while end < len(converted) and (converted[end].isalnum() or converted[end] in "$:_"):
                end += 1
            token = converted[index:end]
            if any(ch.isdigit() for ch in token) and any(ch.isalpha() for ch in token):
                token = token.replace(":", "]:.[")
                token = f"[.{token}]"
            tokens.append(token)
            index = end
            continue
        tokens.append(char)
        index += 1
    return f"of:={''.join(tokens)}"


def _export_raw_value(raw: str) -> str:
    return unescape_literal_text(raw)


def _looks_numeric(text: str) -> bool:
    if not text:
        return False
    try:
        float(text)
    except ValueError:
        return False
    return True


def _iso_timestamp() -> str:
    return datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")
