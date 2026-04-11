#!/usr/bin/env python3
"""Curses spreadsheet application."""

from __future__ import annotations

import argparse
import base64
import curses
import json
import csv
import re
import shlex
import shutil
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from .commands import (
    ALIASES,
    ADVANCED_COMMAND_MENU_OPTIONS,
    COMMAND_MENU_OPTIONS,
    COMMAND_DESCRIPTIONS,
    HELP_TOPICS,
    get_command_help_lines,
    get_formula_help_lines,
    get_key_help_lines,
    LANGUAGE_CODES,
    LANGUAGE_OPTIONS,
    parse_command,
    tr,
)
from .formulas import (
    Evaluator,
    FormulaError,
    format_date_text,
    format_time_text,
    is_formula_text,
    parse_date_text,
    parse_time_text,
    normalize_date_text,
    normalize_time_text,
    shift_formula_references,
)
from .model import Spreadsheet, column_label, parse_cell_reference
from .storage import load_app_settings, load_sheet, save_app_settings, save_pdf_text, save_sheet

APP_NAME = "tui-ss"
DEFAULT_PATH = Path.home() / "scripts" / "tui-ss" / "sheets" / "autosave.tss"
DEFAULT_SETTINGS_PATH = Path.home() / ".config" / "tui-ss" / "tui-ss-settings.toml"
THEMES = ["blue", "cyan", "magenta", "purple", "white", "yellow"]
ACTIVE_CELL_COLORS = ["yellow", "pink", "orange", "white", "lightblue", "cornflower", "lightgrey"]
PROTECTED_COLOR_OPTIONS = ["black", "white", "yellow", "pink", "palepink", "orange", "lightblue", "cornflower", "lightgrey", "blue", "cyan", "green", "magenta", "red"]
TUI_COLOR_OPTIONS = ["black", "white", "yellow", "pink", "palepink", "orange", "lightblue", "cornflower", "lightgrey", "primrose", "gold", "darkgreen", "blue", "cyan", "green", "magenta", "red"]
FORMULA_COLOR_OPTIONS = ["green", "yellow", "cyan", "magenta", "orange", "lightblue", "cornflower", "white", "red", "blue"]
FORMULA_COLOR_SETTING_OPTIONS = ["off"] + FORMULA_COLOR_OPTIONS
FORMAT_STYLES = ["accounting", "background", "bold", "clear-format", "currency", "date", "fixed", "int", "italic", "negative", "percent", "row-background", "sci", "text", "time", "underline"]
TIME_FORMATS = ["24h", "24h-seconds", "12h", "12h-seconds"]
CURRENCY_SYMBOLS = ["£", "€", "$", "¥", "₹"]
DATE_FORMATS = ["european", "uk", "us", "ansi"]
BACKGROUND_COLORS = ["blue", "cyan", "green", "magenta", "none", "red", "white", "yellow"]
JUSTIFY_OPTIONS = ["left", "centre", "right"]
FILE_BROWSER_SORT_OPTIONS = ["name", "time", "type"]
SHEET_BG_OPTIONS = ["none", "white", "lightgrey", "grey", "darkgrey"]
SHEET_FG_OPTIONS = ["none", "white", "yellow", "primrose", "lightgreen"]
COLOR_PAIR_TEXT = 1
COLOR_PAIR_FORMULA = 2
COLOR_PAIR_HEADER = 3
COLOR_PAIR_BAR = 4
COLOR_PAIR_GRID = 5
COLOR_PAIR_MENU_SELECTED = 6
COLOR_PAIR_GRID_ROW = 7
COLOR_PAIR_SELECTION = 8
COLOR_PAIR_ROW_HEADER = 9
COLOR_PAIR_NEGATIVE = 10
THEME_COLOR_MAP = {
    "white": curses.COLOR_WHITE,
    "cyan": curses.COLOR_CYAN,
    "yellow": curses.COLOR_YELLOW,
    "magenta": curses.COLOR_MAGENTA,
    "blue": curses.COLOR_BLUE,
    "purple": curses.COLOR_MAGENTA,
}
CUSTOM_PURPLE_COLOR_ID = 16
CUSTOM_ORANGE_COLOR_ID = 17
CUSTOM_PINK_COLOR_ID = 18
CUSTOM_LIGHTBLUE_COLOR_ID = 19
CUSTOM_CORNFLOWER_COLOR_ID = 20
CUSTOM_LIGHTGREY_COLOR_ID = 21
CUSTOM_PALEPINK_COLOR_ID = 22
CUSTOM_PRIMROSE_COLOR_ID = 23
CUSTOM_GOLD_COLOR_ID = 24
CUSTOM_DARKGREEN_COLOR_ID = 25
CUSTOM_LIGHTGREEN_COLOR_ID = 26
CUSTOM_GREY_COLOR_ID = 27
CUSTOM_DARKGREY_COLOR_ID = 28
CUSTOM_HEX_COLOR_START = 40
CLIPBOARD_MARKER = "TUI-SS-CLIP:"
RECENT_FILES_LIMIT = 10
HEX_COLOR_RE = re.compile(r"^#?[0-9A-Fa-f]{6}$")
BACKGROUND_COLOR_MAP = {
    "blue": curses.COLOR_BLUE,
    "cyan": curses.COLOR_CYAN,
    "green": curses.COLOR_GREEN,
    "magenta": curses.COLOR_MAGENTA,
    "red": curses.COLOR_RED,
    "yellow": curses.COLOR_YELLOW,
    "white": curses.COLOR_WHITE,
}
FORMULA_SIGNATURES = {
    "ABS": "ABS(value)",
    "AND": "AND(value1, value2, ...)",
    "AVERAGE": "AVERAGE(range)",
    "AVG": "AVG(range)",
    "CONCAT": "CONCAT(text1, text2, ...)",
    "COS": "COS(value)",
    "COUNT": "COUNT(range)",
    "DATE": "DATE(year, month, day)",
    "DATEDIFF": "DATEDIFF(start_date, end_date)",
    "DAY": "DAY(date)",
    "HLOOKUP": "HLOOKUP(value, range, row_index)",
    "HOUR": "HOUR(time)",
    "IF": "IF(condition, then_value, else_value)",
    "IFERROR": "IFERROR(value, fallback)",
    "INDEX": "INDEX(range, index)",
    "INT": "INT(value)",
    "LEFT": "LEFT(text, count)",
    "LEN": "LEN(text)",
    "LOOKUP": "LOOKUP(value, lookup_range, result_range)",
    "MATCH": "MATCH(value, range)",
    "MAX": "MAX(range)",
    "MID": "MID(text, start, count)",
    "MINUTE": "MINUTE(time)",
    "MIN": "MIN(range)",
    "MOD": "MOD(value, divisor)",
    "MONTH": "MONTH(date)",
    "NOT": "NOT(value)",
    "OR": "OR(value1, value2, ...)",
    "RIGHT": "RIGHT(text, count)",
    "ROUND": "ROUND(value, digits)",
    "SECOND": "SECOND(time)",
    "SIN": "SIN(value)",
    "SQRT": "SQRT(value)",
    "SUM": "SUM(range)",
    "TAN": "TAN(value)",
    "TIME": "TIME(hour, minute, second)",
    "TIMEVALUE": "TIMEVALUE(value)",
    "TEXT": 'TEXT(value, "format")',
    "NOW": "NOW()",
    "TODAY": "TODAY()",
    "VALUE": "VALUE(text)",
    "VLOOKUP": "VLOOKUP(value, range, column_index)",
    "WEEKDAY": "WEEKDAY(date)",
    "YEAR": "YEAR(date)",
}
FORMULA_ARGUMENT_NAMES = {
    "ABS": ["value"],
    "AND": ["value1", "value2", "..."],
    "AVERAGE": ["range"],
    "AVG": ["range"],
    "CONCAT": ["text1", "text2", "..."],
    "COS": ["value"],
    "COUNT": ["range"],
    "DATE": ["year", "month", "day"],
    "DATEDIFF": ["start_date", "end_date"],
    "DAY": ["date"],
    "HLOOKUP": ["value", "range", "row_index"],
    "HOUR": ["time"],
    "IF": ["condition", "then_value", "else_value"],
    "IFERROR": ["value", "fallback"],
    "INDEX": ["range", "index"],
    "INT": ["value"],
    "LEFT": ["text", "count"],
    "LEN": ["text"],
    "LOOKUP": ["value", "lookup_range", "result_range"],
    "MATCH": ["value", "range"],
    "MAX": ["range"],
    "MID": ["text", "start", "count"],
    "MINUTE": ["time"],
    "MIN": ["range"],
    "MOD": ["value", "divisor"],
    "MONTH": ["date"],
    "NOT": ["value"],
    "OR": ["value1", "value2", "..."],
    "RIGHT": ["text", "count"],
    "ROUND": ["value", "digits"],
    "SECOND": ["time"],
    "SIN": ["value"],
    "SQRT": ["value"],
    "SUM": ["range"],
    "TAN": ["value"],
    "TIME": ["hour", "minute", "second"],
    "TIMEVALUE": ["value"],
    "TEXT": ["value", "format"],
    "NOW": [],
    "TODAY": [],
    "VALUE": ["text"],
    "VLOOKUP": ["value", "range", "column_index"],
    "WEEKDAY": ["date"],
    "YEAR": ["date"],
}
FORMULA_REF_PATTERN = re.compile(r"(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$")


@dataclass
class TabState:
    sheet: Spreadsheet
    path: Path | None
    name: str | None
    dirty: bool
    current_row: int
    current_col: int
    row_offset: int
    col_offset: int
    selection_anchor: tuple[int, int] | None
    selection_range: tuple[int, int, int, int] | None
    mouse_dragging: bool
    undo_stack: list[dict[str, object]]
    redo_stack: list[dict[str, object]]


def build_stamp() -> str:
    try:
        return datetime.fromtimestamp(Path(__file__).stat().st_mtime).strftime("%y%m%d-%H:%M")
    except OSError:
        return datetime.now().strftime("%y%m%d-%H:%M")


def parse_cell_or_current(token: str | None, row: int, col: int) -> tuple[int, int]:
    if not token:
        return row, col
    return parse_cell_reference(token)


def parse_range_spec(
    spec: str,
    current_row: int,
    current_col: int,
    max_rows: int = 100,
    max_cols: int = 52,
) -> tuple[int, int, int, int]:
    token = spec.strip().upper() if spec else ""
    if not token or token == ".":
        return current_row, current_col, current_row, current_col
    if token.isalpha():
        col = parse_cell_reference(f"{token}1")[1]
        return 0, col, max_rows - 1, col
    if token.isdigit():
        row = max(0, int(token) - 1)
        return row, 0, row, max_cols - 1
    if ":" not in token:
        row, col = parse_cell_reference(token)
        return row, col, row, col
    start_text, end_text = token.split(":", 1)
    if start_text.isalpha() and end_text.isalpha():
        start_col = parse_cell_reference(f"{start_text}1")[1]
        end_col = parse_cell_reference(f"{end_text}1")[1]
        col_lo, col_hi = sorted((start_col, end_col))
        return 0, col_lo, max_rows - 1, col_hi
    if start_text.isdigit() and end_text.isdigit():
        start_row = max(0, int(start_text) - 1)
        end_row = max(0, int(end_text) - 1)
        row_lo, row_hi = sorted((start_row, end_row))
        return row_lo, 0, row_hi, max_cols - 1
    if start_text.isalpha() and end_text.isdigit():
        start_row, start_col = parse_cell_reference(f"{start_text}{end_text}")
        return start_row, start_col, max_rows - 1, start_col
    if start_text.isdigit() and end_text.isalpha():
        end_row, end_col = parse_cell_reference(f"{end_text}{start_text}")
        return end_row, end_col, end_row, max_cols - 1
    start_row, start_col = parse_cell_reference(start_text)
    end_row, end_col = parse_cell_reference(end_text)
    row_lo, row_hi = sorted((start_row, end_row))
    col_lo, col_hi = sorted((start_col, end_col))
    return row_lo, col_lo, row_hi, col_hi


def should_auto_right_align(raw: str) -> bool:
    text = raw.strip()
    if not text:
        return False
    if text.startswith("'"):
        return False
    if is_formula_text(text):
        return True
    try:
        float(text.replace(",", ""))
        return True
    except ValueError:
        if parse_date_text(text) is not None:
            return True
        if parse_time_text(text) is not None:
            return True
        return False


class SpreadsheetApp:
    def __init__(self, stdscr, path: Path | None = None, settings_path: Path | None = None) -> None:
        self.stdscr = stdscr
        self.sheet = Spreadsheet(rows=100, cols=52)
        self.evaluator = Evaluator(self.sheet)
        self.current_row = 0
        self.current_col = 0
        self.row_offset = 0
        self.col_offset = 0
        self.message = "Press / for SuperCalc-style commands, Enter to edit."
        self.path = path
        self.settings_path = settings_path or DEFAULT_SETTINGS_PATH
        self.running = True
        self.colors_ready = False
        self.dynamic_color_pairs: dict[tuple[int, int], int] = {}
        self.dynamic_colors: dict[str, int] = {}
        self.next_dynamic_pair = 20
        self.selection_anchor: tuple[int, int] | None = None
        self.selection_range: tuple[int, int, int, int] | None = None
        self.mouse_dragging = False
        self.dirty = False
        self.clipboard_cells: list[tuple[int, int, str, str, str, str, str, str, bool, bool]] = []
        self.clipboard_size: tuple[int, int] = (0, 0)
        self.clipboard_origin: tuple[int, int] = (0, 0)
        self.prefer_internal_clipboard = False
        self.undo_stack: list[dict[str, object]] = []
        self.redo_stack: list[dict[str, object]] = []
        self.max_history = 100
        self.tabs: list[TabState] = []
        self.current_tab_index = 0
        self.recent_files: list[str] = []
        self.cursor_positions: dict[str, str] = {}
        self.command_hint_visible = True
        self.key_overlay_visible = False
        self.raw_sheet_view = False
        self._load_global_settings()
        saved_row, saved_col = self._saved_cursor_position(self.path)
        self.current_row = min(self.sheet.rows - 1, saved_row)
        self.current_col = min(self.sheet.cols - 1, saved_col)
        self.tabs.append(self._capture_tab_state())
        self._scroll_into_view()

    def run(self) -> int:
        curses.curs_set(0)
        self.stdscr.keypad(True)
        curses.raw()
        mouse_events = curses.BUTTON1_PRESSED | curses.BUTTON1_CLICKED
        mouse_events |= getattr(curses, "BUTTON1_RELEASED", 0)
        mouse_events |= getattr(curses, "REPORT_MOUSE_POSITION", 0)
        mouse_events |= getattr(curses, "BUTTON4_PRESSED", 0)
        mouse_events |= getattr(curses, "BUTTON5_PRESSED", 0)
        curses.mousemask(mouse_events)
        curses.mouseinterval(0)
        self._set_bracketed_paste(True)
        self._setup_colors()
        try:
            while self.running:
                self.draw()
                key = self.stdscr.getch()
                self.handle_key(key)
        finally:
            self._store_current_tab_state()
            self._save_global_settings()
            self._set_bracketed_paste(False)
            curses.noraw()
        return 0

    def _set_bracketed_paste(self, enabled: bool) -> None:
        sequence = "\x1b[?2004h" if enabled else "\x1b[?2004l"
        sys.stdout.write(sequence)
        sys.stdout.flush()

    def _setup_colors(self) -> None:
        if not curses.has_colors():
            return
        curses.start_color()
        curses.use_default_colors()
        self.colors_ready = True
        self._refresh_theme_colors()

    def _refresh_theme_colors(self) -> None:
        if not self.colors_ready:
            return
        self.dynamic_color_pairs.clear()
        self.dynamic_colors.clear()
        self.next_dynamic_pair = 20
        text_color = self._sheet_foreground_color()
        bar_foreground = self._tui_foreground_color()
        bar_background = self._tui_background_color()
        column_header_foreground = self._named_color(self.sheet.column_header_foreground_color, curses.COLOR_YELLOW)
        column_header_background = self._named_color(self.sheet.column_header_background_color, curses.COLOR_BLACK)
        row_header_foreground = self._named_color(self.sheet.row_header_foreground_color, curses.COLOR_YELLOW)
        row_header_background = self._named_color(self.sheet.row_header_background_color, curses.COLOR_BLACK)
        sheet_background = self._sheet_background_color()
        curses.init_pair(COLOR_PAIR_TEXT, text_color, sheet_background)
        curses.init_pair(COLOR_PAIR_FORMULA, self._formula_foreground_color(), sheet_background)
        curses.init_pair(COLOR_PAIR_HEADER, column_header_foreground, column_header_background)
        curses.init_pair(COLOR_PAIR_BAR, bar_foreground, bar_background)
        curses.init_pair(COLOR_PAIR_GRID, curses.COLOR_BLACK, -1)
        curses.init_pair(COLOR_PAIR_GRID_ROW, text_color, -1)
        curses.init_pair(COLOR_PAIR_MENU_SELECTED, self._selection_foreground_color(), self._selection_background_color())
        curses.init_pair(COLOR_PAIR_SELECTION, self._selection_foreground_color(), self._selection_background_color())
        curses.init_pair(COLOR_PAIR_ROW_HEADER, row_header_foreground, row_header_background)
        curses.init_pair(COLOR_PAIR_NEGATIVE, curses.COLOR_RED, sheet_background)

    def _settings_payload(self) -> dict[str, str]:
        return {
            "theme_name": self.sheet.theme_name,
            "date_format": self.sheet.date_format,
            "time_format": self.sheet.time_format,
            "active_cell_color": self.sheet.active_cell_color,
            "tui_foreground_color": self.sheet.tui_foreground_color,
            "tui_background_color": self.sheet.tui_background_color,
            "row_header_foreground_color": self.sheet.row_header_foreground_color,
            "row_header_background_color": self.sheet.row_header_background_color,
            "column_header_foreground_color": self.sheet.column_header_foreground_color,
            "column_header_background_color": self.sheet.column_header_background_color,
            "sheet_foreground_color": self.sheet.sheet_foreground_color,
            "sheet_background_color": self.sheet.sheet_background_color,
            "formula_coloration": "on" if self.sheet.formula_coloration else "off",
            "formula_foreground_color": self.sheet.formula_foreground_color,
            "language": self.sheet.language,
            "protected_foreground_color": self.sheet.protected_foreground_color,
            "protected_background_color": self.sheet.protected_background_color,
            "recent_files_json": json.dumps(self.recent_files),
            "cursor_positions_json": json.dumps(self.cursor_positions),
        }

    def _apply_settings_payload(self, settings: dict[str, str]) -> None:
        def _is_named_or_hex(value: str | None, options: list[str]) -> bool:
            return bool(value) and (value in options or bool(HEX_COLOR_RE.match(value)))

        theme_name = settings.get("theme_name")
        if theme_name in THEMES:
            self.sheet.theme_name = theme_name
        raw_date_format = settings.get("date_format")
        if raw_date_format:
            if raw_date_format in DATE_FORMATS:
                self.sheet.date_format = f"date:{raw_date_format}"
            elif raw_date_format.startswith("date:"):
                self.sheet.date_format = raw_date_format
        raw_time_format = settings.get("time_format")
        if raw_time_format:
            if raw_time_format in TIME_FORMATS:
                self.sheet.time_format = f"time:{raw_time_format}"
            elif raw_time_format.startswith("time:"):
                self.sheet.time_format = raw_time_format
        active_cell_color = settings.get("active_cell_color")
        if active_cell_color in ACTIVE_CELL_COLORS:
            self.sheet.active_cell_color = active_cell_color
        tui_fg = settings.get("tui_foreground_color")
        if tui_fg in TUI_COLOR_OPTIONS:
            self.sheet.tui_foreground_color = tui_fg
        tui_bg = settings.get("tui_background_color")
        if tui_bg in TUI_COLOR_OPTIONS:
            self.sheet.tui_background_color = tui_bg
        row_header_fg = settings.get("row_header_foreground_color")
        if _is_named_or_hex(row_header_fg, TUI_COLOR_OPTIONS):
            self.sheet.row_header_foreground_color = str(row_header_fg)
        row_header_bg = settings.get("row_header_background_color")
        if _is_named_or_hex(row_header_bg, TUI_COLOR_OPTIONS):
            self.sheet.row_header_background_color = str(row_header_bg)
        column_header_fg = settings.get("column_header_foreground_color")
        if _is_named_or_hex(column_header_fg, TUI_COLOR_OPTIONS):
            self.sheet.column_header_foreground_color = str(column_header_fg)
        column_header_bg = settings.get("column_header_background_color")
        if _is_named_or_hex(column_header_bg, TUI_COLOR_OPTIONS):
            self.sheet.column_header_background_color = str(column_header_bg)
        sheet_fg = settings.get("sheet_foreground_color")
        if sheet_fg in SHEET_FG_OPTIONS:
            self.sheet.sheet_foreground_color = sheet_fg
        sheet_bg = settings.get("sheet_background_color")
        if sheet_bg in SHEET_BG_OPTIONS:
            self.sheet.sheet_background_color = sheet_bg
        formula_coloration = settings.get("formula_coloration")
        if formula_coloration:
            self.sheet.formula_coloration = formula_coloration.lower() in {"1", "true", "yes", "on"}
        formula_foreground_color = settings.get("formula_foreground_color")
        if formula_foreground_color in FORMULA_COLOR_OPTIONS:
            self.sheet.formula_foreground_color = formula_foreground_color
        language = settings.get("language")
        if language in LANGUAGE_CODES.values():
            self.sheet.language = language
        protected_fg = settings.get("protected_foreground_color")
        if protected_fg in PROTECTED_COLOR_OPTIONS:
            self.sheet.protected_foreground_color = protected_fg
        protected_bg = settings.get("protected_background_color")
        if protected_bg in PROTECTED_COLOR_OPTIONS:
            self.sheet.protected_background_color = protected_bg
        recent_files_json = settings.get("recent_files_json")
        if recent_files_json:
            try:
                payload = json.loads(recent_files_json)
            except json.JSONDecodeError:
                payload = []
            if isinstance(payload, list):
                self.recent_files = [str(item) for item in payload if str(item).strip()][:RECENT_FILES_LIMIT]
        cursor_positions_json = settings.get("cursor_positions_json")
        if cursor_positions_json:
            try:
                payload = json.loads(cursor_positions_json)
            except json.JSONDecodeError:
                payload = {}
            if isinstance(payload, dict):
                self.cursor_positions = {str(key): str(value) for key, value in payload.items() if str(key).strip() and str(value).strip()}

    def _load_global_settings(self) -> None:
        settings = load_app_settings(self.settings_path)
        if settings:
            self._apply_settings_payload(settings)
        save_app_settings(self.settings_path, self._settings_payload())

    def _save_global_settings(self) -> None:
        save_app_settings(self.settings_path, self._settings_payload())

    def _remember_recent_file(self, path: Path) -> None:
        text = str(path.expanduser().resolve())
        self.recent_files = [item for item in self.recent_files if item != text]
        self.recent_files.insert(0, text)
        self.recent_files = self.recent_files[:RECENT_FILES_LIMIT]
        self._save_global_settings()

    def _cursor_position_key(self, path: Path | None) -> str:
        if path is None:
            return "__untitled__"
        return str(path.expanduser().resolve())

    def _remember_cursor_position(self, path: Path | None, row: int, col: int) -> None:
        self.cursor_positions[self._cursor_position_key(path)] = f"{max(0, row)},{max(0, col)}"

    def _saved_cursor_position(self, path: Path | None) -> tuple[int, int]:
        raw = self.cursor_positions.get(self._cursor_position_key(path), "")
        if not raw:
            return 0, 0
        try:
            row_text, col_text = raw.split(",", 1)
            return max(0, int(row_text)), max(0, int(col_text))
        except ValueError:
            return 0, 0

    def _theme_text_color(self) -> int:
        if self.sheet.theme_name == "purple":
            custom_purple = self._custom_purple_color()
            if custom_purple is not None:
                return custom_purple
        return THEME_COLOR_MAP.get(self.sheet.theme_name, curses.COLOR_WHITE)

    def _tui_foreground_color(self) -> int:
        return self._named_color(self.sheet.tui_foreground_color, self._theme_text_color())

    def _tui_background_color(self) -> int:
        return self._named_color(self.sheet.tui_background_color, curses.COLOR_BLACK)

    def _sheet_foreground_color(self) -> int:
        if self.sheet.sheet_foreground_color == "none":
            return self._theme_text_color()
        return self._named_color(self.sheet.sheet_foreground_color, self._theme_text_color())

    def _sheet_background_color(self) -> int:
        if self.sheet.sheet_background_color == "none":
            return -1
        return self._named_color(self.sheet.sheet_background_color, -1)

    def _custom_purple_color(self) -> int | None:
        if not self.colors_ready:
            return None
        if not hasattr(curses, "can_change_color") or not curses.can_change_color():
            return None
        if curses.COLORS <= CUSTOM_PURPLE_COLOR_ID:
            return None
        try:
            curses.init_color(CUSTOM_PURPLE_COLOR_ID, 400, 0, 1000)
        except curses.error:
            return None
        return CUSTOM_PURPLE_COLOR_ID

    def _custom_orange_color(self) -> int | None:
        if not self.colors_ready:
            return None
        if not hasattr(curses, "can_change_color") or not curses.can_change_color():
            return None
        if curses.COLORS <= CUSTOM_ORANGE_COLOR_ID:
            return None
        try:
            curses.init_color(CUSTOM_ORANGE_COLOR_ID, 1000, 400, 0)
        except curses.error:
            return None
        return CUSTOM_ORANGE_COLOR_ID

    def _custom_named_color(self, color_id: int, red: int, green: int, blue: int) -> int | None:
        if not self.colors_ready:
            return None
        if not hasattr(curses, "can_change_color") or not curses.can_change_color():
            return None
        if curses.COLORS <= color_id:
            return None
        try:
            curses.init_color(color_id, red, green, blue)
        except curses.error:
            return None
        return color_id

    def _named_color(self, name: str, default: int = -1) -> int:
        if HEX_COLOR_RE.match(name):
            return self._hex_color(name) or default
        if name == "black":
            return curses.COLOR_BLACK
        if name == "white":
            return curses.COLOR_WHITE
        if name == "yellow":
            return curses.COLOR_YELLOW
        if name == "blue":
            return curses.COLOR_BLUE
        if name == "cyan":
            return curses.COLOR_CYAN
        if name == "green":
            return curses.COLOR_GREEN
        if name == "magenta":
            return curses.COLOR_MAGENTA
        if name == "red":
            return curses.COLOR_RED
        if name == "pink":
            return self._custom_named_color(CUSTOM_PINK_COLOR_ID, 1000, 500, 700) or curses.COLOR_MAGENTA
        if name == "palepink":
            return self._custom_named_color(CUSTOM_PALEPINK_COLOR_ID, 1000, 850, 900) or curses.COLOR_WHITE
        if name == "orange":
            return self._custom_orange_color() or curses.COLOR_YELLOW
        if name == "lightblue":
            return self._custom_named_color(CUSTOM_LIGHTBLUE_COLOR_ID, 500, 700, 1000) or curses.COLOR_CYAN
        if name == "cornflower":
            return self._custom_named_color(CUSTOM_CORNFLOWER_COLOR_ID, 392, 584, 929) or curses.COLOR_BLUE
        if name == "lightgrey":
            return self._custom_named_color(CUSTOM_LIGHTGREY_COLOR_ID, 800, 800, 800) or curses.COLOR_WHITE
        if name == "primrose":
            return self._custom_named_color(CUSTOM_PRIMROSE_COLOR_ID, 1000, 970, 639) or curses.COLOR_YELLOW
        if name == "gold":
            return self._custom_named_color(CUSTOM_GOLD_COLOR_ID, 1000, 843, 0) or curses.COLOR_YELLOW
        if name == "darkgreen":
            return self._custom_named_color(CUSTOM_DARKGREEN_COLOR_ID, 0, 392, 0) or curses.COLOR_GREEN
        if name == "lightgreen":
            return self._custom_named_color(CUSTOM_LIGHTGREEN_COLOR_ID, 564, 933, 564) or curses.COLOR_GREEN
        if name == "grey":
            return self._custom_named_color(CUSTOM_GREY_COLOR_ID, 500, 500, 500) or curses.COLOR_WHITE
        if name == "darkgrey":
            return self._custom_named_color(CUSTOM_DARKGREY_COLOR_ID, 250, 250, 250) or curses.COLOR_BLACK
        return default

    def _hex_color(self, value: str) -> int | None:
        if not self.colors_ready:
            return None
        normalized = value.upper()
        if not normalized.startswith("#"):
            normalized = f"#{normalized}"
        existing = self.dynamic_colors.get(normalized)
        if existing is not None:
            return existing
        red = int(normalized[1:3], 16)
        green = int(normalized[3:5], 16)
        blue = int(normalized[5:7], 16)
        if hasattr(curses, "can_change_color") and curses.can_change_color():
            color_id = CUSTOM_HEX_COLOR_START + len(self.dynamic_colors)
            if color_id < curses.COLORS:
                try:
                    curses.init_color(
                        color_id,
                        int(red * 1000 / 255),
                        int(green * 1000 / 255),
                        int(blue * 1000 / 255),
                    )
                    self.dynamic_colors[normalized] = color_id
                    return color_id
                except curses.error:
                    pass
        # Fallback: approximate to a basic curses color.
        candidates = [
            ("black", (0, 0, 0), curses.COLOR_BLACK),
            ("red", (255, 0, 0), curses.COLOR_RED),
            ("green", (0, 255, 0), curses.COLOR_GREEN),
            ("yellow", (255, 255, 0), curses.COLOR_YELLOW),
            ("blue", (0, 0, 255), curses.COLOR_BLUE),
            ("magenta", (255, 0, 255), curses.COLOR_MAGENTA),
            ("cyan", (0, 255, 255), curses.COLOR_CYAN),
            ("white", (255, 255, 255), curses.COLOR_WHITE),
        ]
        nearest = min(
            candidates,
            key=lambda item: (red - item[1][0]) ** 2 + (green - item[1][1]) ** 2 + (blue - item[1][2]) ** 2,
        )
        self.dynamic_colors[normalized] = nearest[2]
        return nearest[2]

    def draw(self) -> None:
        self.stdscr.erase()
        height, width = self.stdscr.getmaxyx()
        top_grid_row, grid_height, row_header_width, visible_columns = self._grid_layout(height, width)
        self.stdscr.addnstr(0, 0, self._tabs_line(width), width - 1, self._bar_attr(bold=True))
        self.stdscr.addnstr(1, 0, self._title_line(width), width - 1, self._bar_attr(bold=True))
        self.stdscr.addnstr(2, 0, self._top_formula_line(width), width - 1, self._bar_attr())
        header_y = top_grid_row - 1
        display_lines = self._display_lines(grid_height, visible_columns)
        self._draw_grid(header_y, grid_height, row_header_width, visible_columns)
        self._draw_column_headers(header_y, visible_columns)

        visible_rows = [row for kind, row, _edge in display_lines if kind == "row"]
        for screen_offset, (kind, row, _edge) in enumerate(display_lines):
            if kind != "row":
                continue
            y = top_grid_row + screen_offset
            row_header_attr = self._row_header_attr()
            self.stdscr.addnstr(y, 0, (" " * max(0, row_header_width - 1)), row_header_width - 1, row_header_attr)
            self.stdscr.addnstr(y, 0, f"{row + 1:>5}", row_header_width - 1, row_header_attr)
            column_index = 0
            while column_index < len(visible_columns):
                col, x, col_width = visible_columns[column_index]
                render_width, spill_to_index = self._spill_width(row, column_index, visible_columns)
                text = self._cell_text(row, col, render_width)
                in_selection = self._cell_in_selection(row, col)
                attr = self._cell_attr(row, col)
                if in_selection:
                    attr = self._selection_cell_attr(row, col)
                if (row, col) == (self.current_row, self.current_col):
                    attr = self._active_cell_attr(row, col)
                if row < self.sheet.title_rows or col < self.sheet.title_cols:
                    attr |= curses.A_BOLD
                self.stdscr.addnstr(y, x, text, render_width, attr)
                column_index = spill_to_index + 1

        self._draw_cell_borders(top_grid_row, display_lines, visible_columns)
        self.stdscr.addnstr(height - 1, 0, self._status_line(width), width - 1, self._bar_attr(bold=True))
        self._draw_settings_cog(height, width)
        self._draw_key_overlay(height, width)
        self.stdscr.refresh()

    def _settings_label(self) -> str:
        return "[⚙]"

    def _draw_settings_cog(self, height: int, width: int) -> None:
        label = self._settings_label()
        x = max(0, width - len(label) - 1)
        attr = self._bar_attr(bold=True)
        self.stdscr.addnstr(height - 1, x, label, len(label), attr)

    def _status_line(self, width: int) -> str:
        text = self.message
        if self.command_hint_visible:
            hint = "Press / to start"
            if not text:
                text = hint
            elif hint not in text:
                text = f"{text}  |  {hint}"
        if self.selection_range:
            stats = self._selection_stats()
            if stats:
                text = f"{text}  |  {stats}"
        return text[: width - 1].ljust(width - 1)

    def _selection_stats(self) -> str:
        if not self.selection_range:
            return ""
        row_lo, col_lo, row_hi, col_hi = self.selection_range
        rows = row_hi - row_lo + 1
        cols = col_hi - col_lo + 1
        total = rows * cols
        count = 0
        total_sum = 0.0
        for row in range(row_lo, row_hi + 1):
            for col in range(col_lo, col_hi + 1):
                raw = self.sheet.get_raw(row, col)
                if not raw:
                    continue
                if is_formula_text(raw):
                    try:
                        value = self.evaluator.evaluate_cell(row, col, set())
                    except FormulaError:
                        continue
                else:
                    value = raw
                try:
                    number = float(str(value).replace(",", ""))
                except (ValueError, TypeError):
                    continue
                count += 1
                total_sum += number
        if count:
            avg = total_sum / count
            return f"sel {rows}x{cols}={total}  count {count}  sum {total_sum:g}  avg {avg:g}"
        return f"sel {rows}x{cols}={total}  count 0  sum -  avg -"

    def _draw_key_overlay(self, height: int, width: int) -> None:
        if not self.key_overlay_visible:
            return
        lines = get_key_help_lines(self.sheet.language)
        if not lines:
            return
        overlay_lines = lines[: min(8, len(lines))]
        panel_width = min(width - 4, max(len(line) for line in overlay_lines) + 4) if width > 6 else width - 2
        if panel_width <= 0:
            return
        panel_height = len(overlay_lines) + 2
        start_y = max(0, height - panel_height - 3)
        start_x = max(0, (width - panel_width) // 2)
        attr = self._help_attr()
        for row in range(panel_height):
            y = start_y + row
            if y < 0 or y >= height - 1:
                continue
            try:
                self.stdscr.addnstr(y, start_x, " " * panel_width, panel_width, attr)
            except curses.error:
                continue


    def _draw_column_headers(self, y: int, visible_columns: list[tuple[int, int, int]]) -> None:
        for col, x, col_width in visible_columns:
            label = column_label(col)
            if col < self.sheet.title_cols:
                label = f"*{label}"
            if self.sheet.is_col_hidden(col):
                label = f"·{label}"
            attr = curses.A_BOLD
            if self.colors_ready:
                attr |= curses.color_pair(COLOR_PAIR_HEADER)
            self.stdscr.addnstr(y, x, label.center(col_width - 1), col_width - 1, attr)

    def _draw_grid(
        self,
        header_y: int,
        grid_height: int,
        row_header_width: int,
        visible_columns: list[tuple[int, int, int]],
    ) -> None:
        if not visible_columns:
            return
        verticals = [row_header_width] + [x + width - 1 for _col, x, width in visible_columns]
        bottom_y = header_y + grid_height
        for y in range(header_y, bottom_y + 1):
            for x in verticals:
                self.stdscr.addch(y, x, curses.ACS_VLINE, self._grid_attr())

    def _draw_cell_borders(
        self,
        top_grid_row: int,
        display_lines: list[tuple[str, int, str | None]],
        visible_columns: list[tuple[int, int, int]],
    ) -> None:
        if not display_lines or not visible_columns:
            return
        attr = self._border_attr()
        for screen_offset, (kind, row, edge) in enumerate(display_lines):
            y = top_grid_row + screen_offset
            if kind == "row":
                row_border = row
                for col, x, col_width in visible_columns:
                    border = self.sheet.get_border(row_border, col)
                    if border not in {"all", "outline"}:
                        continue
                    left_x = x - 1
                    right_x = x + col_width - 1
                    try:
                        if left_x >= 0:
                            self.stdscr.addch(y, left_x, curses.ACS_VLINE, attr)
                        self.stdscr.addch(y, right_x, curses.ACS_VLINE, attr)
                    except curses.error:
                        pass
                continue
            if kind != "sep" or edge is None:
                continue
            row_border = row
            for col, x, col_width in visible_columns:
                border = self.sheet.get_border(row_border, col)
                if border not in {"all", "outline"}:
                    continue
                above_same = row_border > 0 and self.sheet.get_border(row_border - 1, col) in {"all", "outline"}
                below_same = row_border < self.sheet.rows - 1 and self.sheet.get_border(row_border + 1, col) in {"all", "outline"}
                if edge == "top" and above_same:
                    continue
                if edge == "bottom" and below_same:
                    continue
                left_x = x - 1
                right_x = x + col_width - 1
                try:
                    if left_x >= 0:
                        self.stdscr.addch(y, left_x, curses.ACS_HLINE, attr)
                    for hx in range(x, right_x):
                        self.stdscr.addch(y, hx, curses.ACS_HLINE, attr)
                    self.stdscr.addch(y, right_x, curses.ACS_HLINE, attr)
                except curses.error:
                    pass

    def handle_key(self, key: int) -> None:
        if key == curses.KEY_F1:
            self.key_overlay_visible = not self.key_overlay_visible
            return
        if key in (curses.KEY_UP, ord("k")):
            self.move(-1, 0)
        elif key in (curses.KEY_DOWN, ord("j")):
            self.move(1, 0)
        elif key in (curses.KEY_LEFT, ord("h")):
            self.move(0, -1)
        elif key in (curses.KEY_RIGHT, ord("l")):
            self.move(0, 1)
        elif key in (getattr(curses, "KEY_SR", -1),):
            self.extend_selection(-1, 0)
        elif key in (getattr(curses, "KEY_SF", -1),):
            self.extend_selection(1, 0)
        elif key in (getattr(curses, "KEY_SLEFT", -1),):
            self.extend_selection(0, -1)
        elif key in (getattr(curses, "KEY_SRIGHT", -1),):
            self.extend_selection(0, 1)
        elif key == getattr(curses, "KEY_SPREVIOUS", -1):
            self.switch_tab(-1)
        elif key == getattr(curses, "KEY_SNEXT", -1):
            self.switch_tab(1)
        elif key in (10, 13):
            self.edit_current_cell()
        elif key == 3:
            self.copy_selection_to_clipboard()
        elif key in (22, 25):
            self.paste_clipboard()
        elif key in (5, curses.KEY_F2):
            self.edit_formula_bar()
        elif key == curses.KEY_DC:
            self.clear_current_cell()
        elif key == 0:
            self._select_column(self.current_col)
        elif key == curses.KEY_MOUSE:
            self._handle_mouse()
        elif key == ord("/"):
            self.command_hint_visible = False
            self.run_command_prompt()
        elif key == 27:
            if self.raw_sheet_view:
                self.raw_sheet_view = False
                self.message = "Raw view off."
                return
            if not self._handle_escape_sequence():
                self.message = "Ready."
        elif key == 19:
            self._execute_file_command("save", [])
        elif key == 2:
            self._apply_style_shortcut("bold")
        elif key == 21:
            self._apply_style_shortcut("underline")
        elif key == 9:
            self._apply_style_shortcut("italic")
        elif key == 17:
            self.execute_command("quit", [])
        elif key == 26:
            self.undo_last_action()
        elif key == 18:
            self.redo_last_action()
        elif key == curses.KEY_NPAGE:
            self.move(10, 0)
        elif key == curses.KEY_PPAGE:
            self.move(-10, 0)
        elif key == ord(" "):
            self.edit_current_cell(initial_text=" ", replace=True)
        elif 32 <= key <= 126:
            self.edit_current_cell(initial_text=chr(key), replace=True)

    def move(self, row_delta: int, col_delta: int) -> None:
        self.selection_anchor = None
        self.selection_range = None
        self.mouse_dragging = False
        self.current_row = self._step_visible_row(self.current_row, row_delta)
        self.current_col = self._step_visible_col(self.current_col, col_delta)
        self._scroll_into_view()

    def _apply_style_shortcut(self, style: str) -> None:
        row_lo, col_lo, row_hi, col_hi = self._target_range(None)
        self._save_undo_state()
        for row in range(row_lo, row_hi + 1):
            for col in range(col_lo, col_hi + 1):
                if self.sheet.is_protected(row, col):
                    continue
                enabled = not self.sheet.has_text_style(row, col, style)
                self.sheet.set_text_style(row, col, style, enabled=enabled)
        self.dirty = True
        self.message = f"Style {style} toggled on {self._range_label(row_lo, col_lo, row_hi, col_hi)}"
        self.message = f"Cell {column_label(self.current_col)}{self.current_row + 1}"

    def extend_selection(self, row_delta: int, col_delta: int) -> None:
        if self.selection_anchor is None:
            self.selection_anchor = (self.current_row, self.current_col)
        self.mouse_dragging = False
        self.current_row = self._step_visible_row(self.current_row, row_delta)
        self.current_col = self._step_visible_col(self.current_col, col_delta)
        self.selection_range = self._normalize_range(self.selection_anchor, (self.current_row, self.current_col))
        self._scroll_into_view()
        self.message = f"Selected {self._selection_label()}"

    def _capture_tab_state(self, name: str | None = None) -> TabState:
        return TabState(
            sheet=self.sheet,
            path=self.path,
            name=name,
            dirty=self.dirty,
            current_row=self.current_row,
            current_col=self.current_col,
            row_offset=self.row_offset,
            col_offset=self.col_offset,
            selection_anchor=self.selection_anchor,
            selection_range=self.selection_range,
            mouse_dragging=self.mouse_dragging,
            undo_stack=list(self.undo_stack),
            redo_stack=list(self.redo_stack),
        )

    def _store_current_tab_state(self) -> None:
        self._remember_cursor_position(self.path, self.current_row, self.current_col)
        if not self.tabs:
            self.tabs.append(self._capture_tab_state())
            self.current_tab_index = 0
            return
        current_name = self.tabs[self.current_tab_index].name
        self.tabs[self.current_tab_index] = self._capture_tab_state(name=current_name)

    def _restore_tab_state(self, tab: TabState) -> None:
        self.sheet = tab.sheet
        self.evaluator = Evaluator(self.sheet)
        self.path = tab.path
        self.dirty = tab.dirty
        saved_row, saved_col = self._saved_cursor_position(tab.path)
        self.current_row = min(self.sheet.rows - 1, saved_row if saved_row or saved_col else tab.current_row)
        self.current_col = min(self.sheet.cols - 1, saved_col if saved_row or saved_col else tab.current_col)
        self.row_offset = tab.row_offset
        self.col_offset = tab.col_offset
        self.selection_anchor = tab.selection_anchor
        self.selection_range = tab.selection_range
        self.mouse_dragging = tab.mouse_dragging
        self.undo_stack = list(tab.undo_stack)
        self.redo_stack = list(tab.redo_stack)
        self._refresh_theme_colors()

    def _tab_label(self, index: int, tab: TabState) -> str:
        base = tab.name or (tab.path.name if tab.path else f"untitled-{index + 1}")
        if tab.dirty:
            base += "*"
        return base

    def _tabs_line(self, width: int) -> str:
        chips = []
        for index, tab in enumerate(self.tabs):
            label = self._tab_label(index, tab)
            if index == self.current_tab_index:
                chips.append(f"[{label}]")
            else:
                chips.append(label)
        text = " " + " ".join(chips) + " "
        return text[: width - 1].ljust(width - 1)

    def _switch_to_tab(self, index: int) -> None:
        if not self.tabs or index == self.current_tab_index or not (0 <= index < len(self.tabs)):
            return
        self._store_current_tab_state()
        self.current_tab_index = index
        self._restore_tab_state(self.tabs[index])
        self._scroll_into_view()
        self.message = f"Tab {index + 1}/{len(self.tabs)}: {self._tab_label(index, self.tabs[index])}"

    def switch_tab(self, delta: int) -> None:
        if len(self.tabs) <= 1:
            self.message = "Only one tab open."
            return
        self._switch_to_tab((self.current_tab_index + delta) % len(self.tabs))

    def _add_loaded_tab(self, target: Path, switch: bool = True) -> None:
        loaded_sheet = load_sheet(target, defaults=self._settings_payload())
        saved_row, saved_col = self._saved_cursor_position(target)
        new_tab = TabState(
            sheet=loaded_sheet,
            path=target,
            name=None,
            dirty=False,
            current_row=min(loaded_sheet.rows - 1, saved_row),
            current_col=min(loaded_sheet.cols - 1, saved_col),
            row_offset=0,
            col_offset=0,
            selection_anchor=None,
            selection_range=None,
            mouse_dragging=False,
            undo_stack=[],
            redo_stack=[],
        )
        if len(self.tabs) == 1 and self.path is None and not self.dirty and not any(raw for _r, _c, raw in self.sheet.iter_cells()):
            self.tabs[0] = new_tab
            self.current_tab_index = 0
            self._restore_tab_state(new_tab)
            self._scroll_into_view()
        else:
            self._store_current_tab_state()
            self.tabs.append(new_tab)
            if switch:
                self.current_tab_index = len(self.tabs) - 1
                self._restore_tab_state(new_tab)
                self._scroll_into_view()
        self._remember_recent_file(target)
        self.message = f"Loaded {target} in tab {self.current_tab_index + 1}"

    def _close_current_tab(self) -> None:
        if not self.tabs:
            self.running = False
            self.message = self._tr("bye")
            return
        closing_index = self.current_tab_index
        closing_label = self._tab_label(closing_index, self.tabs[closing_index])
        if len(self.tabs) == 1:
            self.running = False
            self.message = self._tr("bye")
            return
        self.tabs.pop(closing_index)
        new_index = min(closing_index, len(self.tabs) - 1)
        self.current_tab_index = new_index
        self._restore_tab_state(self.tabs[new_index])
        self._scroll_into_view()
        self.message = f"Closed {closing_label}"

    def _duplicate_current_tab(self) -> None:
        if not self.tabs:
            self.message = "No tabs to duplicate."
            return
        self._store_current_tab_state()
        current = self.tabs[self.current_tab_index]
        cloned_sheet = Spreadsheet.from_dict(current.sheet.to_dict())
        cloned_name = f"{self._tab_label(self.current_tab_index, current)} copy"
        new_tab = TabState(
            sheet=cloned_sheet,
            path=None,
            name=cloned_name,
            dirty=True,
            current_row=current.current_row,
            current_col=current.current_col,
            row_offset=current.row_offset,
            col_offset=current.col_offset,
            selection_anchor=None,
            selection_range=None,
            mouse_dragging=False,
            undo_stack=[],
            redo_stack=[],
        )
        insert_index = self.current_tab_index + 1
        self.tabs.insert(insert_index, new_tab)
        self.current_tab_index = insert_index
        self._restore_tab_state(new_tab)
        self._scroll_into_view()
        self.message = f"Duplicated tab to {self._tab_label(self.current_tab_index, new_tab)}"

    def _move_tab(self, delta: int) -> None:
        if len(self.tabs) <= 1:
            self.message = "Only one tab open."
            return
        target_index = (self.current_tab_index + delta) % len(self.tabs)
        if target_index == self.current_tab_index:
            return
        self.tabs[self.current_tab_index], self.tabs[target_index] = (
            self.tabs[target_index],
            self.tabs[self.current_tab_index],
        )
        self.current_tab_index = target_index
        self._restore_tab_state(self.tabs[self.current_tab_index])
        self._scroll_into_view()
        self.message = f"Tab moved to {self.current_tab_index + 1}/{len(self.tabs)}"

    def _handle_mouse(self) -> None:
        try:
            _id, mouse_x, mouse_y, _z, state = curses.getmouse()
        except curses.error:
            return
        wheel_up = getattr(curses, "BUTTON4_PRESSED", 0) | getattr(curses, "BUTTON4_CLICKED", 0)
        wheel_down = getattr(curses, "BUTTON5_PRESSED", 0) | getattr(curses, "BUTTON5_CLICKED", 0)
        if state & wheel_up:
            self._scroll_rows(-3)
            return
        if state & wheel_down:
            self._scroll_rows(3)
            return
        if state & (curses.BUTTON1_PRESSED | curses.BUTTON1_CLICKED):
            if self._handle_settings_click(mouse_y, mouse_x):
                return
            if self._handle_tab_click(mouse_y, mouse_x):
                return
            if self._handle_header_click(mouse_y, mouse_x):
                return
        target = self._cell_from_screen(mouse_y, mouse_x)
        if target is None:
            return
        row, col = target
        self.current_row = row
        self.current_col = col
        self._scroll_into_view()
        report_motion = getattr(curses, "REPORT_MOUSE_POSITION", 0)
        release_mask = getattr(curses, "BUTTON1_RELEASED", 0)
        if state & curses.BUTTON1_PRESSED:
            self.selection_anchor = target
            self.selection_range = (row, col, row, col)
            self.mouse_dragging = True
            self.message = f"Selected {self._selection_label()}"
            return
        if state & report_motion and self.mouse_dragging and self.selection_anchor:
            self.selection_range = self._normalize_range(self.selection_anchor, target)
            self.message = f"Selected {self._selection_label()}"
            return
        if state & release_mask and self.mouse_dragging and self.selection_anchor:
            self.selection_range = self._normalize_range(self.selection_anchor, target)
            self.mouse_dragging = False
            self.message = f"Selected {self._selection_label()}"
            return
        if state & curses.BUTTON1_CLICKED:
            self.selection_anchor = target
            self.selection_range = (row, col, row, col)
            self.mouse_dragging = False
            self.message = f"Selected {self._selection_label()}"
            return
        self.message = f"Cell {column_label(col)}{row + 1}"

    def _scroll_rows(self, delta: int) -> None:
        height, width = self.stdscr.getmaxyx()
        top_grid_row, grid_height, _row_header_width, _visible_columns = self._grid_layout(height, width)
        visible_rows = max(1, len(self._visible_rows(grid_height)))
        max_offset = max(0, self.sheet.rows - visible_rows)
        self.row_offset = max(0, min(max_offset, self.row_offset + delta))
        self.message = f"Scrolled to row {self.row_offset + 1}"

    def _handle_settings_click(self, y: int, x: int) -> bool:
        height, width = self.stdscr.getmaxyx()
        label = self._settings_label()
        start_x = max(0, width - len(label) - 1)
        if y == height - 1 and start_x <= x < start_x + len(label):
            self.show_settings_screen()
            return True
        return False

    def _handle_tab_click(self, y: int, x: int) -> bool:
        if y != 0:
            return False
        position = 1
        for index, tab in enumerate(self.tabs):
            label = self._tab_label(index, tab)
            chip = f"[{label}]"
            end = position + len(chip)
            if position <= x < end:
                self._switch_to_tab(index)
                return True
            position = end + 1
        return False

    def _handle_header_click(self, y: int, x: int) -> bool:
        height, width = self.stdscr.getmaxyx()
        top_grid_row, grid_height, row_header_width, visible_columns = self._grid_layout(height, width)
        header_y = top_grid_row - 1
        if y == header_y:
            for col, col_x, col_width in visible_columns:
                if col_x <= x < col_x + col_width - 1:
                    if self._selection_is_full_col(col):
                        self.current_col = col
                        self._scroll_into_view()
                        self._show_column_header_menu(col)
                    else:
                        self._select_column(col)
                    return True
        if top_grid_row <= y < top_grid_row + grid_height and x < row_header_width:
            display_lines = self._display_lines(grid_height, visible_columns)
            row_line_index = y - top_grid_row
            if 0 <= row_line_index < len(display_lines) and display_lines[row_line_index][0] == "row":
                row = display_lines[row_line_index][1]
                if self._selection_is_full_row(row):
                    self.current_row = row
                    self._scroll_into_view()
                    self._show_row_header_menu(row)
                else:
                    self._select_row(row)
                return True
        return False

    def _selection_is_full_row(self, row: int) -> bool:
        return self.selection_range == (row, 0, row, self.sheet.cols - 1)

    def _selection_is_full_col(self, col: int) -> bool:
        return self.selection_range == (0, col, self.sheet.rows - 1, col)

    def _select_row(self, row: int) -> None:
        self.current_row = row
        self.selection_anchor = (row, 0)
        self.selection_range = (row, 0, row, self.sheet.cols - 1)
        self.mouse_dragging = False
        self._scroll_into_view()
        self.message = f"Selected row {row + 1}"

    def _select_column(self, col: int) -> None:
        self.current_col = col
        self.selection_anchor = (0, col)
        self.selection_range = (0, col, self.sheet.rows - 1, col)
        self.mouse_dragging = False
        self._scroll_into_view()
        self.message = f"Selected column {column_label(col)}"

    def edit_current_cell(self, initial_text: str = "", replace: bool = False) -> None:
        origin_row = self.current_row
        origin_col = self.current_col
        if self.sheet.is_protected(origin_row, origin_col):
            self.message = "Cell is protected."
            return
        raw = "" if replace else self.sheet.get_raw(origin_row, origin_col)
        initial_value = initial_text if replace else (initial_text or raw)
        # Re-read formula cells at prompt-open time so copy/paste mutations always
        # reopen from the latest stored formula instead of any stale prompt state.
        if not replace and is_formula_text(raw):
            initial_value = self.sheet.get_raw(origin_row, origin_col)
        edited = self.prompt(
            f"Edit {column_label(origin_col)}{origin_row + 1}: ",
            initial_value,
            formula_origin=(origin_row, origin_col) if is_formula_text(initial_value) else None,
        )
        self.current_row = origin_row
        self.current_col = origin_col
        if edited is None:
            self.message = "Edit cancelled."
            return
        edited = self._normalize_cell_input(origin_row, origin_col, edited)
        stored_ref = f"{column_label(origin_col)}{origin_row + 1}"
        self._save_undo_state()
        self.sheet.set_raw(origin_row, origin_col, edited)
        self._apply_default_alignment(origin_row, origin_col, edited)
        self.dirty = True
        if origin_row >= self.sheet.rows - 1:
            self.sheet.ensure_size(origin_row + 1, origin_col)
        self.current_row = min(self.sheet.rows - 1, origin_row + 1)
        self.current_col = origin_col
        self._scroll_into_view()
        self.message = f"Stored {stored_ref}; ready for {column_label(self.current_col)}{self.current_row + 1}"

    def edit_formula_bar(self) -> None:
        raw = self.sheet.get_raw(self.current_row, self.current_col)
        if self.sheet.is_protected(self.current_row, self.current_col):
            self.message = "Cell is protected."
            return
        edited = self.prompt(
            f"Formula {column_label(self.current_col)}{self.current_row + 1}: ",
            raw,
            formula_origin=(self.current_row, self.current_col) if is_formula_text(raw) else None,
        )
        if edited is None:
            self.message = "Edit cancelled."
            return
        edited = self._normalize_cell_input(self.current_row, self.current_col, edited)
        self._save_undo_state()
        self.sheet.set_raw(self.current_row, self.current_col, edited)
        self._apply_default_alignment(self.current_row, self.current_col, edited)
        self.dirty = True
        self.message = f"Stored {column_label(self.current_col)}{self.current_row + 1}"

    def _show_column_header_menu(self, col: int) -> None:
        choice = self._choose_from_menu(
            f"Column {column_label(col)}",
            ["duplicate", "freeze", "hide", "insert before", "unhide all", "delete", "width"],
            default_option="freeze",
        )
        if choice is None:
            self.message = "Column action cancelled."
            return
        if choice == "freeze":
            self._save_undo_state()
            self.sheet.title_cols = col + 1
            self.dirty = True
            self.message = f"Frozen columns through {column_label(col)}"
            return
        if choice == "duplicate":
            self._save_undo_state()
            self._duplicate_columns(col, col)
            self.dirty = True
            self.message = f"Duplicated column {column_label(col)}"
            return
        if choice == "hide":
            self._save_undo_state()
            self.sheet.hide_col(col)
            self.current_col = self._first_visible_col()
            self._scroll_into_view()
            self.dirty = True
            self.message = f"Hidden column {column_label(col)}"
            return
        if choice == "unhide all":
            self._save_undo_state()
            self.sheet.hidden_cols.clear()
            self.dirty = True
            self.message = "Unhid all columns."
            return
        if choice == "insert before":
            self._save_undo_state()
            self._rebuild_cols(col, 1)
            self.current_col = min(self.sheet.cols - 1, col)
            self.dirty = True
            self.message = f"Inserted column at {column_label(col)}"
            return
        if choice == "delete":
            self._save_undo_state()
            self._rebuild_cols(col, -1)
            self.current_col = min(self.sheet.cols - 1, col)
            self.dirty = True
            self.message = f"Deleted column {column_label(col)}"
            return
        if choice == "width":
            current_width = str(self.sheet.get_column_width(col))
            width_text = self.prompt(f"Width for {column_label(col)}: ", current_width)
            if width_text is None or not width_text.strip():
                self.message = "Column width cancelled."
                return
            try:
                width = max(8, int(width_text.strip()))
            except ValueError:
                self.message = "Column width must be a number."
                return
            self._save_undo_state()
            self.sheet.set_column_width(col, width)
            self.dirty = True
            self.message = f"Width set to {width} for {column_label(col)}"

    def _show_row_header_menu(self, row: int) -> None:
        choice = self._choose_from_menu(
            f"Row {row + 1}",
            ["duplicate", "freeze", "hide", "insert above", "unhide all", "delete"],
            default_option="freeze",
        )
        if choice is None:
            self.message = "Row action cancelled."
            return
        if choice == "freeze":
            self._save_undo_state()
            self.sheet.title_rows = row + 1
            self.dirty = True
            self.message = f"Frozen rows through {row + 1}"
            return
        if choice == "duplicate":
            self._save_undo_state()
            self._duplicate_rows(row, row)
            self.dirty = True
            self.message = f"Duplicated row {row + 1}"
            return
        if choice == "hide":
            self._save_undo_state()
            self.sheet.hide_row(row)
            self.current_row = self._first_visible_row()
            self._scroll_into_view()
            self.dirty = True
            self.message = f"Hidden row {row + 1}"
            return
        if choice == "unhide all":
            self._save_undo_state()
            self.sheet.hidden_rows.clear()
            self.dirty = True
            self.message = "Unhid all rows."
            return
        if choice == "insert above":
            self._save_undo_state()
            self._rebuild_rows(row, 1)
            self.current_row = min(self.sheet.rows - 1, row)
            self.dirty = True
            self.message = f"Inserted row at {row + 1}"
            return
        if choice == "delete":
            self._save_undo_state()
            self._rebuild_rows(row, -1)
            self.current_row = min(self.sheet.rows - 1, row)
            self.dirty = True
            self.message = f"Deleted row {row + 1}"

    def _normalize_cell_input(self, row: int, col: int, value: str) -> str:
        if not value or value.startswith("'") or is_formula_text(value):
            return value
        style = self.sheet.get_format(row, col)
        if style.startswith("date"):
            return normalize_date_text(value, style)
        try:
            return normalize_date_text(value, self.sheet.date_format)
        except ValueError:
            return value

    def clear_current_cell(self) -> None:
        if self.selection_range is not None:
            row_lo, col_lo, row_hi, col_hi = self.selection_range
        else:
            row_lo = row_hi = self.current_row
            col_lo = col_hi = self.current_col
        editable_cells = [
            (row, col)
            for row in range(row_lo, row_hi + 1)
            for col in range(col_lo, col_hi + 1)
            if not self.sheet.is_protected(row, col)
        ]
        if not editable_cells:
            self.message = "Cell is protected."
            return
        self._save_undo_state()
        for row, col in editable_cells:
            self.sheet.clear(row, col)
            if not self.sheet.is_alignment_manual(row, col):
                self.sheet.set_alignment(row, col, "", manual=False)
        self.dirty = True
        self.message = f"Cleared {self._range_label(row_lo, col_lo, row_hi, col_hi)}"

    def _clear_cell_full(self, row: int, col: int) -> None:
        self.sheet.clear(row, col)
        self.sheet.set_format(row, col, "")
        self.sheet.clear_text_styles(row, col)
        self.sheet.set_background(row, col, "")
        self.sheet.set_border(row, col, "")
        self.sheet.set_alignment(row, col, "", manual=False)
        self.sheet.set_font_size(row, col, 0)
        self.sheet.unprotect(row, col)

    def _cell_state(self, row: int, col: int) -> dict[str, object]:
        return {
            "raw": self.sheet.get_raw(row, col),
            "format": self.sheet.get_format(row, col),
            "text_styles": self.sheet.get_text_styles(row, col),
            "background": self.sheet.get_background(row, col),
            "border": self.sheet.get_border(row, col),
            "alignment": self.sheet.get_alignment(row, col),
            "alignment_manual": self.sheet.is_alignment_manual(row, col),
            "protected": self.sheet.is_protected(row, col),
            "font_size": self.sheet.font_sizes.get(self.sheet.key(row, col), 0),
        }

    def _apply_cell_state(self, row: int, col: int, state: dict[str, object]) -> None:
        self._clear_cell_full(row, col)
        raw = str(state.get("raw", ""))
        if raw:
            self.sheet.set_raw(row, col, raw)
        style = str(state.get("format", ""))
        if style:
            self.sheet.set_format(row, col, style)
        for text_style in state.get("text_styles", set()):
            self.sheet.set_text_style(row, col, str(text_style), enabled=True)
        background = str(state.get("background", ""))
        if background:
            self.sheet.set_background(row, col, background)
        border = str(state.get("border", ""))
        if border:
            self.sheet.set_border(row, col, border)
        align = str(state.get("alignment", ""))
        if align:
            self.sheet.set_alignment(row, col, align, manual=bool(state.get("alignment_manual", False)))
        size = int(state.get("font_size", 0) or 0)
        if size > 0:
            self.sheet.set_font_size(row, col, size)
        if state.get("protected", False):
            self.sheet.protect(row, col)

    def _delete_cells(self, row_lo: int, col_lo: int, row_hi: int, col_hi: int, mode: str) -> None:
        rows = self.sheet.rows
        cols = self.sheet.cols
        count_rows = row_hi - row_lo + 1
        count_cols = col_hi - col_lo + 1
        if mode == "clear":
            for row in range(row_lo, row_hi + 1):
                for col in range(col_lo, col_hi + 1):
                    self._clear_cell_full(row, col)
            return
        if mode == "up":
            for col in range(col_lo, col_hi + 1):
                for row in range(row_lo, rows - count_rows):
                    state = self._cell_state(row + count_rows, col)
                    self._apply_cell_state(row, col, state)
                for row in range(rows - count_rows, rows):
                    self._clear_cell_full(row, col)
            return
        if mode == "left":
            for row in range(row_lo, row_hi + 1):
                for col in range(col_lo, cols - count_cols):
                    state = self._cell_state(row, col + count_cols)
                    self._apply_cell_state(row, col, state)
                for col in range(cols - count_cols, cols):
                    self._clear_cell_full(row, col)
            return

    def copy_selection_to_clipboard(self) -> None:
        if self.selection_range is not None:
            row_lo, col_lo, row_hi, col_hi = self.selection_range
            # Plain Ctrl+C should behave like classic spreadsheet copy:
            # copy the current cell unless there is a real multi-cell selection.
            if row_lo == row_hi and col_lo == col_hi:
                row_lo = row_hi = self.current_row
                col_lo = col_hi = self.current_col
        else:
            row_lo = row_hi = self.current_row
            col_lo = col_hi = self.current_col
        self._load_internal_clipboard_from_range(row_lo, col_lo, row_hi, col_hi)
        self._export_clipboard_to_terminal()
        self._export_clipboard_to_system()
        self.message = f"Copied {self._range_label(row_lo, col_lo, row_hi, col_hi)}"

    def _load_internal_clipboard_from_range(self, row_lo: int, col_lo: int, row_hi: int, col_hi: int) -> None:
        cells: list[tuple[int, int, str, str, str, str, str, str, bool, bool]] = []
        for row in range(row_lo, row_hi + 1):
            for col in range(col_lo, col_hi + 1):
                cells.append(
                    (
                        row - row_lo,
                        col - col_lo,
                        self.sheet.get_raw(row, col),
                        self.sheet.get_format(row, col),
                        ",".join(sorted(self.sheet.get_text_styles(row, col))),
                        self.sheet.get_background(row, col),
                        self.sheet.get_border(row, col),
                        self.sheet.get_alignment(row, col),
                        self.sheet.is_alignment_manual(row, col),
                        self.sheet.is_protected(row, col),
                    )
                )
        self.clipboard_cells = cells
        self.clipboard_size = (row_hi - row_lo + 1, col_hi - col_lo + 1)
        self.clipboard_origin = (row_lo, col_lo)
        self.prefer_internal_clipboard = True

    def paste_clipboard(self) -> None:
        if self.prefer_internal_clipboard and self.clipboard_cells:
            self._paste_internal_clipboard()
            return
        if self._paste_from_system_clipboard():
            return
        if self.clipboard_cells:
            self._paste_internal_clipboard()
            return
        self.message = "Clipboard is empty."

    def _paste_internal_clipboard(self, destination: tuple[int, int, int, int] | None = None) -> None:
        self._save_undo_state()
        cells_by_offset = {
            (row_offset, col_offset): (raw, style, text_styles, background, border, align, align_manual, protected)
            for row_offset, col_offset, raw, style, text_styles, background, border, align, align_manual, protected in self.clipboard_cells
        }
        clip_height, clip_width = self.clipboard_size
        if destination is not None:
            start_row, start_col, end_row, end_col = destination
        elif self.selection_range is not None:
            start_row, start_col, end_row, end_col = self.selection_range
        else:
            start_row = self.current_row
            start_col = self.current_col
            end_row = start_row + clip_height - 1
            end_col = start_col + clip_width - 1
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                if self.sheet.is_protected(row, col):
                    continue
                row_offset = (row - start_row) % max(1, clip_height)
                col_offset = (col - start_col) % max(1, clip_width)
                raw, style, text_styles, background, border, align, align_manual, protected = cells_by_offset[(row_offset, col_offset)]
                src_row = self.clipboard_origin[0] + row_offset
                src_col = self.clipboard_origin[1] + col_offset
                shifted_raw = shift_formula_references(raw, row - src_row, col - src_col)
                self.sheet.set_raw(row, col, shifted_raw)
                self.sheet.set_format(row, col, style)
                self.sheet.clear_text_styles(row, col)
                for text_style in [item for item in text_styles.split(",") if item]:
                    self.sheet.set_text_style(row, col, text_style, enabled=True)
                self.sheet.set_background(row, col, background)
                self.sheet.set_border(row, col, border)
                self.sheet.set_alignment(row, col, align, manual=align_manual)
                if protected:
                    self.sheet.protect(row, col)
                else:
                    self.sheet.unprotect(row, col)
        self.dirty = True
        self.message = f"Pasted to {self._range_label(start_row, start_col, end_row, end_col)}"

    def _save_undo_state(self) -> None:
        state = {
            "sheet": self.sheet.to_dict(),
            "dirty": self.dirty,
            "path": None if self.path is None else str(self.path),
            "current_row": self.current_row,
            "current_col": self.current_col,
            "row_offset": self.row_offset,
            "col_offset": self.col_offset,
            "selection_range": self.selection_range,
            "selection_anchor": self.selection_anchor,
        }
        self.undo_stack.append(state)
        if len(self.undo_stack) > self.max_history:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def _capture_state(self) -> dict[str, object]:
        return {
            "sheet": self.sheet.to_dict(),
            "dirty": self.dirty,
            "path": None if self.path is None else str(self.path),
            "current_row": self.current_row,
            "current_col": self.current_col,
            "row_offset": self.row_offset,
            "col_offset": self.col_offset,
            "selection_range": self.selection_range,
            "selection_anchor": self.selection_anchor,
        }

    def _restore_state(self, state: dict[str, object]) -> None:
        self.sheet = Spreadsheet.from_dict(state["sheet"])
        self.evaluator = Evaluator(self.sheet)
        self._refresh_theme_colors()
        self.dirty = bool(state["dirty"])
        path_text = state["path"]
        self.path = None if path_text is None else Path(str(path_text))
        self.current_row = int(state["current_row"])
        self.current_col = int(state["current_col"])
        self.row_offset = int(state["row_offset"])
        self.col_offset = int(state["col_offset"])
        self.selection_range = state["selection_range"]
        self.selection_anchor = state["selection_anchor"]
        self.mouse_dragging = False

    def undo_last_action(self) -> None:
        if not self.undo_stack:
            self.message = "Nothing to undo."
            return
        self.redo_stack.append(self._capture_state())
        state = self.undo_stack.pop()
        self._restore_state(state)
        self.message = "Undid last action."

    def redo_last_action(self) -> None:
        if not self.redo_stack:
            self.message = "Nothing to redo."
            return
        self.undo_stack.append(self._capture_state())
        state = self.redo_stack.pop()
        self._restore_state(state)
        self.message = "Redid last action."

    def _clipboard_payload(self) -> str:
        payload = {
            "origin": list(self.clipboard_origin),
            "size": list(self.clipboard_size),
            "cells": [list(cell) for cell in self.clipboard_cells],
        }
        encoded = base64.b64encode(json.dumps(payload).encode("utf-8")).decode("ascii")
        return f"{CLIPBOARD_MARKER}{encoded}"

    def _clipboard_plain_text(self) -> str:
        if not self.clipboard_cells:
            return ""
        row_lo, col_lo = self.clipboard_origin
        clip_height, clip_width = self.clipboard_size
        lines: list[str] = []
        for row in range(row_lo, row_lo + clip_height):
            row_values = [self.sheet.get_raw(row, col) for col in range(col_lo, col_lo + clip_width)]
            lines.append("\t".join(row_values))
        return "\n".join(lines)

    def _export_clipboard_to_terminal(self) -> None:
        if not self.clipboard_cells:
            return
        encoded = base64.b64encode(self._clipboard_payload().encode("utf-8")).decode("ascii")
        sys.stdout.write(f"\x1b]52;c;{encoded}\x07")
        sys.stdout.flush()

    def _export_clipboard_to_system(self) -> None:
        if not self.clipboard_cells:
            return
        text = self._clipboard_plain_text()
        try:
            subprocess.run(
                ["/usr/bin/wl-copy"],
                input=text,
                text=True,
                check=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        except (FileNotFoundError, subprocess.SubprocessError):
            return

    def _paste_from_system_clipboard(self) -> bool:
        if self.prefer_internal_clipboard and self.clipboard_cells:
            self._paste_internal_clipboard()
            return True
        try:
            result = subprocess.run(
                ["/usr/bin/wl-paste", "-n"],
                text=True,
                capture_output=True,
                check=True,
            )
        except (FileNotFoundError, subprocess.SubprocessError):
            return False
        text = result.stdout
        if not text:
            return False
        if self.clipboard_cells and text.rstrip("\n") == self._clipboard_plain_text().rstrip("\n"):
            self._paste_internal_clipboard()
            return True
        if self._load_clipboard_payload(text):
            return True
        rows = list(csv.reader(text.splitlines(), delimiter="\t"))
        if not rows:
            return False
        start_row, start_col, end_row, end_col = self._target_range(None)
        if len(rows) == 1 and len(rows[0]) == 1 and self.selection_range is None:
            if self.sheet.is_protected(start_row, start_col):
                self.message = "Cell is protected."
                return True
            self._save_undo_state()
            value = self._normalize_cell_input(start_row, start_col, rows[0][0])
            self.sheet.set_raw(start_row, start_col, value)
            self._apply_default_alignment(start_row, start_col, value)
            self.prefer_internal_clipboard = False
            self.dirty = True
            self.message = f"Pasted to {self._range_label(start_row, start_col, start_row, start_col)}"
            return True
        self._save_undo_state()
        row_count = max(1, len(rows))
        col_count = max(1, max((len(row) for row in rows), default=1))
        if self.selection_range is None:
            end_row = start_row + row_count - 1
            end_col = start_col + col_count - 1
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                if self.sheet.is_protected(row, col):
                    continue
                row_offset = (row - start_row) % row_count
                source_row = rows[row_offset] if row_offset < len(rows) else []
                col_offset = (col - start_col) % col_count
                value = source_row[col_offset] if col_offset < len(source_row) else ""
                self.sheet.set_raw(row, col, value)
        self.prefer_internal_clipboard = False
        self.dirty = True
        self.message = f"Pasted to {self._range_label(start_row, start_col, end_row, end_col)}"
        return True

    def _load_clipboard_payload(self, text: str) -> bool:
        if not text.startswith(CLIPBOARD_MARKER):
            return False
        try:
            payload = json.loads(base64.b64decode(text[len(CLIPBOARD_MARKER) :]).decode("utf-8"))
        except (ValueError, json.JSONDecodeError):
            self.message = "Clipboard data is invalid."
            return True
        origin = payload.get("origin", [0, 0])
        size = payload.get("size", [0, 0])
        cells = payload.get("cells", [])
        self.clipboard_origin = (int(origin[0]), int(origin[1]))
        self.clipboard_size = (max(0, int(size[0])), max(0, int(size[1])))
        self.clipboard_cells = [
            (
                int(row_offset),
                int(col_offset),
                str(raw),
                str(style),
                str(text_styles),
                str(background),
                str(border),
                str(align),
                bool(manual),
                bool(protected),
            )
            for row_offset, col_offset, raw, style, text_styles, background, border, align, manual, protected in cells
        ]
        self.prefer_internal_clipboard = True
        self.paste_clipboard()
        return True

    def run_command_prompt(self) -> None:
        advanced = False
        default_option = "edit"
        while True:
            options = ADVANCED_COMMAND_MENU_OPTIONS if advanced else COMMAND_MENU_OPTIONS
            selected = self._choose_from_menu(
                "//" if advanced else "/",
                options,
                default_option=default_option if default_option in options else options[0],
                descriptions=COMMAND_DESCRIPTIONS,
                toggle_key=ord("/"),
                toggle_value="__toggle__",
                footer_hint=" arrows/type/Enter/Esc  /=advanced ",
            )
            if selected is None:
                self.message = tr(self.sheet.language, "command_cancelled")
                return
            if selected == "__toggle__":
                advanced = not advanced
                default_option = "edit" if "edit" in options else options[0]
                continue
            self._launch_menu_command(selected)
            return

    def _tr(self, key: str) -> str:
        return tr(self.sheet.language, key)

    def _set_sheet_date_format(self, selected_date: str) -> None:
        self._save_undo_state()
        self.sheet.date_format = f"date:{selected_date}"
        for row, col, raw in list(self.sheet.iter_cells()):
            if not raw or raw.startswith("'") or is_formula_text(raw):
                continue
            try:
                self.sheet.set_raw(row, col, normalize_date_text(raw, self.sheet.date_format))
            except ValueError:
                continue
        self._save_global_settings()
        self.dirty = True
        self.message = f"Sheet date format set to {selected_date}"

    def _set_sheet_time_format(self, selected_time: str) -> None:
        self._save_undo_state()
        self.sheet.time_format = f"time:{selected_time}"
        for row, col, raw in list(self.sheet.iter_cells()):
            if not raw or raw.startswith("'") or is_formula_text(raw):
                continue
            try:
                self.sheet.set_raw(row, col, normalize_time_text(raw, self.sheet.time_format))
            except ValueError:
                continue
        self._save_global_settings()
        self.dirty = True
        self.message = f"Sheet time format set to {selected_time}"

    def _set_theme(self, theme_name: str) -> None:
        self._save_undo_state()
        self.sheet.theme_name = theme_name
        self._refresh_theme_colors()
        self._save_global_settings()
        self.dirty = True
        self.message = f"Theme set to {self.sheet.theme_name}"

    def _set_active_cell_color(self, color_name: str) -> None:
        self._save_undo_state()
        self.sheet.active_cell_color = color_name
        self._refresh_theme_colors()
        self._save_global_settings()
        self.dirty = True
        self.message = f"Active cell color set to {self.sheet.active_cell_color}"

    def _set_formula_coloration(self, enabled: bool) -> None:
        self._save_undo_state()
        self.sheet.formula_coloration = enabled
        self._save_global_settings()
        self.dirty = True
        self.message = f"Formula coloration {'on' if enabled else 'off'}"

    def _set_formula_foreground_color(self, color_name: str) -> None:
        self._save_undo_state()
        self.sheet.formula_foreground_color = color_name
        self._refresh_theme_colors()
        self._save_global_settings()
        self.dirty = True
        self.message = f"Formula color set to {self.sheet.formula_foreground_color}"

    def _set_formula_color_setting(self, value: str) -> None:
        self._save_undo_state()
        if value == "off":
            self.sheet.formula_coloration = False
        else:
            self.sheet.formula_coloration = True
            self.sheet.formula_foreground_color = value
        self._refresh_theme_colors()
        self._save_global_settings()
        self.dirty = True
        self.message = f"Formula color set to {value}"

    def _set_protected_colors(self, foreground_name: str | None = None, background_name: str | None = None) -> None:
        self._save_undo_state()
        if foreground_name is not None:
            self.sheet.protected_foreground_color = foreground_name
        if background_name is not None:
            self.sheet.protected_background_color = background_name
        self._refresh_theme_colors()
        self._save_global_settings()
        self.dirty = True
        self.message = (
            f"Protected colors: {self.sheet.protected_foreground_color}"
            f" on {self.sheet.protected_background_color}"
        )

    def _prompt_header_color(self, label: str, current: str, default_value: str) -> str | None:
        initial = "" if current == default_value else current.removeprefix("#")
        value = self.prompt(
            f"{label}: ",
            initial,
            help_lines=[
                " Enter 6 hex chars like FFCC99",
                f" Type none for default {default_value}",
            ],
        ).strip()
        if not value:
            return None
        if value.lower() == "none":
            return default_value
        if HEX_COLOR_RE.match(value):
            normalized = value.upper()
            if not normalized.startswith("#"):
                normalized = f"#{normalized}"
            return normalized
        raise ValueError("need 6 hex characters or none")

    def _set_language(self, language_name: str) -> None:
        self._save_undo_state()
        self.sheet.language = LANGUAGE_CODES.get(language_name, "en")
        self._save_global_settings()
        self.dirty = True
        self.message = language_name

    def show_settings_screen(self) -> None:
        selected_row = 0
        selected_col = 0
        curses.curs_set(0)
        while True:
            height, width = self.stdscr.getmaxyx()
            self.stdscr.erase()
            self.stdscr.addnstr(0, 0, f" {self._tr('settings')} ".ljust(width - 1), width - 1, self._bar_attr(bold=True))
            help_lines = [
                f" {self._tr('settings_help_1')}",
                f" {self._tr('settings_help_2')}",
                f" {self._tr('settings_help_3')}",
            ]
            for offset, line in enumerate(help_lines, start=2):
                if offset >= height - 1:
                    break
                self.stdscr.addnstr(offset, 0, line.ljust(width - 1), width - 1, self._help_attr())

            current_language = next((name for name, code in LANGUAGE_CODES.items() if code == self.sheet.language), "english")
            rows = [
                (self._tr("theme"), self.sheet.theme_name, None),
                (self._tr("date_format"), self.sheet.date_format.split(":", 1)[1], None),
                (self._tr("time_format"), self.sheet.time_format.split(":", 1)[1], None),
                (self._tr("active_cell"), self.sheet.active_cell_color, None),
                (self._tr("sheet_fg"), self.sheet.sheet_foreground_color, self.sheet.sheet_background_color),
                (self._tr("tui_fg"), self.sheet.tui_foreground_color, self.sheet.tui_background_color),
                (self._tr("row_header_fg"), self.sheet.row_header_foreground_color, self.sheet.row_header_background_color),
                (self._tr("column_header_fg"), self.sheet.column_header_foreground_color, self.sheet.column_header_background_color),
                (self._tr("formula_colour"), self.sheet.formula_foreground_color if self.sheet.formula_coloration else "off", None),
                (self._tr("protected_fg"), self.sheet.protected_foreground_color, self.sheet.protected_background_color),
                (self._tr("language"), current_language, None),
            ]
            if selected_row >= len(rows):
                selected_row = max(0, len(rows) - 1)
            if rows and rows[selected_row][2] is None and selected_col == 1:
                selected_col = 0

            label_width = min(24, max(len(label) for label, _value, _bg in rows) + 2)
            value_width = max(10, min(18, (width - label_width - 8) // 2)) if width > 40 else 10
            label_x = 2
            fg_x = label_x + label_width + 2
            bg_x = fg_x + value_width + 2
            first_row_y = 6
            row_gap = 1
            for index, (label, fg_value, bg_value) in enumerate(rows):
                y = first_row_y + (index * row_gap)
                if y >= height - 2:
                    break
                self.stdscr.addnstr(y, 0, (" " * (width - 1)), width - 1, self._help_attr())
                self.stdscr.addnstr(y, label_x, label.ljust(label_width), max(0, width - 3), self._help_attr())
                fg_attr = self._menu_selected_attr() if (index == selected_row and selected_col == 0) else self._help_attr()
                bg_attr = self._menu_selected_attr() if (index == selected_row and selected_col == 1) else self._help_attr()
                if fg_x < width - 1:
                    fg_text = str(fg_value).ljust(value_width)
                    self.stdscr.addnstr(y, fg_x, fg_text[: max(0, width - 1 - fg_x)], max(0, width - 1 - fg_x), fg_attr)
                if bg_x < width - 1:
                    bg_text = "-" if bg_value is None else str(bg_value)
                    bg_text = bg_text.ljust(value_width)
                    self.stdscr.addnstr(y, bg_x, bg_text[: max(0, width - 1 - bg_x)], max(0, width - 1 - bg_x), bg_attr if bg_value is not None else self._help_attr())

            hint = f" {self._tr('settings_help_1')}  {self._tr('settings_help_2')}  {self._tr('settings_help_3')} "
            if height - 2 > 0:
                settings_path_text = f" {self.settings_path.expanduser()} "
                self.stdscr.addnstr(height - 2, 0, settings_path_text.ljust(width - 1), width - 1, self._bar_attr())
            self.stdscr.addnstr(height - 1, 0, hint.ljust(width - 1), width - 1, self._bar_attr())
            self.stdscr.refresh()
            key = self.stdscr.getch()
            if key == 27:
                self.message = self._tr("settings_closed")
                return
            if key in (curses.KEY_UP, ord("k")):
                selected_row = max(0, selected_row - 1)
                continue
            if key in (curses.KEY_DOWN, ord("j"), 9):
                selected_row = min(len(rows) - 1, selected_row + 1)
                continue
            if key in (curses.KEY_LEFT, ord("h"), curses.KEY_RIGHT, ord("l")):
                direction = -1 if key in (curses.KEY_LEFT, ord("h")) else 1
                if rows[selected_row][2] is None:
                    selected_col = 0
                else:
                    selected_col = 0 if direction < 0 else 1
                continue
            if key in (10, 13):
                direction = 1
                row_key = rows[selected_row][0]
                row_map = {self._tr("theme"): "theme",
                           self._tr("date_format"): "date",
                           self._tr("time_format"): "time",
                           self._tr("active_cell"): "active",
                           self._tr("sheet_fg"): "sheet",
                           self._tr("tui_fg"): "tui",
                           self._tr("row_header_fg"): "row_header",
                           self._tr("column_header_fg"): "column_header",
                           self._tr("formula_colour"): "formula",
                           self._tr("protected_fg"): "protected",
                           self._tr("language"): "language"}
                selected_key = row_map.get(row_key, "")
                if selected_key == "theme":
                    current = THEMES.index(self.sheet.theme_name) if self.sheet.theme_name in THEMES else 0
                    self._set_theme(THEMES[(current + direction) % len(THEMES)])
                elif selected_key == "date":
                    current_style = self.sheet.date_format.split(":", 1)[1]
                    current = DATE_FORMATS.index(current_style) if current_style in DATE_FORMATS else 0
                    self._set_sheet_date_format(DATE_FORMATS[(current + direction) % len(DATE_FORMATS)])
                elif selected_key == "time":
                    current_style = self.sheet.time_format.split(":", 1)[1]
                    current = TIME_FORMATS.index(current_style) if current_style in TIME_FORMATS else 0
                    self._set_sheet_time_format(TIME_FORMATS[(current + direction) % len(TIME_FORMATS)])
                elif selected_key == "active":
                    current = ACTIVE_CELL_COLORS.index(self.sheet.active_cell_color) if self.sheet.active_cell_color in ACTIVE_CELL_COLORS else 0
                    self._set_active_cell_color(ACTIVE_CELL_COLORS[(current + direction) % len(ACTIVE_CELL_COLORS)])
                elif selected_key == "sheet":
                    if selected_col == 0:
                        current = SHEET_FG_OPTIONS.index(self.sheet.sheet_foreground_color) if self.sheet.sheet_foreground_color in SHEET_FG_OPTIONS else 0
                        self.sheet.sheet_foreground_color = SHEET_FG_OPTIONS[(current + direction) % len(SHEET_FG_OPTIONS)]
                        self._refresh_theme_colors()
                        self._save_global_settings()
                        self.dirty = True
                        self.message = self.sheet.sheet_foreground_color
                    else:
                        current = SHEET_BG_OPTIONS.index(self.sheet.sheet_background_color) if self.sheet.sheet_background_color in SHEET_BG_OPTIONS else 0
                        self.sheet.sheet_background_color = SHEET_BG_OPTIONS[(current + direction) % len(SHEET_BG_OPTIONS)]
                        self._refresh_theme_colors()
                        self._save_global_settings()
                        self.dirty = True
                        self.message = self.sheet.sheet_background_color
                elif selected_key == "tui":
                    if selected_col == 0:
                        current = TUI_COLOR_OPTIONS.index(self.sheet.tui_foreground_color) if self.sheet.tui_foreground_color in TUI_COLOR_OPTIONS else 0
                        self.sheet.tui_foreground_color = TUI_COLOR_OPTIONS[(current + direction) % len(TUI_COLOR_OPTIONS)]
                        self._refresh_theme_colors()
                        self._save_global_settings()
                        self.dirty = True
                        self.message = self.sheet.tui_foreground_color
                    else:
                        current = TUI_COLOR_OPTIONS.index(self.sheet.tui_background_color) if self.sheet.tui_background_color in TUI_COLOR_OPTIONS else 0
                        self.sheet.tui_background_color = TUI_COLOR_OPTIONS[(current + direction) % len(TUI_COLOR_OPTIONS)]
                        self._refresh_theme_colors()
                        self._save_global_settings()
                        self.dirty = True
                        self.message = self.sheet.tui_background_color
                elif selected_key == "row_header":
                    if selected_col == 0:
                        value = self._prompt_header_color(self._tr("row_header_fg"), self.sheet.row_header_foreground_color, "yellow")
                        if value is not None:
                            self.sheet.row_header_foreground_color = value
                            self._refresh_theme_colors()
                            self._save_global_settings()
                            self.dirty = True
                            self.message = self.sheet.row_header_foreground_color
                    else:
                        value = self._prompt_header_color(self._tr("row_header_bg"), self.sheet.row_header_background_color, "black")
                        if value is not None:
                            self.sheet.row_header_background_color = value
                            self._refresh_theme_colors()
                            self._save_global_settings()
                            self.dirty = True
                            self.message = self.sheet.row_header_background_color
                elif selected_key == "column_header":
                    if selected_col == 0:
                        value = self._prompt_header_color(self._tr("column_header_fg"), self.sheet.column_header_foreground_color, "yellow")
                        if value is not None:
                            self.sheet.column_header_foreground_color = value
                            self._refresh_theme_colors()
                            self._save_global_settings()
                            self.dirty = True
                            self.message = self.sheet.column_header_foreground_color
                    else:
                        value = self._prompt_header_color(self._tr("column_header_bg"), self.sheet.column_header_background_color, "black")
                        if value is not None:
                            self.sheet.column_header_background_color = value
                            self._refresh_theme_colors()
                            self._save_global_settings()
                            self.dirty = True
                            self.message = self.sheet.column_header_background_color
                elif selected_key == "formula":
                    current_value = self.sheet.formula_foreground_color if self.sheet.formula_coloration else "off"
                    current = FORMULA_COLOR_SETTING_OPTIONS.index(current_value) if current_value in FORMULA_COLOR_SETTING_OPTIONS else 0
                    self._set_formula_color_setting(FORMULA_COLOR_SETTING_OPTIONS[(current + direction) % len(FORMULA_COLOR_SETTING_OPTIONS)])
                elif selected_key == "protected":
                    if selected_col == 0:
                        current = PROTECTED_COLOR_OPTIONS.index(self.sheet.protected_foreground_color) if self.sheet.protected_foreground_color in PROTECTED_COLOR_OPTIONS else 0
                        self._set_protected_colors(foreground_name=PROTECTED_COLOR_OPTIONS[(current + direction) % len(PROTECTED_COLOR_OPTIONS)])
                    else:
                        current = PROTECTED_COLOR_OPTIONS.index(self.sheet.protected_background_color) if self.sheet.protected_background_color in PROTECTED_COLOR_OPTIONS else 0
                        self._set_protected_colors(background_name=PROTECTED_COLOR_OPTIONS[(current + direction) % len(PROTECTED_COLOR_OPTIONS)])
                elif selected_key == "language":
                    current_name = next((name for name, code in LANGUAGE_CODES.items() if code == self.sheet.language), "english")
                    current = LANGUAGE_OPTIONS.index(current_name) if current_name in LANGUAGE_OPTIONS else 0
                    self._set_language(LANGUAGE_OPTIONS[(current + direction) % len(LANGUAGE_OPTIONS)])
                continue

    def _confirm_quit(self) -> None:
        self._store_current_tab_state()
        current_tab = self.tabs[self.current_tab_index] if self.tabs else None
        if current_tab is None:
            self.running = False
            self.message = self._tr("bye")
            return
        if not current_tab.dirty:
            self._close_current_tab()
            return
        choice = self._choose_from_menu(
            self._tr("unsaved_quit"),
            [self._tr("save_quit"), self._tr("discard_quit")],
            default_option=self._tr("save_quit"),
        )
        if choice is None:
            self.message = self._tr("quit_cancelled")
            return
        if choice == self._tr("save_quit"):
            self._execute_file_command("save", [])
            self._store_current_tab_state()
            if self.tabs[self.current_tab_index].dirty:
                return
            self._close_current_tab()
            if not self.running:
                self.message = self._tr("saved_and_exited")
            return
        last_tab = len(self.tabs) == 1
        self._close_current_tab()
        if last_tab:
            self.message = self._tr("exited_without_saving")

    def execute_command(self, name: str, args: list[str]) -> None:
        try:
            if name == "quit":
                self._confirm_quit()
            elif name == "help":
                self._command_help(args)
            elif name == "save":
                self._command_save(args)
            elif name in {"load", "saveas"}:
                self._execute_file_command(name, args)
            elif name == "goto":
                spec = self._resolve_named_spec(args[0])
                self.current_row, self.current_col = parse_cell_reference(spec.split(":", 1)[0])
                self.sheet.ensure_size(self.current_row, self.current_col)
                self._scroll_into_view()
                self.message = f"Jumped to {args[0].upper()}"
            elif name == "find":
                self._command_find(args)
            elif name == "fill":
                self._command_fill(args)
            elif name == "edit":
                self._command_edit(args)
            elif name in {"blank", "protect", "unprotect"}:
                self._command_range_flag(name, args)
            elif name in {"copy", "replicate"}:
                self._command_copy(args)
            elif name == "duplicate":
                self._command_duplicate(args)
            elif name == "hide":
                self._command_hide(args)
            elif name == "unhide":
                self._command_unhide(args)
            elif name == "arrange":
                self._command_arrange(args)
            elif name == "delete":
                self._command_delete(args)
            elif name == "insert":
                self._command_insert(args)
            elif name == "move":
                self._command_move(args)
            elif name == "format":
                self._command_format(args)
            elif name == "justify":
                self._command_justify(args)
            elif name == "global":
                self._command_global(args)
            elif name == "name":
                self._command_name(args)
            elif name == "title":
                self._command_title(args)
            elif name == "output":
                self._command_output(args)
            elif name == "execute":
                self._command_execute(args)
            elif name == "redo":
                self.redo_last_action()
            elif name == "replace":
                self._command_replace(args)
            elif name == "raw":
                self.raw_sheet_view = not self.raw_sheet_view
                if self.raw_sheet_view:
                    self.message = "Raw view on (Esc to close)."
                else:
                    self.message = "Raw view off."
            elif name == "window":
                self._command_help(["commands"])
            elif name == "zap":
                self._command_zap()
            else:
                self.message = f"Unknown command: {name}"
        except (IndexError, OSError, ValueError, FormulaError) as exc:
            self.message = f"Command error: {exc}"

    def _command_save(self, args: list[str]) -> None:
        if args:
            self._execute_file_command("save", args)
            return
        choice = self._choose_from_menu("Save", ["save", "save-as", "save-quit"], default_option="save")
        if choice is None:
            self.message = self._tr("save_cancelled")
            return
        if choice == "save":
            self._execute_file_command("save", [])
            return
        if choice == "save-as":
            self._execute_file_command("saveas", [])
            return
        self._execute_file_command("save", [])
        if not self.dirty:
            self.running = False
            self.message = "Saved and exited."

    def _execute_file_command(self, name: str, args: list[str]) -> None:
        if name in {"save", "saveas"}:
            if name == "saveas":
                target = Path(args[0]).expanduser() if args else self._choose_save_target()
                if target is None:
                    self.message = self._tr("save_as_cancelled")
                    return
            else:
                if args:
                    target = Path(args[0]).expanduser()
                elif self.path is not None:
                    target = self.path
                else:
                    target = self._choose_save_target()
                    if target is None:
                        self.message = self._tr("save_cancelled")
                        return
            save_sheet(self.sheet, target)
            self.path = target
            self.dirty = False
            self._store_current_tab_state()
            self._remember_recent_file(target)
            self.message = f"Saved {target}"
            return
        if not args:
            target = self._choose_load_target()
            if target is None:
                self.message = "Load cancelled."
                return
        else:
            target = Path(args[0]).expanduser()
        self._add_loaded_tab(target, switch=True)

    def _choose_load_target(self) -> Path | None:
        recent_existing = [Path(item).expanduser() for item in self.recent_files if Path(item).expanduser().exists()]
        if recent_existing:
            duplicate_counts: dict[str, int] = {}
            for path in recent_existing:
                duplicate_counts[path.name] = duplicate_counts.get(path.name, 0) + 1
            recent_labels: list[str] = []
            label_to_path: dict[str, Path] = {}
            for path in recent_existing:
                label = path.name
                if duplicate_counts[path.name] > 1:
                    label = str(path)
                recent_labels.append(label)
                label_to_path[label] = path
            options = ["browse"] + recent_labels
            descriptions = {"browse": "Browse the filesystem for a sheet file."}
            for label, path in label_to_path.items():
                descriptions[label] = str(path)
            choice = self._choose_from_menu(
                "Load",
                options,
                default_option="browse",
                descriptions=descriptions,
                footer_hint=" arrows/type/Enter/Esc ",
            )
            if choice is None:
                return None
            if choice != "browse":
                return label_to_path.get(choice)
        return self._browse_for_file(
            "Load File",
            self.path.parent if self.path else Path.cwd(),
            suffixes={".tss", ".csv", ".tsv"},
        )

    def _choose_save_target(self) -> Path | None:
        initial = self.path if self.path is not None else DEFAULT_PATH
        return self._browse_for_save("Save File", initial)

    def _launch_menu_command(self, name: str) -> None:
        if name in {"format", "justify", "save", "help", "redo", "load"}:
            self.execute_command(name, [])
            return
        if name == "edit":
            self.show_settings_screen()
            return
        if name == "tab":
            self._command_tab()
            return
        if name == "delete":
            self._command_delete([])
            return
        if name == "quit":
            self.execute_command("quit", [])
            return
        if name in {"copy", "replicate"}:
            source = self.prompt(f"{name.title()} source: ", "", reference_origin=(self.current_row, self.current_col), range_snap=True)
            if source is None:
                self.message = self._tr("command_cancelled")
                return
            source = source.strip()
            if not source:
                self.message = self._tr("command_cancelled")
                return
            destination = self.prompt(f"{name.title()} destination: ", "", reference_origin=(self.current_row, self.current_col), range_snap=True)
            if destination is None:
                self.message = self._tr("command_cancelled")
                return
            destination = destination.strip()
            if not destination:
                self.message = self._tr("command_cancelled")
                return
            self.execute_command(name, [source, destination])
            return
        prompt_map = {
            "arrange": ("Arrange range [col] [desc]: ", "A1:C10 0"),
            "blank": ("Blank range (empty=current/selection): ", ""),
            "duplicate": ("Duplicate row|col range: ", "row 3:3"),
            "execute": ("Execute file: ", ""),
            "fill": ("Fill down|right [series] [range]: ", "down"),
            "find": ("Find text [range]: ", ""),
            "global": ("Global width n or width COL n: ", "width 14"),
            "goto": ("Goto cell: ", "A1"),
            "hide": ("Hide row|col range: ", "row 3:3"),
            "move": ("Move row|col a b [n]: ", "row 1 2 1"),
            "name": ("Name define NAME RANGE | delete NAME | list: ", "define Sales A1:A10"),
            "output": ("Output screen or file PATH: ", "screen"),
            "protect": ("Protect range (empty=current/selection): ", ""),
            "replace": ("Replace old new [range]: ", ""),
            "title": ("Title rows [cols]: ", "1 0"),
            "unhide": ("Unhide row|col range: ", "row 3:3"),
            "unprotect": ("Unprotect range (empty=current/selection): ", ""),
            "zap": ("Type YES to clear workspace: ", "NO"),
        }
        if name not in prompt_map:
            self.execute_command(name, [])
            return
        label, initial = prompt_map[name]
        reference_prompt_commands = {"blank", "goto", "protect", "unprotect"}
        text = self.prompt(
            label,
            initial,
            reference_origin=(self.current_row, self.current_col) if name in reference_prompt_commands else None,
            range_snap=name in reference_prompt_commands or name in {"arrange", "duplicate", "fill", "hide", "move", "title", "unhide"},
        )
        if text is None:
            self.message = f"{name.title()} cancelled."
            return
        cleaned = text.strip()
        if name == "zap":
            if cleaned.upper() == "YES":
                self.execute_command("zap", [])
            else:
                self.message = "Zap cancelled."
            return
        args = shlex.split(cleaned) if cleaned else []
        self.execute_command(name, args)

    def _command_tab(self) -> None:
        options = ["rename", "duplicate", "close", "next", "prev", "move left", "move right"]
        choice = self._choose_from_menu("Tab", options, default_option="next")
        if choice is None:
            self.message = "Tab menu cancelled."
            return
        if choice == "rename":
            current = self.tabs[self.current_tab_index] if self.tabs else None
            if current is None:
                self.message = "No tabs to rename."
                return
            current_label = current.name or (current.path.name if current.path else f"untitled-{self.current_tab_index + 1}")
            text = self.prompt("Tab name (blank=default): ", current_label)
            if text is None:
                self.message = "Rename cancelled."
                return
            cleaned = text.strip()
            current.name = cleaned if cleaned else None
            self.message = f"Tab renamed to {self._tab_label(self.current_tab_index, current)}"
            return
        if choice == "duplicate":
            self._duplicate_current_tab()
            return
        if choice == "close":
            self._close_current_tab()
            return
        if choice == "next":
            self.switch_tab(1)
            return
        if choice == "prev":
            self.switch_tab(-1)
            return
        if choice == "move left":
            self._move_tab(-1)
            return
        if choice == "move right":
            self._move_tab(1)
            return

    def _current_insert_range_text(self, axis: str) -> str:
        if axis.startswith("r"):
            row_text = str(self.current_row + 1)
            return f"{row_text}:{row_text}"
        col_text = column_label(self.current_col)
        return f"{col_text}:{col_text}"

    def _command_insert_interactive(self) -> None:
        axis_choice = self._choose_from_menu("Insert", ["column", "row"], default_option="column")
        if axis_choice is None:
            self.message = "Insert cancelled."
            return
        range_text = self.prompt(
            f"Insert {axis_choice} range: ",
            self._current_insert_range_text(axis_choice),
            reference_origin=(self.current_row, self.current_col) if axis_choice == "column" else None,
            range_snap=True,
        )
        if range_text is None or not range_text.strip():
            self.message = "Insert cancelled."
            return
        self.execute_command("insert", [axis_choice, range_text.strip()])

    def _command_help(self, args: list[str]) -> None:
        if args:
            topic = args[0].lower()
        else:
            topic = self._choose_from_menu("Help", HELP_TOPICS, default_option="commands")
            if topic is None:
                self.message = self._tr("help_cancelled")
                return
        topic_map = {
            "commands": (self._tr("commands"), get_command_help_lines(self.sheet.language)),
            "keys": (self._tr("keys"), get_key_help_lines(self.sheet.language)),
            "formulas": (self._tr("formulas"), get_formula_help_lines(self.sheet.language)),
            "formula": (self._tr("formulas"), get_formula_help_lines(self.sheet.language)),
        }
        if topic not in topic_map:
            raise ValueError("help topics: commands, keys, formulas")
        title, lines = topic_map[topic]
        self._show_text_page(title, lines)
        if topic.startswith("command"):
            self.message = self._tr("help_commands")
        elif topic.startswith("formula"):
            self.message = self._tr("help_formulas")
        else:
            self.message = self._tr("help_keys")

    def _show_text_page(self, title: str, lines: list[str]) -> None:
        curses.curs_set(0)
        while True:
            self.stdscr.erase()
            height, width = self.stdscr.getmaxyx()
            self.stdscr.addnstr(0, 0, f" {title} ".ljust(width - 1), width - 1, self._bar_attr(bold=True))
            body_height = max(1, height - 2)
            visible = lines[:body_height]
            for index, line in enumerate(visible, start=1):
                self.stdscr.addnstr(index, 0, line.ljust(width - 1), width - 1, self._help_attr())
            hint = " Esc/Enter/Space closes "
            self.stdscr.addnstr(height - 1, 0, hint.ljust(width - 1), width - 1, self._bar_attr())
            self.stdscr.refresh()
            key = self.stdscr.getch()
            if key in (27, 10, 13, ord(" ")):
                return

    def _draw_help_panel(self, start_y: int, panel_height: int, width: int) -> None:
        lines = get_formula_help_lines(self.sheet.language)[: max(0, panel_height - 2)]
        if panel_height <= 0:
            return
        for offset in range(panel_height):
            self.stdscr.addnstr(start_y + offset, 0, (" " * (width - 1)), width - 1, self._bar_attr())
        title = " Formula Help "
        self.stdscr.addnstr(start_y, 0, title.ljust(width - 1), width - 1, self._bar_attr(bold=True))
        for index, line in enumerate(lines, start=1):
            if start_y + index >= start_y + panel_height:
                break
            self.stdscr.addnstr(start_y + index, 0, line.ljust(width - 1), width - 1, self._help_attr())

    def _command_edit(self, args: list[str]) -> None:
        if not args:
            self.show_settings_screen()
            return
        row, col = self.current_row, self.current_col
        value_args = args
        if args[0][0].isalpha() and any(char.isdigit() for char in args[0]):
            row, col = parse_cell_reference(args[0])
            value_args = args[1:]
        if self.sheet.is_protected(row, col):
            raise ValueError("cell is protected")
        value = self._normalize_cell_input(row, col, " ".join(value_args))
        self._save_undo_state()
        self.sheet.set_raw(row, col, value)
        self._apply_default_alignment(row, col, value)
        self.dirty = True
        self.message = f"Stored {column_label(col)}{row + 1}"

    def _command_find(self, args: list[str]) -> None:
        if not args:
            raise ValueError("find needs text")
        needle = args[0]
        if len(args) > 1:
            row_lo, col_lo, row_hi, col_hi = self._target_range(args[1])
        else:
            row_lo, col_lo, row_hi, col_hi = 0, 0, self.sheet.rows - 1, self.sheet.cols - 1
        matches = [
            (row, col)
            for row in range(row_lo, row_hi + 1)
            for col in range(col_lo, col_hi + 1)
            if needle.lower() in self.sheet.get_raw(row, col).lower()
        ]
        if not matches:
            self.message = f"No match for {needle!r}"
            return
        current = (self.current_row, self.current_col)
        for row, col in matches:
            if (row, col) > current:
                self.current_row, self.current_col = row, col
                self.selection_range = (row, col, row, col)
                self._scroll_into_view()
                self.message = f"Found {needle!r} at {column_label(col)}{row + 1}"
                return
        row, col = matches[0]
        self.current_row, self.current_col = row, col
        self.selection_range = (row, col, row, col)
        self._scroll_into_view()
        self.message = f"Found {needle!r} at {column_label(col)}{row + 1}"

    def _command_replace(self, args: list[str]) -> None:
        if len(args) < 2:
            raise ValueError("replace needs old and new text")
        old_text, new_text = args[0], args[1]
        if len(args) > 2:
            row_lo, col_lo, row_hi, col_hi = self._target_range(args[2])
        elif self.selection_range is not None:
            row_lo, col_lo, row_hi, col_hi = self.selection_range
        else:
            row_lo, col_lo, row_hi, col_hi = 0, 0, self.sheet.rows - 1, self.sheet.cols - 1
        changed = 0
        self._save_undo_state()
        for row in range(row_lo, row_hi + 1):
            for col in range(col_lo, col_hi + 1):
                raw = self.sheet.get_raw(row, col)
                if old_text not in raw:
                    continue
                self.sheet.set_raw(row, col, raw.replace(old_text, new_text))
                changed += 1
        if not changed:
            self.undo_stack.pop()
            self.message = f"No replacements for {old_text!r}"
            return
        self.dirty = True
        self.message = f"Replaced {changed} cell(s) in {self._range_label(row_lo, col_lo, row_hi, col_hi)}"

    def _command_range_flag(self, name: str, args: list[str]) -> None:
        row_lo, col_lo, row_hi, col_hi = self._target_range(args[0] if args else None)
        self._save_undo_state()
        for row in range(row_lo, row_hi + 1):
            for col in range(col_lo, col_hi + 1):
                if name == "blank":
                    if not self.sheet.is_protected(row, col):
                        self.sheet.clear(row, col)
                        self.sheet.set_format(row, col, "")
                        self.sheet.clear_text_styles(row, col)
                        self.sheet.set_background(row, col, "")
                        self.sheet.set_border(row, col, "")
                        if not self.sheet.is_alignment_manual(row, col):
                            self.sheet.set_alignment(row, col, "", manual=False)
                elif name == "protect":
                    self.sheet.protect(row, col)
                else:
                    self.sheet.unprotect(row, col)
        self.dirty = True
        self.message = f"{name.title()} applied to {self._range_label(row_lo, col_lo, row_hi, col_hi)}"

    def _command_hide(self, args: list[str]) -> None:
        if len(args) < 2:
            raise ValueError("hide needs row|col and a range")
        axis = args[0].lower()
        row_lo, col_lo, row_hi, col_hi = self._parse_range_spec(args[1])
        self._save_undo_state()
        if axis.startswith("r"):
            for row in range(row_lo, row_hi + 1):
                self.sheet.hide_row(row)
            self.current_row = self._first_visible_row()
            self.message = f"Hidden row(s) {row_lo + 1}:{row_hi + 1}"
        else:
            for col in range(col_lo, col_hi + 1):
                self.sheet.hide_col(col)
            self.current_col = self._first_visible_col()
            self.message = f"Hidden column(s) {column_label(col_lo)}:{column_label(col_hi)}"
        self._scroll_into_view()
        self.dirty = True

    def _command_unhide(self, args: list[str]) -> None:
        if not args:
            raise ValueError("unhide needs row|col and a range or 'all'")
        axis = args[0].lower()
        self._save_undo_state()
        if len(args) > 1 and args[1].lower() == "all":
            if axis.startswith("r"):
                self.sheet.hidden_rows.clear()
                self.message = "Unhid all rows."
            else:
                self.sheet.hidden_cols.clear()
                self.message = "Unhid all columns."
            self.dirty = True
            return
        if len(args) < 2:
            raise ValueError("unhide needs row|col and a range")
        row_lo, col_lo, row_hi, col_hi = self._parse_range_spec(args[1])
        if axis.startswith("r"):
            for row in range(row_lo, row_hi + 1):
                self.sheet.unhide_row(row)
            self.message = f"Unhid row(s) {row_lo + 1}:{row_hi + 1}"
        else:
            for col in range(col_lo, col_hi + 1):
                self.sheet.unhide_col(col)
            self.message = f"Unhid column(s) {column_label(col_lo)}:{column_label(col_hi)}"
        self.dirty = True

    def _duplicate_rows(self, row_lo: int, row_hi: int) -> None:
        count = row_hi - row_lo + 1
        destination = row_hi + 1
        self._rebuild_rows(destination, count)
        for offset in range(count):
            src_row = row_lo + offset
            dst_row = destination + offset
            for col in range(self.sheet.cols):
                raw = self.sheet.get_raw(src_row, col)
                self.sheet.set_raw(dst_row, col, shift_formula_references(raw, dst_row - src_row, 0))
                self.sheet.set_format(dst_row, col, self.sheet.get_format(src_row, col))
                self.sheet.clear_text_styles(dst_row, col)
                for style in self.sheet.get_text_styles(src_row, col):
                    self.sheet.set_text_style(dst_row, col, style, enabled=True)
                self.sheet.set_background(dst_row, col, self.sheet.get_background(src_row, col))
                self.sheet.set_border(dst_row, col, self.sheet.get_border(src_row, col))
                self.sheet.set_alignment(dst_row, col, self.sheet.get_alignment(src_row, col), manual=self.sheet.is_alignment_manual(src_row, col))
                if self.sheet.is_protected(src_row, col):
                    self.sheet.protect(dst_row, col)
                else:
                    self.sheet.unprotect(dst_row, col)

    def _duplicate_columns(self, col_lo: int, col_hi: int) -> None:
        count = col_hi - col_lo + 1
        destination = col_hi + 1
        self._rebuild_cols(destination, count)
        for offset in range(count):
            src_col = col_lo + offset
            dst_col = destination + offset
            self.sheet.set_column_width(dst_col, self.sheet.get_column_width(src_col))
            for row in range(self.sheet.rows):
                raw = self.sheet.get_raw(row, src_col)
                self.sheet.set_raw(row, dst_col, shift_formula_references(raw, 0, dst_col - src_col))
                self.sheet.set_format(row, dst_col, self.sheet.get_format(row, src_col))
                self.sheet.clear_text_styles(row, dst_col)
                for style in self.sheet.get_text_styles(row, src_col):
                    self.sheet.set_text_style(row, dst_col, style, enabled=True)
                self.sheet.set_background(row, dst_col, self.sheet.get_background(row, src_col))
                self.sheet.set_border(row, dst_col, self.sheet.get_border(row, src_col))
                self.sheet.set_alignment(row, dst_col, self.sheet.get_alignment(row, src_col), manual=self.sheet.is_alignment_manual(row, src_col))
                if self.sheet.is_protected(row, src_col):
                    self.sheet.protect(row, dst_col)
                else:
                    self.sheet.unprotect(row, dst_col)

    def _command_duplicate(self, args: list[str]) -> None:
        if len(args) < 2:
            raise ValueError("duplicate needs row|col and a range")
        axis = args[0].lower()
        row_lo, col_lo, row_hi, col_hi = self._parse_range_spec(args[1])
        self._save_undo_state()
        if axis.startswith("r"):
            self._duplicate_rows(row_lo, row_hi)
            self.message = f"Duplicated row(s) {row_lo + 1}:{row_hi + 1}"
        else:
            self._duplicate_columns(col_lo, col_hi)
            self.message = f"Duplicated column(s) {column_label(col_lo)}:{column_label(col_hi)}"
        self.dirty = True

    def _coerce_series_value(self, raw: str) -> float | None:
        if not raw or raw.startswith("'") or is_formula_text(raw):
            return None
        try:
            return float(raw.replace(",", ""))
        except ValueError:
            try:
                normalized = normalize_date_text(raw, self.sheet.date_format)
                date_obj = parse_date_text(normalized, "date:ansi")
                return float(date_obj.toordinal()) if date_obj else None
            except ValueError:
                return None

    def _render_series_value(self, value: float, template: str) -> str:
        try:
            normalized = normalize_date_text(template, self.sheet.date_format)
            if parse_date_text(normalized, "date:ansi") is not None:
                from datetime import date
                return date.fromordinal(int(round(value))).strftime("%Y-%m-%d")
        except ValueError:
            pass
        if template.strip().isdigit():
            return str(int(round(value)))
        if "." in template:
            digits = len(template.split(".", 1)[1])
            return f"{value:.{digits}f}"
        return str(value)

    def _command_fill(self, args: list[str]) -> None:
        direction = args[0].lower() if args else "down"
        smart_series = len(args) > 1 and args[1].lower() == "series"
        row_lo, col_lo, row_hi, col_hi = self._target_range(args[2] if len(args) > 2 else None)
        if direction not in {"down", "right"}:
            raise ValueError("fill direction must be down or right")
        self._save_undo_state()
        if direction == "down":
            if smart_series and row_hi > row_lo:
                for col in range(col_lo, col_hi + 1):
                    first = self.sheet.get_raw(row_lo, col)
                    second = self.sheet.get_raw(row_lo + 1, col)
                    first_num = self._coerce_series_value(first)
                    second_num = self._coerce_series_value(second)
                    step = 1.0 if first_num is None or second_num is None else second_num - first_num
                    current = second_num if second_num is not None else first_num
                    template = second or first
                    for row in range(row_lo + 2, row_hi + 1):
                        if current is None or not template:
                            self.sheet.set_raw(row, col, first)
                        else:
                            current += step
                            self.sheet.set_raw(row, col, self._render_series_value(current, template))
                        self.sheet.set_format(row, col, self.sheet.get_format(row_lo + 1, col) or self.sheet.get_format(row_lo, col))
            else:
                src_row = row_lo
                for row in range(row_lo + 1, row_hi + 1):
                    for col in range(col_lo, col_hi + 1):
                        raw = self.sheet.get_raw(src_row, col)
                        self.sheet.set_raw(row, col, shift_formula_references(raw, row - src_row, 0))
                        self.sheet.set_format(row, col, self.sheet.get_format(src_row, col))
                        self.sheet.clear_text_styles(row, col)
                        for style in self.sheet.get_text_styles(src_row, col):
                            self.sheet.set_text_style(row, col, style, enabled=True)
                        self.sheet.set_background(row, col, self.sheet.get_background(src_row, col))
                        self.sheet.set_border(row, col, self.sheet.get_border(src_row, col))
        else:
            if smart_series and col_hi > col_lo:
                for row in range(row_lo, row_hi + 1):
                    first = self.sheet.get_raw(row, col_lo)
                    second = self.sheet.get_raw(row, col_lo + 1)
                    first_num = self._coerce_series_value(first)
                    second_num = self._coerce_series_value(second)
                    step = 1.0 if first_num is None or second_num is None else second_num - first_num
                    current = second_num if second_num is not None else first_num
                    template = second or first
                    for col in range(col_lo + 2, col_hi + 1):
                        if current is None or not template:
                            self.sheet.set_raw(row, col, first)
                        else:
                            current += step
                            self.sheet.set_raw(row, col, self._render_series_value(current, template))
                        self.sheet.set_format(row, col, self.sheet.get_format(row, col_lo + 1) or self.sheet.get_format(row, col_lo))
            else:
                src_col = col_lo
                for row in range(row_lo, row_hi + 1):
                    for col in range(col_lo + 1, col_hi + 1):
                        raw = self.sheet.get_raw(row, src_col)
                        self.sheet.set_raw(row, col, shift_formula_references(raw, 0, col - src_col))
                        self.sheet.set_format(row, col, self.sheet.get_format(row, src_col))
                        self.sheet.clear_text_styles(row, col)
                        for style in self.sheet.get_text_styles(row, src_col):
                            self.sheet.set_text_style(row, col, style, enabled=True)
                        self.sheet.set_background(row, col, self.sheet.get_background(row, src_col))
                        self.sheet.set_border(row, col, self.sheet.get_border(row, src_col))
        self.dirty = True
        mode = "series " if smart_series else ""
        self.message = f"Filled {mode}{direction} on {self._range_label(row_lo, col_lo, row_hi, col_hi)}"

    def _command_copy(self, args: list[str]) -> None:
        if len(args) < 2:
            default_source = self._selection_label() or f"{column_label(self.current_col)}{self.current_row + 1}"
            source = self.prompt("Copy source: ", default_source, reference_origin=(self.current_row, self.current_col), range_snap=True)
            if source is None or not source.strip():
                self.message = "Copy cancelled."
                return
            destination = self.prompt("Copy destination: ", "", reference_origin=(self.current_row, self.current_col), range_snap=True)
            if destination is None or not destination.strip():
                self.message = "Copy cancelled."
                return
            args = [source.strip(), destination.strip()]
        src_lo_r, src_lo_c, src_hi_r, src_hi_c = self._parse_range_spec(args[0])
        dst_lo_r, dst_lo_c, dst_hi_r, dst_hi_c = self._parse_range_spec(args[1])
        self._load_internal_clipboard_from_range(src_lo_r, src_lo_c, src_hi_r, src_hi_c)
        self._paste_internal_clipboard((dst_lo_r, dst_lo_c, dst_hi_r, dst_hi_c))
        self.message = f"Copied {args[0].upper()} to {args[1].upper()}"

    def _command_arrange(self, args: list[str]) -> None:
        if not args:
            raise ValueError("arrange needs a range")
        row_lo, col_lo, row_hi, col_hi = self._parse_range_spec(args[0])
        sort_offset = int(args[1]) if len(args) > 1 else 0
        descending = len(args) > 2 and args[2].lower().startswith("d")
        sort_col = col_lo + sort_offset
        self._save_undo_state()
        records: list[list[tuple[str, str, str, str, bool]]] = []
        for row in range(row_lo, row_hi + 1):
            record = []
            for col in range(col_lo, col_hi + 1):
                record.append(
                    (
                        self.sheet.get_raw(row, col),
                        self.sheet.get_format(row, col),
                        ",".join(sorted(self.sheet.get_text_styles(row, col))),
                        self.sheet.get_background(row, col),
                        self.sheet.is_protected(row, col),
                    )
                )
            records.append(record)
        def sort_key(record: list[tuple[str, str, str, str, bool]]) -> str:
            index = max(0, min(len(record) - 1, sort_col - col_lo))
            return record[index][0]
        records.sort(key=sort_key, reverse=descending)
        for row_index, record in enumerate(records, start=row_lo):
            for col_index, (raw, style, text_styles, background, protected) in enumerate(record, start=col_lo):
                self.sheet.set_raw(row_index, col_index, raw)
                self.sheet.set_format(row_index, col_index, style)
                self.sheet.clear_text_styles(row_index, col_index)
                for text_style in [item for item in text_styles.split(",") if item]:
                    self.sheet.set_text_style(row_index, col_index, text_style, enabled=True)
                self.sheet.set_background(row_index, col_index, background)
                if protected:
                    self.sheet.protect(row_index, col_index)
                else:
                    self.sheet.unprotect(row_index, col_index)
        self.dirty = True
        self.message = f"Arranged {self._range_label(row_lo, col_lo, row_hi, col_hi)}"

    def _command_delete(self, args: list[str]) -> None:
        if not args:
            axis = self._choose_from_menu("Delete", ["row", "column", "cells"], default_option="row")
            if axis is None:
                self.message = "Delete cancelled."
                return
            args = [axis]
        axis = args[0].lower()
        if axis in {"cell", "cells"}:
            if len(args) < 2:
                default_text = self._selection_label() if self.selection_range else f"{column_label(self.current_col)}{self.current_row + 1}"
                range_text = self.prompt("Delete cells: ", default_text, reference_origin=(self.current_row, self.current_col), range_snap=True)
                if range_text is None or not range_text.strip():
                    self.message = "Delete cancelled."
                    return
            else:
                range_text = args[1]
            row_lo, col_lo, row_hi, col_hi = self._parse_range_spec(range_text)
            choice = None
            if len(args) >= 3:
                choice = args[2].lower()
            if choice not in {"up", "left", "clear"}:
                choice = self._choose_from_menu("Delete cells", ["move up", "move left", "clear"], default_option="clear")
                if choice is None:
                    self.message = "Delete cancelled."
                    return
                if choice.startswith("move up"):
                    choice = "up"
                elif choice.startswith("move left"):
                    choice = "left"
                else:
                    choice = "clear"
            self._save_undo_state()
            self._delete_cells(row_lo, col_lo, row_hi, col_hi, choice)
            self.dirty = True
            self.message = f"Deleted cells {self._range_label(row_lo, col_lo, row_hi, col_hi)} ({choice})"
            return
        if axis in {"row", "rows", "r"}:
            if len(args) < 2:
                default_text = self._selection_label() if self._selection_label().isdigit() else str(self.current_row + 1)
                range_text = self.prompt(
                    "Delete row: ",
                    default_text,
                    help_lines=[" Use e.g. 1:5"],
                    reference_origin=(self.current_row, self.current_col),
                    range_snap=True,
                )
                if range_text is None or not range_text.strip():
                    self.message = "Delete cancelled."
                    return
            else:
                range_text = args[1]
            row_lo, _col_lo, row_hi, _col_hi = self._parse_range_spec(range_text)
            count = row_hi - row_lo + 1
            self._save_undo_state()
            self._rebuild_rows(row_lo, -count)
            self.message = f"Deleted {count} row(s) at {row_lo + 1}"
            self.dirty = True
            return
        if axis in {"col", "column", "c"}:
            if len(args) < 2:
                default_text = self._selection_label() if self._selection_label().isalpha() else column_label(self.current_col)
                range_text = self.prompt(
                    "Delete column: ",
                    default_text,
                    help_lines=[" Use e.g. A:D"],
                    reference_origin=(self.current_row, self.current_col),
                    range_snap=True,
                )
                if range_text is None or not range_text.strip():
                    self.message = "Delete cancelled."
                    return
            else:
                range_text = args[1]
            _row_lo, col_lo, _row_hi, col_hi = self._parse_range_spec(range_text)
            count = col_hi - col_lo + 1
            self._save_undo_state()
            self._rebuild_cols(col_lo, -count)
            self.message = f"Deleted {count} column(s) at {column_label(col_lo)}"
            self.dirty = True
            return
        raise ValueError("delete needs row or column")

    def _command_insert(self, args: list[str]) -> None:
        if not args:
            self._command_insert_interactive()
            return
        axis = args[0].lower()
        if axis in {"col", "column", "c"}:
            index = self.current_col
            if len(args) < 2:
                range_text = self._current_insert_range_text(axis)
                col_lo = col_hi = self.current_col
            else:
                range_text = args[1]
                _row_lo, col_lo, _row_hi, col_hi = self._parse_range_spec(range_text)
            count = col_hi - col_lo + 1
            if count <= 0:
                raise ValueError("insert column range must not be empty")
            self._save_undo_state()
            self._rebuild_cols(index, count)
            self.dirty = True
            self.message = f"Inserted {count} column(s) at {column_label(index)} from {range_text.upper()}"
            return
        if axis in {"row", "r"}:
            index = self.current_row
            if len(args) < 2:
                range_text = self._current_insert_range_text(axis)
                row_lo = row_hi = self.current_row
            else:
                range_text = args[1]
                row_lo, _col_lo, row_hi, _col_hi = self._parse_range_spec(range_text)
            count = row_hi - row_lo + 1
            if count <= 0:
                raise ValueError("insert row range must not be empty")
            self._save_undo_state()
            self._rebuild_rows(index, count)
            self.dirty = True
            self.message = f"Inserted {count} row(s) at {index + 1} from {range_text.upper()}"
            return
        if len(args) < 2:
            raise ValueError("insert needs row/column and a range")
        index = int(args[1]) - 1
        count = int(args[2]) if len(args) > 2 else 1
        self._save_undo_state()
        if axis.startswith("r"):
            self._rebuild_rows(index, count)
            self.message = f"Inserted {count} row(s) at {index + 1}"
        else:
            self._rebuild_cols(index, count)
            self.message = f"Inserted {count} column(s) at {index + 1}"
        self.dirty = True

    def _command_move(self, args: list[str]) -> None:
        axis = args[0].lower()
        start = int(args[1]) - 1
        destination = int(args[2]) - 1
        count = int(args[3]) if len(args) > 3 else 1
        self._save_undo_state()
        if axis.startswith("r"):
            self._move_rows(start, destination, count)
            self.message = f"Moved {count} row(s) from {start + 1} to {destination + 1}"
        else:
            self._move_cols(start, destination, count)
            self.message = f"Moved {count} column(s) from {start + 1} to {destination + 1}"
        self.dirty = True

    def _command_format(self, args: list[str]) -> None:
        if not args:
            style = self._choose_from_menu("Format", FORMAT_STYLES)
            if style is None:
                self.message = "Format cancelled."
                return
            args = [style]
        style = args[0].lower()
        if style in {"b", "bg", "back", "background"}:
            self._command_background(args[1:])
            return
        if style in {"row-background", "rowbg", "bgrow", "row-bg"}:
            self._command_row_background(args[1:])
            return
        if style in {"border", "borders"}:
            self._command_border(args[1:])
            return
        if style == "date":
            date_arg = args[1].lower() if len(args) > 1 else ""
            if date_arg in DATE_FORMATS:
                selected_date = date_arg
            else:
                selected_date = self._choose_from_menu("Date", DATE_FORMATS, default_option="european")
                if selected_date is None:
                    self.message = "Date format cancelled."
                    return
            self._set_sheet_date_format(selected_date)
            return
        if style == "time":
            time_arg = args[1].lower() if len(args) > 1 else ""
            if time_arg in TIME_FORMATS:
                selected_time = time_arg
                range_arg = args[2] if len(args) > 2 else None
            else:
                selected_time = self._choose_from_menu("Time", TIME_FORMATS, default_option="24h")
                if selected_time is None:
                    self.message = "Time format cancelled."
                    return
                range_arg = args[1] if len(args) > 1 else None
            format_value = f"time:{selected_time}"
            row_lo, col_lo, row_hi, col_hi = self._target_range(range_arg)
            self._save_undo_state()
            for row in range(row_lo, row_hi + 1):
                for col in range(col_lo, col_hi + 1):
                    self.sheet.set_format(row, col, format_value)
            self.dirty = True
            self.message = f"Format {format_value} set on {self._range_label(row_lo, col_lo, row_hi, col_hi)}"
            return
        if style in {"clear", "clear-format", "remove-format", "none"}:
            row_lo, col_lo, row_hi, col_hi = self._target_range(args[1] if len(args) > 1 else None)
            self._save_undo_state()
            for row in range(row_lo, row_hi + 1):
                for col in range(col_lo, col_hi + 1):
                    self.sheet.set_format(row, col, "")
                    self.sheet.clear_text_styles(row, col)
                    self.sheet.set_background(row, col, "")
                    self.sheet.set_border(row, col, "")
            self.dirty = True
            self.message = f"Formatting cleared on {self._range_label(row_lo, col_lo, row_hi, col_hi)}"
            return
        if style not in {"text", "bold", "underline", "italic", "currency", "fixed", "percent", "int", "negative", "accounting", "sci", "scientific"}:
            raise ValueError("format must be clear-format, text, bold, underline, italic, border, row-background, currency, date, time, fixed, percent, int, negative, accounting, sci, or b")
        format_value = "" if style == "text" else style
        range_arg = "."
        if style in {"bold", "underline", "italic"}:
            range_arg = args[1] if len(args) > 1 else None
            row_lo, col_lo, row_hi, col_hi = self._target_range(range_arg)
            self._save_undo_state()
            for row in range(row_lo, row_hi + 1):
                for col in range(col_lo, col_hi + 1):
                    enabled = not self.sheet.has_text_style(row, col, style)
                    self.sheet.set_text_style(row, col, style, enabled=enabled)
            self.dirty = True
            self.message = f"Style {style} toggled on {self._range_label(row_lo, col_lo, row_hi, col_hi)}"
            return
        if style == "currency":
            symbol_arg = args[1] if len(args) > 1 else ""
            if symbol_arg in CURRENCY_SYMBOLS:
                symbol = symbol_arg
                range_arg = args[2] if len(args) > 2 else "."
            else:
                symbol = self._choose_from_menu("Currency", CURRENCY_SYMBOLS, default_option="£")
                if symbol is None:
                    self.message = "Currency format cancelled."
                    return
                range_arg = args[1] if len(args) > 1 else None
            format_value = f"currency:{symbol}"
        else:
            range_arg = args[1] if len(args) > 1 else None
        row_lo, col_lo, row_hi, col_hi = self._target_range(range_arg)
        self._save_undo_state()
        for row in range(row_lo, row_hi + 1):
            for col in range(col_lo, col_hi + 1):
                self.sheet.set_format(row, col, format_value)
        self.dirty = True
        label = format_value if format_value else "text"
        self.message = f"Format {label} set on {self._range_label(row_lo, col_lo, row_hi, col_hi)}"

    def _command_background(self, args: list[str]) -> None:
        color_arg = args[0].lower() if args else ""
        if color_arg in BACKGROUND_COLORS:
            color = "" if color_arg == "none" else color_arg
            range_arg = args[1] if len(args) > 1 else None
        else:
            selected = self._choose_from_menu("Background", BACKGROUND_COLORS, default_option="blue")
            if selected is None:
                self.message = "Background cancelled."
                return
            color = "" if selected == "none" else selected
            range_arg = args[0] if args else None
        row_lo, col_lo, row_hi, col_hi = self._target_range(range_arg)
        self._save_undo_state()
        for row in range(row_lo, row_hi + 1):
            for col in range(col_lo, col_hi + 1):
                self.sheet.set_background(row, col, color)
        self.dirty = True
        label = color or "none"
        self.message = f"Background {label} set on {self._range_label(row_lo, col_lo, row_hi, col_hi)}"

    def _command_row_background(self, args: list[str]) -> None:
        color_value = args[0].strip() if args else ""
        if not color_value:
            color_value = self.prompt(
                "Row background hex: ",
                "",
                help_lines=[
                    " Enter 6 hex chars like FFCC99",
                    " Type none to clear selected rows",
                ],
            ).strip()
        if not color_value:
            self.message = "Row background cancelled."
            return
        if color_value.lower() == "none":
            color = ""
        elif HEX_COLOR_RE.match(color_value):
            color = color_value.upper()
            if not color.startswith("#"):
                color = f"#{color}"
        else:
            raise ValueError("row background needs 6 hex characters or none")
        if self.selection_range is not None:
            row_lo, _col_lo, row_hi, _col_hi = self.selection_range
        else:
            row_lo = row_hi = self.current_row
        self._save_undo_state()
        for row in range(row_lo, row_hi + 1):
            self.sheet.set_row_background(row, color)
        self.dirty = True
        if row_lo == row_hi:
            self.message = f"Row background {(color or 'none')} set on row {row_lo + 1}"
        else:
            self.message = f"Row background {(color or 'none')} set on rows {row_lo + 1}:{row_hi + 1}"

    def _command_border(self, args: list[str]) -> None:
        border_arg = args[0].lower() if args else ""
        if border_arg in BORDER_STYLES:
            border = "" if border_arg == "none" else border_arg
            range_arg = args[1] if len(args) > 1 else None
        else:
            selected = self._choose_from_menu("Border", BORDER_STYLES, default_option="outline")
            if selected is None:
                self.message = "Border cancelled."
                return
            border = "" if selected == "none" else selected
            range_arg = args[0] if args else None
        row_lo, col_lo, row_hi, col_hi = self._target_range(range_arg)
        self._save_undo_state()
        for row in range(row_lo, row_hi + 1):
            for col in range(col_lo, col_hi + 1):
                self.sheet.set_border(row, col, border)
        self.dirty = True
        self.message = f"Border {(border or 'none')} set on {self._range_label(row_lo, col_lo, row_hi, col_hi)}"

    def _command_justify(self, args: list[str]) -> None:
        if not args:
            selected = self._choose_from_menu("Justify", JUSTIFY_OPTIONS)
            if selected is None:
                self.message = "Justify cancelled."
                return
            args = [selected]
        align = args[0].lower()
        align_map = {"l": "left", "c": "center", "r": "right", "left": "left", "centre": "center", "center": "center", "right": "right"}
        if align not in align_map:
            raise ValueError("justify must be left, centre, or right")
        resolved = align_map[align]
        row_lo, col_lo, row_hi, col_hi = self._target_range(args[1] if len(args) > 1 else None)
        self._save_undo_state()
        for row in range(row_lo, row_hi + 1):
            for col in range(col_lo, col_hi + 1):
                self.sheet.set_alignment(row, col, resolved)
        self.dirty = True
        self.message = f"Justify {resolved} set on {self._range_label(row_lo, col_lo, row_hi, col_hi)}"

    def _command_theme(self, args: list[str]) -> None:
        if args:
            requested = args[0].lower()
            if requested not in THEMES:
                raise ValueError(f"theme must be one of: {', '.join(THEMES)}")
            self._save_undo_state()
            self.sheet.theme_name = requested
        else:
            selected = self._choose_from_menu("Theme", THEMES, default_option=self.sheet.theme_name)
            if selected is None:
                self.message = "Theme cancelled."
                return
            self._save_undo_state()
            self.sheet.theme_name = selected
        self._refresh_theme_colors()
        self.dirty = True
        self.message = f"Theme set to {self.sheet.theme_name}"

    def _command_global(self, args: list[str]) -> None:
        if not args or args[0].lower() != "width":
            raise ValueError("global supports: width N or width COL N")
        if len(args) >= 3:
            width = max(8, int(args[2]))
            self._save_undo_state()
            _row_lo, col_lo, _row_hi, col_hi = self._parse_range_spec(args[1])
            for col in range(col_lo, col_hi + 1):
                self.sheet.set_column_width(col, width)
            self.dirty = True
            self.message = f"Width set to {width} for {column_label(col_lo)}:{column_label(col_hi)}"
            return
        self._save_undo_state()
        self.sheet.column_width = max(8, int(args[1]))
        self.dirty = True
        self.message = f"Default column width set to {self.sheet.column_width}"

    def _command_title(self, args: list[str]) -> None:
        self._save_undo_state()
        self.sheet.title_rows = max(0, int(args[0])) if args else 0
        self.sheet.title_cols = max(0, int(args[1])) if len(args) > 1 else 0
        self.dirty = True
        self.message = f"Title freeze rows={self.sheet.title_rows} cols={self.sheet.title_cols}"

    def _command_name(self, args: list[str]) -> None:
        if not args:
            text = self.prompt("Name define NAME RANGE | delete NAME | list: ", "list")
            if text is None or not text.strip():
                self.message = "Name cancelled."
                return
            args = shlex.split(text.strip())
        action = args[0].lower()
        if action == "list":
            lines = [f"{name} = {spec}" for name, spec in sorted(self.sheet.named_ranges.items())] or ["No named ranges."]
            self._show_text_page("Named Ranges", lines)
            self.message = "Named ranges shown."
            return
        if action == "define":
            if len(args) < 3:
                raise ValueError("name define needs NAME and RANGE")
            name = args[1]
            spec = self._resolve_named_spec(args[2])
            self._parse_range_spec(spec)
            self._save_undo_state()
            self.sheet.set_named_range(name, spec)
            self.dirty = True
            self.message = f"Named range {name.upper()} = {spec.upper()}"
            return
        if action in {"delete", "remove"}:
            if len(args) < 2:
                raise ValueError("name delete needs NAME")
            self._save_undo_state()
            self.sheet.set_named_range(args[1], "")
            self.dirty = True
            self.message = f"Removed named range {args[1].upper()}"
            return
        raise ValueError("name must be define, delete, or list")

    def _command_output(self, args: list[str]) -> None:
        if not args or args[0].lower() == "screen":
            snapshot = self.render_delimited_snapshot("\t").splitlines() or ["[No populated cells]"]
            self._show_text_page("Output", snapshot)
            self.message = "Output shown on screen."
            return
        if args[0].lower() == "file":
            if len(args) < 2:
                raise ValueError("output needs: screen or PATH")
            target_text = args[1]
        else:
            target_text = args[0]
        target = Path(target_text).expanduser()
        if target.suffix.lower() in {".csv", ".tsv"}:
            save_sheet(self.sheet, target)
        elif target.suffix.lower() == ".pdf":
            lines = self.render_fixed_width_snapshot()
            title = str(self.path) if self.path else APP_NAME
            save_pdf_text(lines, target, title=title)
        else:
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(self.render_delimited_snapshot("\t"), encoding="utf-8")
        self.message = f"Output written to {target}"

    def _command_execute(self, args: list[str]) -> None:
        path = Path(args[0]).expanduser()
        for line in path.read_text(encoding="utf-8").splitlines():
            cleaned = line.strip()
            if not cleaned or cleaned.startswith("#"):
                continue
            command = parse_command(cleaned)
            if command.name == "execute":
                continue
            self.execute_command(command.name, command.args)
        self.message = f"Executed {path}"

    def _command_zap(self) -> None:
        self._save_undo_state()
        self.sheet = Spreadsheet(rows=100, cols=52)
        self.evaluator = Evaluator(self.sheet)
        self._refresh_theme_colors()
        self.current_row = 0
        self.current_col = 0
        self.row_offset = 0
        self.col_offset = 0
        self.path = None
        self.dirty = False
        self.message = "Workspace cleared."

    def prompt(
        self,
        label: str,
        initial: str,
        help_lines: list[str] | None = None,
        formula_origin: tuple[int, int] | None = None,
        reference_origin: tuple[int, int] | None = None,
        range_snap: bool = False,
    ) -> str | None:
        if formula_origin is not None:
            # Always seed formula editing from the latest stored cell text so
            # copy/paste mutations cannot reopen with stale prompt content.
            stored = self.sheet.get_raw(*formula_origin)
            if not initial or initial == stored:
                initial = stored
        text = list(initial)
        position = len(text)
        curses.curs_set(1)
        original_current = (self.current_row, self.current_col)
        effective_reference = reference_origin
        if range_snap and effective_reference is None and formula_origin is None:
            effective_reference = original_current
        ref_row, ref_col = formula_origin if formula_origin is not None else (effective_reference if effective_reference is not None else original_current)
        inserted_ref: tuple[int, int] | None = None
        while True:
            height, width = self.stdscr.getmaxyx()
            if formula_origin is not None or effective_reference is not None:
                self.current_row = ref_row
                self.current_col = ref_col
                self._scroll_into_view()
            if formula_origin is not None or effective_reference is not None or help_lines:
                self.draw()
            if help_lines:
                panel_height = min(len(help_lines) + 2, max(4, height - 4))
                self._draw_help_panel(2, panel_height, width)
            display = "".join(text)
            if formula_origin is not None:
                hint_lines = self._formula_prompt_hints(display, position, ref_row, ref_col)
                error_line = self._formula_prompt_error(display)
                prompt_lines = hint_lines + ([error_line] if error_line else [])
                self._draw_formula_prompt_hints(height - 1 - len(prompt_lines), width, prompt_lines)
            elif effective_reference is not None:
                pointer_line = f" Pointing: {self._formula_reference_text(ref_row, ref_col)}   arrows point to cells, : starts a range, Enter accepts "
                self._draw_formula_prompt_hints(height - 2, width, [pointer_line])
            prompt_attr = self._prompt_attr(bold=True)
            self.stdscr.addnstr(height - 1, 0, (" " * (width - 1)), width - 1, curses.A_NORMAL)
            self.stdscr.addnstr(height - 1, 0, f"{label}{display}", width - 1, prompt_attr)
            cursor_x = min(width - 2, len(label) + position)
            self.stdscr.move(height - 1, cursor_x)
            self.stdscr.refresh()
            key = self.stdscr.getch()
            if key in (10, 13):
                curses.curs_set(0)
                result = "".join(text)
                if formula_origin is not None:
                    if inserted_ref is None and self._formula_reference_context(result, position):
                        result += self._formula_reference_text(ref_row, ref_col)
                    if is_formula_text(result):
                        result += ")" * max(0, result.count("(") - result.count(")"))
                    self.current_row, self.current_col = original_current
                elif effective_reference is not None:
                    self.current_row, self.current_col = original_current
                return result
            if key == 27:
                curses.curs_set(0)
                if formula_origin is not None or effective_reference is not None:
                    self.current_row, self.current_col = original_current
                return None
            if (formula_origin is not None or effective_reference is not None) and key in (curses.KEY_UP, curses.KEY_DOWN, curses.KEY_LEFT, curses.KEY_RIGHT):
                row_delta, col_delta = {
                    curses.KEY_UP: (-1, 0),
                    curses.KEY_DOWN: (1, 0),
                    curses.KEY_LEFT: (0, -1),
                    curses.KEY_RIGHT: (0, 1),
                }[key]
                if formula_origin is not None:
                    active_reference_context = self._formula_reference_context("".join(text), position) or inserted_ref is not None
                else:
                    active_reference_context = position == len(text) or inserted_ref is not None
                if active_reference_context:
                    ref_row = max(0, min(self.sheet.rows - 1, ref_row + row_delta))
                    ref_col = max(0, min(self.sheet.cols - 1, ref_col + col_delta))
                    ref_text = self._formula_reference_text(ref_row, ref_col)
                    text, position, inserted_ref = self._insert_or_replace_formula_reference(text, position, inserted_ref, ref_text)
                    continue
            if key in (curses.KEY_BACKSPACE, 127):
                if position > 0:
                    del text[position - 1]
                    position -= 1
                inserted_ref = None
                continue
            if key == curses.KEY_DC and position < len(text):
                del text[position]
                inserted_ref = None
                continue
            if key == curses.KEY_LEFT and position > 0:
                position -= 1
                inserted_ref = None
                continue
            if key == curses.KEY_RIGHT and position < len(text):
                position += 1
                inserted_ref = None
                continue
            if 32 <= key <= 126:
                text.insert(position, chr(key))
                position += 1
                inserted_ref = None

    def _formula_reference_context(self, text: str, position: int) -> bool:
        if not is_formula_text(text) or position != len(text):
            return False
        stripped = text.rstrip()
        if not stripped:
            return False
        return stripped[-1] in "=(:,+-*/^<>"

    def _formula_reference_text(self, row: int, col: int) -> str:
        return f"{column_label(col)}{row + 1}"

    def _formula_context(self, text: str, position: int) -> tuple[str | None, int]:
        prefix = text[:position]
        stack: list[tuple[str, int]] = []
        for index, char in enumerate(prefix):
            if char == "(":
                probe = index - 1
                while probe >= 0 and prefix[probe].isspace():
                    probe -= 1
                end = probe + 1
                while probe >= 0 and (prefix[probe].isalpha() or prefix[probe] == "_"):
                    probe -= 1
                function_name = prefix[probe + 1 : end].upper()
                stack.append((function_name, index))
            elif char == ")" and stack:
                stack.pop()
        if not stack:
            return None, 0
        function_name, open_index = stack[-1]
        argument_count = 1 + prefix[open_index + 1 :].count(",")
        return function_name or None, argument_count

    def _formula_pointing_text(self, text: str, ref_row: int, ref_col: int) -> str:
        match = FORMULA_REF_PATTERN.search(text.rstrip())
        if match:
            return match.group(1)
        return self._formula_reference_text(ref_row, ref_col)

    def _formula_prompt_hints(self, text: str, position: int, ref_row: int, ref_col: int) -> list[str]:
        function_name, argument_count = self._formula_context(text, position)
        if function_name:
            signature = FORMULA_SIGNATURES.get(function_name, f"{function_name}(...)")
            arguments = FORMULA_ARGUMENT_NAMES.get(function_name, [])
            if arguments:
                argument_index = min(max(0, argument_count - 1), len(arguments) - 1)
                signature_line = f" Formula: {signature}   Arg {argument_count}: {arguments[argument_index]} "
            else:
                signature_line = f" Formula: {signature} "
        else:
            signature_line = " Formula: arrows point to cells, : starts a range, Enter accepts "
        pointing_text = self._formula_pointing_text(text, ref_row, ref_col)
        pointer_line = f" Pointing: {pointing_text} "
        return [signature_line, pointer_line]

    def _formula_prompt_error(self, text: str) -> str | None:
        if not is_formula_text(text):
            return None
        expression = text[1:].strip()
        if not expression:
            return None
        try:
            self.evaluator.evaluate_expression(expression, set())
        except FormulaError as exc:
            return f" Error: {exc} "
        return None

    def _draw_formula_prompt_hints(self, start_y: int, width: int, lines: list[str]) -> None:
        if start_y < 0 or width <= 1:
            return
        for index, line in enumerate(lines):
            y = start_y + index
            if y < 0:
                continue
            self.stdscr.addnstr(y, 0, (" " * (width - 1)), width - 1, self._bar_attr())
            self.stdscr.addnstr(y, 0, line.ljust(width - 1), width - 1, self._bar_attr())

    def _insert_or_replace_formula_reference(
        self,
        text: list[str],
        position: int,
        inserted_ref: tuple[int, int] | None,
        ref_text: str,
    ) -> tuple[list[str], int, tuple[int, int]]:
        if inserted_ref is not None and inserted_ref[1] == len(text) and position == len(text):
            start, end = inserted_ref
            text = text[:start] + list(ref_text)
            position = start + len(ref_text)
            return text, position, (start, position)
        start = position
        text = text[:position] + list(ref_text) + text[position:]
        position += len(ref_text)
        return text, position, (start, position)

    def render_text_snapshot(self) -> str:
        return self.render_delimited_snapshot("|")

    def render_fixed_width_snapshot(self) -> list[str]:
        max_row = -1
        max_col = -1
        for row, col, raw in self.sheet.iter_cells():
            if raw:
                max_row = max(max_row, row)
                max_col = max(max_col, col)
        if max_row < 0 or max_col < 0:
            return ["[No populated cells]"]

        values: list[list[str]] = []
        widths: list[int] = []
        for col in range(max_col + 1):
            widths.append(max(3, self.sheet.get_column_width(col) - 1))

        for row in range(max_row + 1):
            row_values = []
            for col in range(max_col + 1):
                display = self._display_value(row, col)
                width = widths[col]
                row_values.append(display[:width])
            values.append(row_values)

        for col in range(max_col + 1):
            header_width = len(column_label(col))
            content_width = max((len(row_values[col]) for row_values in values), default=0)
            widths[col] = max(3, min(max(widths[col], header_width, content_width), 32))

        header = " | ".join(column_label(col).center(widths[col]) for col in range(max_col + 1))
        separator = "-+-".join("-" * widths[col] for col in range(max_col + 1))
        lines = [header, separator]
        for row_index, row_values in enumerate(values):
            padded = []
            for col, value in enumerate(row_values):
                align = self.sheet.get_alignment(row_index, col) or "left"
                if align == "right":
                    padded.append(value.rjust(widths[col]))
                elif align == "centre":
                    padded.append(value.center(widths[col]))
                else:
                    padded.append(value.ljust(widths[col]))
            lines.append(" | ".join(padded).rstrip())
        return lines

    def render_delimited_snapshot(self, delimiter: str = "\t") -> str:
        max_row = -1
        max_col = -1
        for row, col, raw in self.sheet.iter_cells():
            if raw:
                max_row = max(max_row, row)
                max_col = max(max_col, col)
        if max_row < 0 or max_col < 0:
            return ""
        lines = []
        for row in range(max_row + 1):
            row_values = [self._display_value(row, col) for col in range(max_col + 1)]
            lines.append(delimiter.join(row_values).rstrip())
        return "\n".join(lines) + ("\n" if lines else "")

    def _title_line(self, width: int) -> str:
        target = str(self.path) if self.path else "[unsaved]"
        current_width = self.sheet.get_column_width(self.current_col)
        dirty_flag = "*" if self.dirty else ""
        title = f" {APP_NAME}{dirty_flag}  {target}  defw={self.sheet.column_width}  {column_label(self.current_col)}w={current_width}  build={build_stamp()} "
        return title.ljust(width - 1)

    def _options_line(self, width: int) -> str:
        current_format = self.sheet.get_format(self.current_row, self.current_col) or "text"
        current_styles = ",".join(sorted(self.sheet.get_text_styles(self.current_row, self.current_col))) or "plain"
        current_align = self.sheet.get_alignment(self.current_row, self.current_col) or "left"
        text = (
            " Formats: /F text /F currency /F fixed /F percent /F int /F sci"
            f"  |  Theme: /V [Enter]=cycle /V white|cyan|yellow|magenta|blue"
            f"  |  Justify: /J l /J c /J r"
            f"  |  Width: /G width 14  or  /G width {column_label(self.current_col)} 18"
            f"  |  Cell format={current_format} styles={current_styles} align={current_align} "
        )
        return text[: width - 1].ljust(width - 1)

    def _formula_bar(self, width: int) -> str:
        ref = f"{column_label(self.current_col)}{self.current_row + 1}"
        raw = self.sheet.get_raw(self.current_row, self.current_col)
        error_text = ""
        try:
            value = self.evaluator.display_value(self.current_row, self.current_col)
        except FormulaError as exc:
            value = f"#ERR {exc}"
            error_text = f" error={exc}"
        value = self._apply_format(value, self.sheet.get_format(self.current_row, self.current_col))
        flags = []
        if self.sheet.is_protected(self.current_row, self.current_col):
            flags.append("PROT")
        cell_format = self.sheet.get_format(self.current_row, self.current_col)
        if cell_format:
            flags.append(cell_format.upper())
        cell_text_styles = sorted(self.sheet.get_text_styles(self.current_row, self.current_col))
        if cell_text_styles:
            flags.extend(style.upper() for style in cell_text_styles)
        cell_background = self.sheet.get_background(self.current_row, self.current_col)
        if cell_background:
            flags.append(f"BG={cell_background.upper()}")
        cell_border = self.sheet.get_border(self.current_row, self.current_col)
        if cell_border:
            flags.append(f"BORDER={cell_border.upper()}")
        meta = f"[{' '.join(flags)}]" if flags else ""
        selection = f" sel={self._selection_label()}" if self.selection_range else ""
        text = f" {ref} raw={raw or ' '} value={value or ' '} {meta}{selection}{error_text}"
        return text[: width - 1].ljust(width - 1)

    def _top_formula_line(self, width: int) -> str:
        ref = f"{column_label(self.current_col)}{self.current_row + 1}"
        raw = self.sheet.get_raw(self.current_row, self.current_col) or ""
        if is_formula_text(raw):
            try:
                value = self.evaluator.display_value(self.current_row, self.current_col)
                error_suffix = ""
            except FormulaError as exc:
                value = f"#ERR {exc}"
                error_suffix = f"   error={exc}"
            function_name, argument_count = self._formula_context(raw, len(raw))
            references = re.findall(r"\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?", raw)
            refs_text = f"   refs={','.join(references[:4])}" if references else ""
            if function_name:
                signature = FORMULA_SIGNATURES.get(function_name, f"{function_name}(...)")
                arguments = FORMULA_ARGUMENT_NAMES.get(function_name, [])
                if arguments:
                    argument_index = min(max(0, argument_count - 1), len(arguments) - 1)
                    arg_text = f"   arg {argument_count}: {arguments[argument_index]}"
                else:
                    arg_text = ""
                text = f" Fx {ref}: {raw}   => {value}   {signature}{arg_text}{refs_text}{error_suffix}"
            else:
                text = f" Fx {ref}: {raw}   => {value}{refs_text}{error_suffix}"
        else:
            display = self._display_value(self.current_row, self.current_col)
            style = self.sheet.get_format(self.current_row, self.current_col) or "text"
            styles = ",".join(sorted(self.sheet.get_text_styles(self.current_row, self.current_col))) or "plain"
            align = self.sheet.get_alignment(self.current_row, self.current_col) or "left"
            text = f" Cell {ref}: {raw or display or '(empty)'}   format={style} styles={styles} align={align}"
        return text[: width - 1].ljust(width - 1)

    def _display_value(self, row: int, col: int) -> str:
        try:
            text = self.evaluator.display_value(row, col)
        except FormulaError as exc:
            return f"#ERR {exc}"
        return self._apply_format(text, self.sheet.get_format(row, col))

    def _cell_text(self, row: int, col: int, width: int) -> str:
        if self.raw_sheet_view:
            text = self.sheet.get_raw(row, col)
        else:
            text = self._display_value(row, col)
        text = (text or "")[:width]
        align = self.sheet.get_alignment(row, col)
        if align == "right":
            rendered = text.rjust(width)
        elif align == "center":
            rendered = text.center(width)
        else:
            rendered = text.ljust(width)
        return self._decorate_cell_border_text(row, col, rendered, width)

    def _decorate_cell_border_text(self, row: int, col: int, text: str, width: int) -> str:
        border = self.sheet.get_border(row, col)
        if border not in {"outline", "all"} or width <= 0:
            return text
        return text

    def _spill_width(self, row: int, column_index: int, visible_columns: list[tuple[int, int, int]]) -> tuple[int, int]:
        col, _x, col_width = visible_columns[column_index]
        raw = self.sheet.get_raw(row, col)
        if not raw:
            return col_width - 1, column_index
        align = self.sheet.get_alignment(row, col)
        if align in {"right", "center"}:
            return col_width - 1, column_index
        text = self._display_value(row, col)
        if len(text) <= col_width - 1:
            return col_width - 1, column_index
        total_width = col_width - 1
        spill_to_index = column_index
        for next_index in range(column_index + 1, len(visible_columns)):
            next_col, _next_x, next_width = visible_columns[next_index]
            if self.sheet.get_raw(row, next_col):
                break
            if self.sheet.get_background(row, next_col) or self.sheet.get_row_background(row):
                break
            if self.sheet.is_protected(row, next_col):
                break
            if self._cell_in_selection(row, next_col):
                break
            if (row, next_col) == (self.current_row, self.current_col):
                break
            total_width += next_width - 1
            spill_to_index = next_index
            if len(text) <= total_width:
                break
        return total_width, spill_to_index

    def _cell_attr(self, row: int, col: int) -> int:
        attr = curses.A_NORMAL
        if not self.colors_ready:
            return attr
        raw = self.sheet.get_raw(row, col)
        style = self.sheet.get_format(row, col)
        text_style_attr = self._text_style_attr(self.sheet.get_text_styles(row, col), self.sheet.get_border(row, col))
        background_name = self.sheet.get_background(row, col) or self.sheet.get_row_background(row)
        if self.sheet.is_protected(row, col):
            protected_fg = self._named_color(self.sheet.protected_foreground_color, curses.COLOR_BLACK)
            protected_bg = self._named_color(self.sheet.protected_background_color, curses.COLOR_WHITE)
            pair_number = self._ensure_color_pair(protected_fg, protected_bg)
            if pair_number is not None:
                attr |= curses.color_pair(pair_number) | text_style_attr
                if is_formula_text(raw):
                    attr |= curses.A_BOLD
                return attr
        if background_name:
            background_color = self._named_color(background_name, -1)
            foreground_color = self._background_foreground_color(background_name, is_formula_text(raw), row, col, style)
            pair_number = self._ensure_color_pair(foreground_color, background_color)
            if pair_number is not None:
                attr |= curses.color_pair(pair_number) | text_style_attr
                if is_formula_text(raw):
                    attr |= curses.A_BOLD
                return attr
        if style == "negative" and self._cell_numeric_value(row, col) is not None and self._cell_numeric_value(row, col) < 0:
            return attr | curses.color_pair(COLOR_PAIR_NEGATIVE) | text_style_attr
        if is_formula_text(raw) and self.sheet.formula_coloration:
            return attr | curses.color_pair(COLOR_PAIR_FORMULA) | curses.A_BOLD | text_style_attr
        return attr | curses.color_pair(COLOR_PAIR_TEXT) | text_style_attr

    def _selection_cell_attr(self, row: int, col: int) -> int:
        attr = curses.A_NORMAL
        if not self.colors_ready:
            return attr | curses.A_REVERSE
        raw = self.sheet.get_raw(row, col)
        text_style_attr = self._text_style_attr(self.sheet.get_text_styles(row, col), self.sheet.get_border(row, col))
        selection_background = self._selection_background_color()
        if is_formula_text(raw) and self.sheet.formula_coloration:
            pair_number = self._ensure_color_pair(self._formula_foreground_color(), selection_background)
            if pair_number is not None:
                return attr | curses.color_pair(pair_number) | curses.A_BOLD | text_style_attr
            return attr | curses.color_pair(COLOR_PAIR_SELECTION) | curses.A_BOLD | text_style_attr
        foreground = self._selection_foreground_color()
        pair_number = self._ensure_color_pair(foreground, selection_background)
        if pair_number is not None:
            return attr | curses.color_pair(pair_number) | text_style_attr
        return attr | curses.color_pair(COLOR_PAIR_SELECTION) | text_style_attr

    def _hex_luminance(self, background_name: str) -> float | None:
        if not HEX_COLOR_RE.match(background_name):
            return None
        normalized = background_name.upper()
        if not normalized.startswith("#"):
            normalized = f"#{normalized}"
        red = int(normalized[1:3], 16)
        green = int(normalized[3:5], 16)
        blue = int(normalized[5:7], 16)
        return (0.299 * red) + (0.587 * green) + (0.114 * blue)

    def _active_cell_attr(self, row: int, col: int) -> int:
        attr = self._selection_cell_attr(row, col)
        return attr | curses.A_BOLD

    def _text_style_attr(self, styles: set[str], border: str = "") -> int:
        attr = curses.A_NORMAL
        if "bold" in styles:
            attr |= curses.A_BOLD
        if "underline" in styles:
            attr |= curses.A_UNDERLINE
        if "italic" in styles:
            attr |= getattr(curses, "A_ITALIC", curses.A_DIM)
        if border == "underline":
            attr |= curses.A_UNDERLINE
        if border in {"outline", "all"}:
            attr |= curses.A_BOLD
        return attr

    def _selection_foreground_color(self) -> int:
        if self.sheet.active_cell_color in {"white", "yellow", "lightblue", "lightgrey"}:
            return curses.COLOR_BLACK
        return curses.COLOR_WHITE

    def _formula_foreground_color(self) -> int:
        return self._named_color(self.sheet.formula_foreground_color, curses.COLOR_GREEN)

    def _background_foreground_color(self, background_name: str, is_formula: bool, row: int, col: int, style: str | None) -> int:
        numeric_value = self._cell_numeric_value(row, col)
        luminance = self._hex_luminance(background_name)
        if style == "negative" and numeric_value is not None and numeric_value < 0:
            return curses.COLOR_RED
        if is_formula and self.sheet.formula_coloration:
            if luminance is not None and luminance >= 160:
                return self._named_color(self.sheet.formula_foreground_color, curses.COLOR_GREEN)
            return self._formula_foreground_color()
        if luminance is not None:
            return curses.COLOR_BLACK if luminance >= 160 else curses.COLOR_WHITE
        if background_name in {"cyan", "green", "lightblue", "lightgrey", "palepink", "white", "yellow"}:
            return curses.COLOR_BLACK
        return curses.COLOR_WHITE

    def _selection_background_color(self) -> int:
        name = self.sheet.active_cell_color
        if name == "yellow":
            return curses.COLOR_YELLOW
        if name == "pink":
            return self._custom_named_color(CUSTOM_PINK_COLOR_ID, 1000, 500, 700) or curses.COLOR_MAGENTA
        if name == "orange":
            return self._custom_orange_color() or curses.COLOR_YELLOW
        if name == "white":
            return curses.COLOR_WHITE
        if name == "lightblue":
            return self._custom_named_color(CUSTOM_LIGHTBLUE_COLOR_ID, 500, 700, 1000) or curses.COLOR_CYAN
        if name == "cornflower":
            return self._custom_named_color(CUSTOM_CORNFLOWER_COLOR_ID, 392, 584, 929) or curses.COLOR_BLUE
        if name == "lightgrey":
            return self._custom_named_color(CUSTOM_LIGHTGREY_COLOR_ID, 800, 800, 800) or curses.COLOR_WHITE
        return self._custom_orange_color() or curses.COLOR_YELLOW

    def _cell_in_selection(self, row: int, col: int) -> bool:
        if self.selection_range is None:
            return False
        row_lo, col_lo, row_hi, col_hi = self.selection_range
        return row_lo <= row <= row_hi and col_lo <= col <= col_hi

    def _cell_numeric_value(self, row: int, col: int) -> float | None:
        raw = self.sheet.get_raw(row, col)
        if not raw:
            return None
        try:
            value = self.evaluator.evaluate_cell(row, col, set()) if is_formula_text(raw) else raw
        except FormulaError:
            return None
        try:
            return float(str(value).replace(",", ""))
        except ValueError:
            return None

    def _ensure_color_pair(self, foreground: int, background: int) -> int | None:
        key = (foreground, background)
        existing = self.dynamic_color_pairs.get(key)
        if existing is not None:
            return existing
        if self.next_dynamic_pair >= curses.COLOR_PAIRS:
            return None
        curses.init_pair(self.next_dynamic_pair, foreground, background)
        pair_number = self.next_dynamic_pair
        self.dynamic_color_pairs[key] = pair_number
        self.next_dynamic_pair += 1
        return pair_number

    def _grid_attr(self) -> int:
        attr = curses.A_DIM
        if self.colors_ready:
            attr |= curses.color_pair(COLOR_PAIR_GRID)
        return attr

    def _row_header_attr(self) -> int:
        attr = curses.A_BOLD
        if self.colors_ready:
            attr |= curses.color_pair(COLOR_PAIR_ROW_HEADER)
        return attr

    def _border_attr(self) -> int:
        attr = curses.A_BOLD
        if self.colors_ready:
            attr |= curses.color_pair(COLOR_PAIR_HEADER)
        return attr

    def _bar_attr(self, bold: bool = False) -> int:
        attr = curses.A_NORMAL
        if bold:
            attr |= curses.A_BOLD
        if self.colors_ready:
            attr |= curses.color_pair(COLOR_PAIR_BAR)
        return attr

    def _prompt_attr(self, bold: bool = False) -> int:
        attr = curses.A_NORMAL
        if bold:
            attr |= curses.A_BOLD
        if self.colors_ready:
            attr |= curses.color_pair(COLOR_PAIR_TEXT)
        return attr

    def _file_browser_entry_attr(self, kind: str, selected: bool = False) -> int:
        foreground = curses.COLOR_WHITE
        if kind == "dir":
            foreground = curses.COLOR_BLUE
        elif kind == "file":
            foreground = curses.COLOR_GREEN
        elif kind == "unsupported":
            foreground = curses.COLOR_RED
        if selected:
            background = self._selection_background_color() if self.colors_ready else -1
            pair_number = self._ensure_color_pair(foreground, background) if self.colors_ready else None
            attr = curses.A_BOLD
            if pair_number is not None:
                attr |= curses.color_pair(pair_number)
            else:
                attr |= self._menu_selected_attr()
            return attr
        attr = curses.A_BOLD if kind in {"dir", "file"} else curses.A_NORMAL
        if self.colors_ready:
            pair_number = self._ensure_color_pair(foreground, -1)
            if pair_number is not None:
                attr |= curses.color_pair(pair_number)
            else:
                attr |= curses.color_pair(COLOR_PAIR_BAR)
        return attr

    def _file_browser_sort_key(self, path: Path, kind: str, sort_mode: str) -> tuple[object, ...]:
        if sort_mode == "time":
            try:
                mtime = path.stat().st_mtime
            except OSError:
                mtime = 0.0
            return (0 if kind == "dir" else 1, -mtime, path.name.lower())
        if sort_mode == "type":
            suffix = "" if kind == "dir" else path.suffix.lower()
            return (0 if kind == "dir" else 1, suffix, path.name.lower())
        return (0 if kind == "dir" else 1, path.name.lower())

    def _browser_item_kind(self, path: Path, suffixes: set[str] | None) -> str:
        if path.is_dir():
            return "dir"
        if not suffixes or path.suffix.lower() in suffixes:
            return "file"
        return "unsupported"

    def _format_file_size(self, path: Path, kind: str) -> str:
        if kind == "dir":
            return "--"
        try:
            size = path.stat().st_size
        except OSError:
            return "?"
        units = ["B", "K", "M", "G"]
        value = float(size)
        unit = units[0]
        for candidate in units:
            unit = candidate
            if value < 1024.0 or candidate == units[-1]:
                break
            value /= 1024.0
        if unit == "B":
            return f"{int(value)}B"
        return f"{value:.1f}{unit}"

    def _format_file_mtime(self, path: Path) -> str:
        try:
            stamp = datetime.fromtimestamp(path.stat().st_mtime)
        except OSError:
            return "--/--/-- --:--"
        date_text = stamp.strftime("%Y-%m-%d")
        display_date = format_date_text(date_text, self.sheet.date_format)
        return f"{display_date} {stamp:%H:%M}"

    def _browser_icon(self, kind: str, path: Path) -> str:
        if kind == "dir":
            return ""
        if kind == "unsupported":
            return ""
        suffix = path.suffix.lower()
        if suffix == ".tss":
            return "󰈙"
        if suffix in {".csv", ".tsv"}:
            return "󰈛"
        return ""

    def _rename_browser_item(self, path: Path) -> Path | None:
        if not path.exists():
            self.message = f"Path not found: {path}"
            return None
        typed = self.prompt("Rename to: ", path.name)
        if typed is None or not typed.strip():
            return None
        new_name = Path(typed).name
        if not new_name:
            return None
        target = path.with_name(new_name)
        if target.exists():
            self.message = f"Rename error: {target.name} already exists"
            return None
        try:
            path.rename(target)
        except OSError as exc:
            self.message = f"Rename error: {exc}"
            return None
        self.message = f"Renamed to {target.name}"
        return target

    def _delete_browser_item(self, path: Path) -> bool:
        if not path.exists():
            self.message = f"Path not found: {path}"
            return False
        label = f"{path.name}/" if path.is_dir() else path.name
        choice = self._choose_from_menu(
            "Delete",
            ["delete", "cancel"],
            default_option="cancel",
            descriptions={
                "delete": f"Delete {label}",
                "cancel": "Cancel deletion.",
            },
            footer_hint=" arrows/Enter/Esc ",
        )
        if choice != "delete":
            self.message = "Delete cancelled"
            return False
        try:
            if path.is_dir():
                shutil.rmtree(path)
            else:
                path.unlink()
        except OSError as exc:
            self.message = f"Delete error: {exc}"
            return False
        self.message = f"Deleted {label}"
        return True

    def _help_attr(self) -> int:
        attr = curses.A_DIM
        if self.colors_ready:
            attr |= curses.color_pair(COLOR_PAIR_TEXT)
        return attr

    def _visible_columns(self, total_width: int, row_header_width: int) -> list[tuple[int, int, int]]:
        columns: list[tuple[int, int, int]] = []
        x = row_header_width
        col = self.col_offset
        while col < self.sheet.cols and x < total_width - 1:
            if self.sheet.is_col_hidden(col):
                col += 1
                continue
            col_width = self.sheet.get_column_width(col)
            if x + col_width > total_width:
                break
            columns.append((col, x, col_width))
            x += col_width
            col += 1
        return columns

    def _visible_rows(self, grid_height: int) -> list[int]:
        rows: list[int] = []
        row = self.row_offset
        while row < self.sheet.rows and len(rows) < grid_height:
            if not self.sheet.is_row_hidden(row):
                rows.append(row)
            row += 1
        return rows

    def _grid_layout(self, height: int, width: int) -> tuple[int, int, int, list[tuple[int, int, int]]]:
        row_header_width = 6
        bottom_bars = 1
        top_grid_row = 4
        grid_height = max(3, height - top_grid_row - bottom_bars)
        visible_columns = self._visible_columns(width, row_header_width)
        return top_grid_row, grid_height, row_header_width, visible_columns

    def _cell_from_screen(self, y: int, x: int) -> tuple[int, int] | None:
        height, width = self.stdscr.getmaxyx()
        top_grid_row, grid_height, row_header_width, visible_columns = self._grid_layout(height, width)
        if y < top_grid_row or y >= top_grid_row + grid_height:
            return None
        display_lines = self._display_lines(grid_height, visible_columns)
        screen_index = y - top_grid_row
        if not (0 <= screen_index < len(display_lines)):
            return None
        kind, row, _edge = display_lines[screen_index]
        if kind != "row":
            return None
        if x < row_header_width:
            return None
        for col, col_x, col_width in visible_columns:
            if col_x <= x < col_x + col_width - 1:
                return row, col
        return None

    def _display_lines(self, grid_height: int, visible_columns: list[tuple[int, int, int]]) -> list[tuple[str, int, str | None]]:
        lines: list[tuple[str, int, str | None]] = []
        row = self.row_offset
        while row < self.sheet.rows and len(lines) < grid_height:
            if self.sheet.is_row_hidden(row):
                row += 1
                continue
            if self._row_has_top_border(row, visible_columns) and len(lines) < grid_height:
                lines.append(("sep", row, "top"))
            if len(lines) < grid_height:
                lines.append(("row", row, None))
            if self._row_has_bottom_border(row, visible_columns) and len(lines) < grid_height:
                lines.append(("sep", row, "bottom"))
            row += 1
        return lines

    def _row_has_top_border(self, row: int, visible_columns: list[tuple[int, int, int]]) -> bool:
        for col, _x, _width in visible_columns:
            border = self.sheet.get_border(row, col)
            if border in {"all", "outline"} and not (row > 0 and self.sheet.get_border(row - 1, col) in {"all", "outline"}):
                return True
        return False

    def _row_has_bottom_border(self, row: int, visible_columns: list[tuple[int, int, int]]) -> bool:
        for col, _x, _width in visible_columns:
            border = self.sheet.get_border(row, col)
            if border in {"all", "outline"} and not (row < self.sheet.rows - 1 and self.sheet.get_border(row + 1, col) in {"all", "outline"}):
                return True
        return False

    def _first_visible_row(self) -> int:
        for row in range(self.sheet.rows):
            if not self.sheet.is_row_hidden(row):
                return row
        return 0

    def _first_visible_col(self) -> int:
        for col in range(self.sheet.cols):
            if not self.sheet.is_col_hidden(col):
                return col
        return 0

    def _step_visible_row(self, start: int, delta: int) -> int:
        if delta == 0:
            return start if not self.sheet.is_row_hidden(start) else self._first_visible_row()
        direction = 1 if delta > 0 else -1
        row = start
        remaining = abs(delta)
        while remaining > 0:
            row = max(0, min(self.sheet.rows - 1, row + direction))
            while 0 <= row < self.sheet.rows and self.sheet.is_row_hidden(row):
                next_row = row + direction
                if not (0 <= next_row < self.sheet.rows):
                    break
                row = next_row
            remaining -= 1
        return row

    def _step_visible_col(self, start: int, delta: int) -> int:
        if delta == 0:
            return start if not self.sheet.is_col_hidden(start) else self._first_visible_col()
        direction = 1 if delta > 0 else -1
        col = start
        remaining = abs(delta)
        while remaining > 0:
            col = max(0, min(self.sheet.cols - 1, col + direction))
            while 0 <= col < self.sheet.cols and self.sheet.is_col_hidden(col):
                next_col = col + direction
                if not (0 <= next_col < self.sheet.cols):
                    break
                col = next_col
            remaining -= 1
        return col

    def _normalize_range(self, start: tuple[int, int], end: tuple[int, int]) -> tuple[int, int, int, int]:
        start_row, start_col = start
        end_row, end_col = end
        row_lo, row_hi = sorted((start_row, end_row))
        col_lo, col_hi = sorted((start_col, end_col))
        return row_lo, col_lo, row_hi, col_hi

    def _selection_label(self) -> str:
        if self.selection_range is None:
            return ""
        row_lo, col_lo, row_hi, col_hi = self.selection_range
        if col_lo == 0 and col_hi == self.sheet.cols - 1 and row_lo == row_hi:
            return f"{row_lo + 1}:{row_hi + 1}"
        if row_lo == 0 and row_hi == self.sheet.rows - 1 and col_lo == col_hi:
            col_text = column_label(col_lo)
            return f"{col_text}:{col_text}"
        return self._range_label(row_lo, col_lo, row_hi, col_hi)

    def _target_range(self, spec: str | None) -> tuple[int, int, int, int]:
        if spec and spec.strip() and spec.strip() != ".":
            return self._parse_range_spec(spec)
        if self.selection_range is not None:
            return self.selection_range
        return self.current_row, self.current_col, self.current_row, self.current_col

    def _resolve_named_spec(self, spec: str) -> str:
        resolved = self.sheet.get_named_range(spec)
        return resolved or spec

    def _parse_range_spec(self, spec: str) -> tuple[int, int, int, int]:
        return parse_range_spec(self._resolve_named_spec(spec), self.current_row, self.current_col, self.sheet.rows, self.sheet.cols)

    def _choose_from_menu(
        self,
        title: str,
        options: list[str],
        default_option: str | None = None,
        descriptions: dict[str, str] | None = None,
        toggle_key: int | None = None,
        toggle_value: str | None = None,
        footer_hint: str | None = None,
    ) -> str | None:
        height, width = self.stdscr.getmaxyx()
        ranked_options = list(options)
        selected_option = default_option if default_option in ranked_options else ranked_options[0]
        typed = ""
        last_normalized = ""
        curses.curs_set(0)
        while True:
            self.draw()
            normalized = typed.lstrip("/").strip().lower()
            if normalized:
                alias_target = ALIASES.get(normalized)
                exact = [option for option in options if option.lower() == normalized]
                alias_match = [alias_target] if alias_target and alias_target in options else []
                starts = [
                    option
                    for option in options
                    if option.lower().startswith(normalized) and option not in exact and option not in alias_match
                ]
                contains = [
                    option
                    for option in options
                    if normalized in option.lower() and option not in exact and option not in alias_match and option not in starts
                ]
                ranked_options = []
                for bucket in (exact, alias_match, starts, contains):
                    for option in bucket:
                        if option not in ranked_options:
                            ranked_options.append(option)
                if not ranked_options:
                    ranked_options = list(options)
                if normalized != last_normalized or selected_option not in ranked_options:
                    selected_option = ranked_options[0]
            else:
                ranked_options = list(options)
                if selected_option not in ranked_options:
                    selected_option = ranked_options[0]
            last_normalized = normalized
            selected = ranked_options.index(selected_option)
            query = f" [{normalized}]" if normalized else ""
            title_text = f"{title}{query}: "
            hint = footer_hint or " arrows/type/Enter/Esc "
            rows: list[list[tuple[int, str]]] = [[]]
            row_widths = [len(title_text)]
            for index, option in enumerate(ranked_options):
                chip = f" {option} "
                needed = len(chip) + (1 if rows[-1] else 0)
                if row_widths[-1] + needed > width - 1 and rows[-1]:
                    rows.append([])
                    row_widths.append(0)
                rows[-1].append((index, chip))
                row_widths[-1] += len(chip) + (1 if len(rows[-1]) > 1 else 0)
            if row_widths[-1] + len(hint) + (1 if rows[-1] else 0) > width - 1 and rows[-1]:
                rows.append([])
                row_widths.append(0)
            rows[-1].append((-1, hint))
            menu_rows = len(rows)
            description = descriptions.get(ranked_options[selected], "") if descriptions else ""
            start_y = max(0, height - menu_rows - (1 if description else 0))
            for offset in range(menu_rows):
                y = start_y + offset
                if 0 <= y < height and width > 1:
                    self.stdscr.addnstr(y, 0, (" " * (width - 1)), width - 1, self._bar_attr(bold=True))
            if description:
                if 0 <= start_y < height and width > 1:
                    self.stdscr.addnstr(start_y, 0, description.ljust(width - 1), width - 1, self._help_attr())
            for offset, row_items in enumerate(rows):
                y = start_y + offset + (1 if description else 0)
                if not (0 <= y < height):
                    continue
                x = 0
                if offset == 0:
                    available = width - 1 - x
                    if available > 0:
                        clipped = title_text[:available]
                        self.stdscr.addnstr(y, x, clipped, len(clipped), self._bar_attr(bold=True))
                        x += len(clipped)
                for item_index, chip in row_items:
                    if x and x < width - 1:
                        self.stdscr.addch(y, x, ord(" "), self._bar_attr())
                        x += 1
                    attr = self._help_attr() if item_index == -1 else (self._menu_selected_attr() if ranked_options[item_index] == selected_option else self._bar_attr())
                    available = width - 1 - x
                    if available > 0:
                        clipped = chip[:available]
                        self.stdscr.addnstr(y, x, clipped, len(clipped), attr)
                        x += len(clipped)
                    else:
                        break
            self.stdscr.refresh()
            key = self.stdscr.getch()
            if key in (27,):
                return None
            if toggle_key is not None and key == toggle_key:
                return toggle_value
            if key in (10, 13):
                return selected_option
            if key == curses.KEY_LEFT:
                selected = (selected - 1) % len(ranked_options)
                selected_option = ranked_options[selected]
                typed = ""
                continue
            if key == curses.KEY_RIGHT:
                selected = (selected + 1) % len(ranked_options)
                selected_option = ranked_options[selected]
                typed = ""
                continue
            if key in (curses.KEY_BACKSPACE, 127):
                typed = typed[:-1]
                continue
            if 32 <= key <= 126:
                typed += chr(key).lower()
                continue

    def _browse_for_file(
        self,
        title: str,
        start_path: Path,
        suffixes: set[str] | None = None,
    ) -> Path | None:
        current_dir = start_path.expanduser().resolve()
        if current_dir.is_file():
            current_dir = current_dir.parent
        if not current_dir.exists():
            current_dir = Path.home()
        selected = 0
        offset = 0
        show_hidden = False
        filter_text = ""
        sort_mode = "name"
        curses.curs_set(0)
        while True:
            try:
                entries = list(current_dir.iterdir())
            except OSError as exc:
                self.message = f"Browse error: {exc}"
                return None
            visible_entries: list[Path] = [current_dir.parent] + entries if current_dir.parent != current_dir else entries
            filtered: list[Path] = []
            kinds: list[str] = []
            for index, entry in enumerate(visible_entries):
                if index == 0 and current_dir.parent != current_dir:
                    filtered.append(entry)
                    kinds.append("dir")
                    continue
                if not show_hidden and entry.name.startswith("."):
                    continue
                if filter_text and filter_text.lower() not in entry.name.lower():
                    continue
                filtered.append(entry)
                kinds.append(self._browser_item_kind(entry, suffixes))
            if current_dir.parent != current_dir and filtered:
                parent = filtered[0]
                parent_kind = kinds[0]
                paired = list(zip(filtered[1:], kinds[1:]))
                paired.sort(key=lambda item: self._file_browser_sort_key(item[0], item[1], sort_mode))
                filtered = [parent] + [item[0] for item in paired]
                kinds = [parent_kind] + [item[1] for item in paired]
            else:
                paired = list(zip(filtered, kinds))
                paired.sort(key=lambda item: self._file_browser_sort_key(item[0], item[1], sort_mode))
                filtered = [item[0] for item in paired]
                kinds = [item[1] for item in paired]
            if not filtered:
                filtered = [current_dir]
                kinds = ["empty"]
                selected = 0
            selected = max(0, min(selected, len(filtered) - 1))
            height, width = self.stdscr.getmaxyx()
            self.stdscr.erase()
            self.stdscr.addnstr(0, 0, f" {title} ".ljust(width - 1), width - 1, self._bar_attr(bold=True))
            self.stdscr.addnstr(1, 0, str(current_dir).ljust(width - 1), width - 1, self._help_attr())
            hint = (
                f" Up/Down select  Enter open/select  .. up  . hidden:{'on' if show_hidden else 'off'}"
                f"  s sort:{sort_mode}  p path  n mkdir  r rename  Del delete  type=filter  Esc cancel "
            )
            self.stdscr.addnstr(2, 0, hint.ljust(width - 1), width - 1, self._help_attr())
            filter_label = f" Filter: {filter_text or '(all)'} "
            self.stdscr.addnstr(3, 0, filter_label.ljust(width - 1), width - 1, self._help_attr())

            top = 5
            content_height = max(6, height - top - 2)
            list_width = max(28, min(width * 3 // 5, width - 24))
            detail_x = min(width - 2, list_width + 2)
            detail_width = max(12, width - detail_x - 1)
            visible_rows = max(1, content_height - 2)
            if selected < offset:
                offset = selected
            elif selected >= offset + visible_rows:
                offset = selected - visible_rows + 1

            header = f" Files [{sort_mode}] "
            self.stdscr.addnstr(top, 0, header.ljust(list_width - 1), list_width - 1, self._bar_attr(bold=True))
            self.stdscr.addch(top, list_width - 1, ord(" "), self._bar_attr())
            if detail_width > 0:
                self.stdscr.addnstr(top, detail_x, " Details ".ljust(detail_width), detail_width, self._bar_attr(bold=True))
            for y in range(top + 1, top + content_height):
                if list_width - 1 < width - 1:
                    self.stdscr.addch(y, list_width - 1, ord("|"), self._help_attr())

            name_width = max(10, list_width - 1 - 20)
            size_width = 8
            time_width = max(10, list_width - 1 - name_width - size_width - 4)
            list_header = f" {'Name'.ljust(name_width)} {'Size'.rjust(size_width)} {'Modified'.ljust(time_width)}"
            self.stdscr.addnstr(top + 1, 0, list_header.ljust(list_width - 1), list_width - 1, self._help_attr())

            for row_index in range(max(0, visible_rows - 1)):
                entry_index = offset + row_index
                if entry_index >= len(filtered):
                    break
                y = top + 2 + row_index
                entry = filtered[entry_index]
                kind = kinds[entry_index]
                attr = self._file_browser_entry_attr(kinds[entry_index], selected=(entry_index == selected))
                if kind == "dir" and entry == current_dir.parent and current_dir.parent != current_dir:
                    name = ".."
                else:
                    suffix = "/" if kind == "dir" else ""
                    name = f"{self._browser_icon(kind, entry)}  {entry.name}{suffix}"
                size = self._format_file_size(entry, kind)
                modified = self._format_file_mtime(entry)
                row_text = f" {name[:name_width].ljust(name_width)} {size.rjust(size_width)} {modified[:time_width].ljust(time_width)}"
                self.stdscr.addnstr(y, 0, row_text.ljust(list_width - 1), list_width - 1, attr)

            for row_index in range(max(0, visible_rows - 1), content_height - 2):
                y = top + 2 + row_index
                if y >= height - 1:
                    break
                self.stdscr.addnstr(y, 0, " ".ljust(list_width - 1), list_width - 1, self._bar_attr())

            if filtered and kinds[selected] != "empty":
                preview_path = filtered[selected]
                selected_kind = kinds[selected]
                details = [
                    f"Name: {preview_path.name or str(preview_path)}",
                    f"Type: {'folder' if selected_kind == 'dir' else preview_path.suffix.lower().lstrip('.') or 'file'}",
                    f"Path: {preview_path}",
                    f"Items: {len(filtered)} shown",
                ]
                if preview_path.exists() and not preview_path.is_dir():
                    try:
                        size = preview_path.stat().st_size
                        details.append(f"Size: {size} bytes")
                    except OSError:
                        pass
                if selected_kind == "file":
                    details.append("Status: loadable")
                elif selected_kind == "unsupported":
                    details.append("Status: unsupported file type")
                if preview_path == current_dir.parent and current_dir.parent != current_dir:
                    details[1] = "Type: parent folder"
            else:
                details = [
                    "No matching files.",
                    "Try clearing the filter",
                    "or toggle hidden files",
                    "with .",
                ]
            for index, line in enumerate(details):
                y = top + 1 + index
                if y >= top + content_height or detail_width <= 0:
                    break
                self.stdscr.addnstr(y, detail_x, line.ljust(detail_width), detail_width, self._help_attr())

            footer = " Enter=open/select   r=rename   Del=delete   s=sort   p=path   n=mkdir   .=hidden   type to filter   Backspace=clear   Esc=cancel "
            self.stdscr.addnstr(height - 1, 0, footer.ljust(width - 1), width - 1, self._bar_attr())
            self.stdscr.refresh()
            key = self.stdscr.getch()
            if key == 27:
                return None
            if key == ord("."):
                show_hidden = not show_hidden
                selected = 0
                offset = 0
                continue
            if key in (ord("s"), ord("S")):
                current_index = FILE_BROWSER_SORT_OPTIONS.index(sort_mode)
                sort_mode = FILE_BROWSER_SORT_OPTIONS[(current_index + 1) % len(FILE_BROWSER_SORT_OPTIONS)]
                selected = 0
                offset = 0
                continue
            if key in (ord("p"), ord("P")):
                target_text = self.prompt("Open path: ", str(current_dir))
                if target_text is None or not target_text.strip():
                    continue
                target = Path(target_text).expanduser()
                if target.is_dir():
                    current_dir = target.resolve()
                    filter_text = ""
                    selected = 0
                    offset = 0
                    continue
                if target.exists():
                    if suffixes and target.suffix.lower() not in suffixes:
                        self.message = f"Unsupported file type: {target.name}"
                        continue
                    return target.resolve()
                self.message = f"Path not found: {target}"
                continue
            if key in (ord("n"), ord("N")):
                folder_text = self.prompt("New folder: ", str(current_dir / "new-folder"))
                if folder_text is None or not folder_text.strip():
                    continue
                target = Path(folder_text).expanduser()
                if not target.is_absolute():
                    target = current_dir / target
                try:
                    target.mkdir(parents=True, exist_ok=False)
                except OSError as exc:
                    self.message = f"Create folder error: {exc}"
                    continue
                current_dir = target.parent.resolve()
                filter_text = ""
                selected = 0
                offset = 0
                self.message = f"Created folder {target.name}"
                continue
            if key in (ord("r"), ord("R")) and filtered and kinds[selected] != "empty":
                choice = filtered[selected]
                if choice == current_dir.parent and current_dir.parent != current_dir:
                    self.message = "Cannot rename parent entry"
                    continue
                renamed = self._rename_browser_item(choice)
                if renamed is not None:
                    selected = 0
                    offset = 0
                continue
            if key == curses.KEY_DC and filtered and kinds[selected] != "empty":
                choice = filtered[selected]
                if choice == current_dir.parent and current_dir.parent != current_dir:
                    self.message = "Cannot delete parent entry"
                    continue
                deleted = self._delete_browser_item(choice)
                if deleted:
                    selected = 0
                    offset = 0
                continue
            if key in (curses.KEY_UP, ord("k")):
                selected = (selected - 1) % len(filtered)
                continue
            if key in (curses.KEY_DOWN, ord("j")):
                selected = (selected + 1) % len(filtered)
                continue
            if key in (curses.KEY_BACKSPACE, 127):
                if filter_text:
                    filter_text = filter_text[:-1]
                    selected = 0
                    offset = 0
                continue
            if key in (curses.KEY_LEFT, ord("h")):
                parent = current_dir.parent
                if parent != current_dir:
                    current_dir = parent
                    selected = 0
                    offset = 0
                continue
            if key in (10, 13, curses.KEY_RIGHT, ord("l")):
                choice = filtered[selected]
                if choice.is_dir():
                    current_dir = choice
                    filter_text = ""
                    selected = 0
                    offset = 0
                    continue
                if kinds[selected] == "unsupported":
                    self.message = f"Unsupported file type: {choice.name}"
                    continue
                return choice
            if 32 <= key <= 126:
                filter_text += chr(key)
                selected = 0
                offset = 0

    def _browse_for_save(self, title: str, start_path: Path) -> Path | None:
        current_dir = start_path.expanduser().resolve().parent if start_path.suffix else start_path.expanduser().resolve()
        filename = start_path.name if start_path.suffix else "sheet.tss"
        filename_typed = False
        if not current_dir.exists():
            current_dir = Path.home()
        selected = 0
        offset = 0
        show_hidden = False
        filter_text = ""
        sort_mode = "name"
        curses.curs_set(0)
        while True:
            try:
                entries = list(current_dir.iterdir())
            except OSError as exc:
                self.message = f"Browse error: {exc}"
                return None
            visible_entries: list[Path] = [current_dir.parent] + entries if current_dir.parent != current_dir else entries
            filtered: list[Path] = []
            kinds: list[str] = []
            for index, entry in enumerate(visible_entries):
                if index == 0 and current_dir.parent != current_dir:
                    filtered.append(entry)
                    kinds.append("dir")
                    continue
                if not show_hidden and entry.name.startswith("."):
                    continue
                if filter_text and filter_text.lower() not in entry.name.lower():
                    continue
                filtered.append(entry)
                kinds.append(self._browser_item_kind(entry, None))
            if current_dir.parent != current_dir and filtered:
                parent = filtered[0]
                parent_kind = kinds[0]
                paired = list(zip(filtered[1:], kinds[1:]))
                paired.sort(key=lambda item: self._file_browser_sort_key(item[0], item[1], sort_mode))
                filtered = [parent] + [item[0] for item in paired]
                kinds = [parent_kind] + [item[1] for item in paired]
            else:
                paired = list(zip(filtered, kinds))
                paired.sort(key=lambda item: self._file_browser_sort_key(item[0], item[1], sort_mode))
                filtered = [item[0] for item in paired]
                kinds = [item[1] for item in paired]
            if not filtered:
                filtered = [current_dir]
                kinds = ["empty"]
                selected = 0
            selected = max(0, min(selected, len(filtered) - 1))
            height, width = self.stdscr.getmaxyx()
            self.stdscr.erase()
            self.stdscr.addnstr(0, 0, f" {title} ".ljust(width - 1), width - 1, self._bar_attr(bold=True))
            self.stdscr.addnstr(1, 0, str(current_dir).ljust(width - 1), width - 1, self._help_attr())
            hint = (
                f" Up/Down select  Enter=save/use  .. up  . hidden:{'on' if show_hidden else 'off'}"
                f"  s sort:{sort_mode}  p path  n mkdir  f filename  r rename  Del delete  type=filter  Esc cancel "
            )
            self.stdscr.addnstr(2, 0, hint.ljust(width - 1), width - 1, self._help_attr())
            filter_label = f" Filter: {filter_text or '(all)'} "
            self.stdscr.addnstr(3, 0, filter_label.ljust(width - 1), width - 1, self._help_attr())
            filename_label = f" Save as: {filename} "
            self.stdscr.addnstr(4, 0, filename_label.ljust(width - 1), width - 1, self._bar_attr(bold=True))

            top = 6
            content_height = max(6, height - top - 2)
            list_width = max(28, min(width * 3 // 5, width - 24))
            detail_x = min(width - 2, list_width + 2)
            detail_width = max(12, width - detail_x - 1)
            visible_rows = max(1, content_height - 2)
            if selected < offset:
                offset = selected
            elif selected >= offset + visible_rows:
                offset = selected - visible_rows + 1

            self.stdscr.addnstr(top, 0, f" Files [{sort_mode}] ".ljust(list_width - 1), list_width - 1, self._bar_attr(bold=True))
            if detail_width > 0:
                self.stdscr.addnstr(top, detail_x, " Details ".ljust(detail_width), detail_width, self._bar_attr(bold=True))
            name_width = max(10, list_width - 1 - 20)
            size_width = 8
            time_width = max(10, list_width - 1 - name_width - size_width - 4)
            list_header = f" {'Name'.ljust(name_width)} {'Size'.rjust(size_width)} {'Modified'.ljust(time_width)}"
            self.stdscr.addnstr(top + 1, 0, list_header.ljust(list_width - 1), list_width - 1, self._help_attr())
            for row_index in range(max(0, visible_rows - 1)):
                entry_index = offset + row_index
                if entry_index >= len(filtered):
                    break
                y = top + 2 + row_index
                entry = filtered[entry_index]
                kind = kinds[entry_index]
                attr = self._file_browser_entry_attr(kind, selected=(entry_index == selected))
                name = ".." if kind == "dir" and entry == current_dir.parent and current_dir.parent != current_dir else f"{self._browser_icon(kind, entry)}  {entry.name}{'/' if kind == 'dir' else ''}"
                size = self._format_file_size(entry, kind)
                modified = self._format_file_mtime(entry)
                row_text = f" {name[:name_width].ljust(name_width)} {size.rjust(size_width)} {modified[:time_width].ljust(time_width)}"
                self.stdscr.addnstr(y, 0, row_text.ljust(list_width - 1), list_width - 1, attr)
            choice = filtered[selected]
            selected_kind = kinds[selected]
            details = [
                f"Name: {choice.name or str(choice)}",
                f"Type: {'folder' if selected_kind == 'dir' else choice.suffix.lower().lstrip('.') or 'file'}",
                f"Target: {current_dir / filename}",
                "Enter saves the current target.",
                "Right/l on folder opens it.",
                "Enter on file uses its name then saves.",
                "Press f to type filename.",
            ]
            for index, line in enumerate(details):
                y = top + 1 + index
                if y >= top + content_height or detail_width <= 0:
                    break
                self.stdscr.addnstr(y, detail_x, line.ljust(detail_width), detail_width, self._help_attr())
            footer = " Enter=save/use   f=filename   r=rename   Del=delete   s=sort   p=path   n=mkdir   .=hidden   type to filter   Esc=cancel "
            self.stdscr.addnstr(height - 1, 0, footer.ljust(width - 1), width - 1, self._bar_attr())
            self.stdscr.refresh()
            key = self.stdscr.getch()
            if key == 27:
                return None
            if key == ord("."):
                show_hidden = not show_hidden
                selected = 0
                offset = 0
                continue
            if key in (ord("s"), ord("S")):
                current_index = FILE_BROWSER_SORT_OPTIONS.index(sort_mode)
                sort_mode = FILE_BROWSER_SORT_OPTIONS[(current_index + 1) % len(FILE_BROWSER_SORT_OPTIONS)]
                selected = 0
                offset = 0
                continue
            if key in (ord("f"), ord("F")):
                typed = self.prompt("Filename: ", filename)
                if typed is None or not typed.strip():
                    continue
                filename = Path(typed).name or filename
                filename_typed = True
                return (current_dir / filename).expanduser()
            if key in (ord("p"), ord("P")):
                typed = self.prompt("Save path: ", str(current_dir / filename))
                if typed is None or not typed.strip():
                    continue
                target = Path(typed).expanduser()
                if target.is_dir():
                    current_dir = target.resolve()
                    filename_typed = False
                    continue
                return target
            if key in (ord("n"), ord("N")):
                folder_text = self.prompt("New folder: ", str(current_dir / "new-folder"))
                if folder_text is None or not folder_text.strip():
                    continue
                target = Path(folder_text).expanduser()
                if not target.is_absolute():
                    target = current_dir / target
                try:
                    target.mkdir(parents=True, exist_ok=False)
                except OSError as exc:
                    self.message = f"Create folder error: {exc}"
                    continue
                current_dir = target.resolve()
                selected = 0
                offset = 0
                filter_text = ""
                filename_typed = False
                continue
            if key in (ord("r"), ord("R")) and filtered and kinds[selected] != "empty":
                choice = filtered[selected]
                if choice == current_dir.parent and current_dir.parent != current_dir:
                    self.message = "Cannot rename parent entry"
                    continue
                renamed = self._rename_browser_item(choice)
                if renamed is not None:
                    if not renamed.is_dir():
                        filename = renamed.name
                    selected = 0
                    offset = 0
                continue
            if key == curses.KEY_DC and filtered and kinds[selected] != "empty":
                choice = filtered[selected]
                if choice == current_dir.parent and current_dir.parent != current_dir:
                    self.message = "Cannot delete parent entry"
                    continue
                deleted = self._delete_browser_item(choice)
                if deleted:
                    selected = 0
                    offset = 0
                continue
            if key in (curses.KEY_UP, ord("k")):
                selected = (selected - 1) % len(filtered)
                continue
            if key in (curses.KEY_DOWN, ord("j")):
                selected = (selected + 1) % len(filtered)
                continue
            if key in (curses.KEY_BACKSPACE, 127):
                if filter_text:
                    filter_text = filter_text[:-1]
                    selected = 0
                    offset = 0
                continue
            if key in (curses.KEY_LEFT, ord("h")):
                parent = current_dir.parent
                if parent != current_dir:
                    current_dir = parent
                    selected = 0
                    offset = 0
                    filename_typed = False
                continue
            if key in (curses.KEY_RIGHT, ord("l")):
                choice = filtered[selected]
                if choice.is_dir():
                    current_dir = choice.resolve()
                    filter_text = ""
                    selected = 0
                    offset = 0
                    filename_typed = False
                    continue
                if choice.exists():
                    filename = choice.name
                    filename_typed = False
                return (current_dir / filename).expanduser()
            if key in (10, 13):
                choice = filtered[selected]
                if not choice.is_dir() and choice.exists():
                    filename = choice.name
                    filename_typed = False
                return (current_dir / filename).expanduser()
            if 32 <= key <= 126:
                filter_text += chr(key)
                selected = 0
                offset = 0

    def _menu_selected_attr(self) -> int:
        attr = curses.A_BOLD
        if self.colors_ready:
            attr |= curses.color_pair(COLOR_PAIR_MENU_SELECTED)
        else:
            attr |= curses.A_REVERSE
        return attr

    def _handle_escape_sequence(self) -> bool:
        self.stdscr.timeout(20)
        try:
            next_key = self.stdscr.getch()
            if next_key == -1:
                return False
            if next_key in (ord("="), ord("+")):
                self._sum_above_into_current_cell()
                return True
            if next_key != ord("["):
                return False
            sequence = []
            while True:
                key = self.stdscr.getch()
                if key == -1:
                    return False
                sequence.append(chr(key))
                if 64 <= key <= 126:
                    break
            sequence_text = "".join(sequence)
            if sequence_text == "200~":
                pasted = self._read_bracketed_paste()
                if self.prefer_internal_clipboard and self.clipboard_cells:
                    self._paste_internal_clipboard()
                    return True
                if self.clipboard_cells and pasted.rstrip("\n") == self._clipboard_plain_text().rstrip("\n"):
                    self._paste_internal_clipboard()
                    return True
                if self._load_clipboard_payload(pasted):
                    return True
                if pasted:
                    start_row, start_col, _end_row, _end_col = self._target_range(None)
                    if self.sheet.is_protected(start_row, start_col):
                        self.message = "Cell is protected."
                        return True
                    self._save_undo_state()
                    value = self._normalize_cell_input(start_row, start_col, pasted)
                    self.sheet.set_raw(start_row, start_col, value)
                    self._apply_default_alignment(start_row, start_col, value)
                    self.dirty = True
                    self.message = f"Pasted to {self._range_label(start_row, start_col, start_row, start_col)}"
                    return True
                return False
            if sequence_text.endswith("u"):
                payload = sequence_text[:-1]
                parts = payload.split(";")
                if parts and parts[0] == "32":
                    modifier = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 1
                    if modifier in {2, 6}:
                        self._select_row(self.current_row)
                        return True
                    if modifier in {5}:
                        self._select_column(self.current_col)
                        return True
            return False
        finally:
            self.stdscr.timeout(-1)

    def _read_bracketed_paste(self) -> str:
        chars: list[str] = []
        while True:
            key = self.stdscr.getch()
            if key == -1:
                break
            if key != 27:
                chars.append(chr(key))
                continue
            next_key = self.stdscr.getch()
            if next_key != ord("["):
                chars.append(chr(27))
                if next_key != -1:
                    chars.append(chr(next_key))
                continue
            trailer = []
            for _ in range(4):
                follow = self.stdscr.getch()
                if follow == -1:
                    break
                trailer.append(chr(follow))
            if "".join(trailer) == "201~":
                break
            chars.append(chr(27))
            chars.append("[")
            chars.extend(trailer)
        return "".join(chars)

    def _sum_above_into_current_cell(self) -> None:
        if self.sheet.is_protected(self.current_row, self.current_col):
            self.message = "Cell is protected."
            return
        self._save_undo_state()
        if self.current_row <= 0:
            self.sheet.set_raw(self.current_row, self.current_col, "0")
            self._apply_default_alignment(self.current_row, self.current_col, "0")
            self.dirty = True
            self.message = f"Stored 0 in {column_label(self.current_col)}{self.current_row + 1}"
            return
        label = column_label(self.current_col)
        formula = f"=SUM({label}1:{label}{self.current_row})"
        self.sheet.set_raw(self.current_row, self.current_col, formula)
        self._apply_default_alignment(self.current_row, self.current_col, formula)
        self.dirty = True
        self.message = f"Summed above into {label}{self.current_row + 1}"

    def _apply_default_alignment(self, row: int, col: int, raw: str) -> None:
        if self.sheet.is_alignment_manual(row, col):
            return
        if should_auto_right_align(raw):
            self.sheet.set_alignment(row, col, "right", manual=False)
        else:
            self.sheet.set_alignment(row, col, "", manual=False)

    def _apply_format(self, text: str, style: str) -> str:
        if not text:
            return text
        normalized_style = style.lower()
        if style.startswith("date"):
            return format_date_text(text, style)
        if style.startswith("time"):
            return format_time_text(text, style)
        if self.sheet.date_format.startswith("date") and parse_date_text(text, self.sheet.date_format) is not None:
            try:
                return format_date_text(text, self.sheet.date_format)
            except ValueError:
                pass
        if self.sheet.time_format.startswith("time") and parse_time_text(text) is not None:
            try:
                return format_time_text(text, self.sheet.time_format)
            except ValueError:
                pass
        if style == "time" and parse_time_text(text) is not None:
            try:
                return format_time_text(text, "time:24h")
            except ValueError:
                pass
        try:
            number = float(text)
        except ValueError:
            return text
        if normalized_style == "currency" or normalized_style.startswith("currency:"):
            symbol = style.split(":", 1)[1] if ":" in style else "£"
            return f"{symbol}{number:,.2f}"
        if normalized_style == "fixed":
            return f"{number:.2f}"
        if normalized_style == "percent":
            return f"{number:.2f}%"
        if normalized_style == "int":
            return str(int(round(number)))
        if normalized_style == "negative":
            return text
        if normalized_style == "accounting":
            abs_value = abs(number)
            rendered = f"{abs_value:,.2f}"
            return f"({rendered})" if number < 0 else rendered
        if normalized_style in {"sci", "scientific"}:
            return f"{number:.2e}"
        return text

    def _scroll_into_view(self) -> None:
        height, width = self.stdscr.getmaxyx()
        _top_grid_row, grid_height, _row_header_width, _visible_columns = self._grid_layout(height, width)
        visible_rows = max(1, len(self._visible_rows(grid_height)))
        if self.sheet.is_row_hidden(self.current_row):
            self.current_row = self._first_visible_row()
        if self.sheet.is_col_hidden(self.current_col):
            self.current_col = self._first_visible_col()
        if self.current_row < self.row_offset:
            self.row_offset = self.current_row
        else:
            while True:
                visible = self._visible_rows(visible_rows)
                if self.current_row in visible:
                    break
                self.row_offset = min(self.sheet.rows - 1, self.row_offset + 1)
        if self.current_col < self.col_offset:
            self.col_offset = self.current_col
        else:
            while True:
                visible = self._visible_columns(width, 6)
                cols = [col for col, _x, _w in visible]
                if self.current_col in cols:
                    break
                self.col_offset += 1

    def _range_label(self, row_lo: int, col_lo: int, row_hi: int, col_hi: int) -> str:
        start = f"{column_label(col_lo)}{row_lo + 1}"
        end = f"{column_label(col_hi)}{row_hi + 1}"
        return start if start == end else f"{start}:{end}"

    def _rebuild_rows(self, index: int, delta: int) -> None:
        new_sheet = Spreadsheet(
            rows=max(1, self.sheet.rows + delta),
            cols=self.sheet.cols,
            column_width=self.sheet.column_width,
            column_widths=self.sheet.column_widths.copy(),
            row_backgrounds=self.sheet.row_backgrounds.copy(),
            hidden_rows=self.sheet.hidden_rows.copy(),
            hidden_cols=self.sheet.hidden_cols.copy(),
            title_rows=self.sheet.title_rows,
            title_cols=self.sheet.title_cols,
            theme_name=self.sheet.theme_name,
            date_format=self.sheet.date_format,
            active_cell_color=self.sheet.active_cell_color,
            formula_coloration=self.sheet.formula_coloration,
            formula_foreground_color=self.sheet.formula_foreground_color,
            language=self.sheet.language,
            protected_foreground_color=self.sheet.protected_foreground_color,
            protected_background_color=self.sheet.protected_background_color,
            named_ranges=self.sheet.named_ranges.copy(),
        )
        shifted_hidden_rows: set[int] = set()
        shifted_row_backgrounds: dict[int, str] = {}
        for row in self.sheet.hidden_rows:
            new_row = row
            if delta > 0 and row >= index:
                new_row = row + delta
            elif delta < 0:
                if index <= row < index - delta:
                    continue
                if row >= index - delta:
                    new_row = row + delta
            shifted_hidden_rows.add(new_row)
        new_sheet.hidden_rows = shifted_hidden_rows
        for row, color in self.sheet.row_backgrounds.items():
            new_row = row
            if delta > 0 and row >= index:
                new_row = row + delta
            elif delta < 0:
                if index <= row < index - delta:
                    continue
                if row >= index - delta:
                    new_row = row + delta
            shifted_row_backgrounds[new_row] = color
        new_sheet.row_backgrounds = shifted_row_backgrounds
        for row, col, raw in self.sheet.iter_cells():
            new_row = row
            if delta > 0 and row >= index:
                new_row = row + delta
            elif delta < 0:
                if index <= row < index - delta:
                    continue
                if row >= index - delta:
                    new_row = row + delta
            new_sheet.set_raw(new_row, col, raw)
        for row, col, style in self.sheet.iter_formats():
            new_row = row
            if delta > 0 and row >= index:
                new_row = row + delta
            elif delta < 0:
                if index <= row < index - delta:
                    continue
                if row >= index - delta:
                    new_row = row + delta
            new_sheet.set_format(new_row, col, style)
        for row, col, styles in self.sheet.iter_text_styles():
            new_row = row
            if delta > 0 and row >= index:
                new_row = row + delta
            elif delta < 0:
                if index <= row < index - delta:
                    continue
                if row >= index - delta:
                    new_row = row + delta
            for style in styles:
                new_sheet.set_text_style(new_row, col, style, enabled=True)
        for row, col, background in self.sheet.iter_backgrounds():
            new_row = row
            if delta > 0 and row >= index:
                new_row = row + delta
            elif delta < 0:
                if index <= row < index - delta:
                    continue
                if row >= index - delta:
                    new_row = row + delta
            new_sheet.set_background(new_row, col, background)
        for row, col, border in self.sheet.iter_borders():
            new_row = row
            if delta > 0 and row >= index:
                new_row = row + delta
            elif delta < 0:
                if index <= row < index - delta:
                    continue
                if row >= index - delta:
                    new_row = row + delta
            new_sheet.set_border(new_row, col, border)
        for row, col, align in self.sheet.iter_alignments():
            new_row = row
            if delta > 0 and row >= index:
                new_row = row + delta
            elif delta < 0:
                if index <= row < index - delta:
                    continue
                if row >= index - delta:
                    new_row = row + delta
            new_sheet.set_alignment(new_row, col, align, manual=self.sheet.is_alignment_manual(row, col))
        for row, col in self.sheet.iter_protected():
            new_row = row
            if delta > 0 and row >= index:
                new_row = row + delta
            elif delta < 0:
                if index <= row < index - delta:
                    continue
                if row >= index - delta:
                    new_row = row + delta
            new_sheet.protect(new_row, col)
        self.sheet = new_sheet
        self.evaluator = Evaluator(self.sheet)

    def _rebuild_cols(self, index: int, delta: int) -> None:
        new_sheet = Spreadsheet(
            rows=self.sheet.rows,
            cols=max(1, self.sheet.cols + delta),
            column_width=self.sheet.column_width,
            row_backgrounds=self.sheet.row_backgrounds.copy(),
            hidden_rows=self.sheet.hidden_rows.copy(),
            hidden_cols=self.sheet.hidden_cols.copy(),
            title_rows=self.sheet.title_rows,
            title_cols=self.sheet.title_cols,
            theme_name=self.sheet.theme_name,
            date_format=self.sheet.date_format,
            active_cell_color=self.sheet.active_cell_color,
            formula_coloration=self.sheet.formula_coloration,
            formula_foreground_color=self.sheet.formula_foreground_color,
            language=self.sheet.language,
            protected_foreground_color=self.sheet.protected_foreground_color,
            protected_background_color=self.sheet.protected_background_color,
            named_ranges=self.sheet.named_ranges.copy(),
        )
        shifted_hidden_cols: set[int] = set()
        for col in self.sheet.hidden_cols:
            new_col = col
            if delta > 0 and col >= index:
                new_col = col + delta
            elif delta < 0:
                if index <= col < index - delta:
                    continue
                if col >= index - delta:
                    new_col = col + delta
            shifted_hidden_cols.add(new_col)
        new_sheet.hidden_cols = shifted_hidden_cols
        for key, value in self.sheet.column_widths.items():
            old_col = int(key)
            new_col = old_col
            if delta > 0 and old_col >= index:
                new_col = old_col + delta
            elif delta < 0:
                if index <= old_col < index - delta:
                    continue
                if old_col >= index - delta:
                    new_col = old_col + delta
            new_sheet.column_widths[str(new_col)] = value
        for row, col, raw in self.sheet.iter_cells():
            if delta > 0 and col >= index:
                col += delta
            elif delta < 0:
                if index <= col < index - delta:
                    continue
                if col >= index - delta:
                    col += delta
            new_sheet.set_raw(row, col, raw)
        for row, col, style in self.sheet.iter_formats():
            if delta > 0 and col >= index:
                col += delta
            elif delta < 0:
                if index <= col < index - delta:
                    continue
                if col >= index - delta:
                    col += delta
            new_sheet.set_format(row, col, style)
        for row, col, styles in self.sheet.iter_text_styles():
            if delta > 0 and col >= index:
                col += delta
            elif delta < 0:
                if index <= col < index - delta:
                    continue
                if col >= index - delta:
                    col += delta
            for style in styles:
                new_sheet.set_text_style(row, col, style, enabled=True)
        for row, col, background in self.sheet.iter_backgrounds():
            if delta > 0 and col >= index:
                col += delta
            elif delta < 0:
                if index <= col < index - delta:
                    continue
                if col >= index - delta:
                    col += delta
            new_sheet.set_background(row, col, background)
        for row, col, border in self.sheet.iter_borders():
            if delta > 0 and col >= index:
                col += delta
            elif delta < 0:
                if index <= col < index - delta:
                    continue
                if col >= index - delta:
                    col += delta
            new_sheet.set_border(row, col, border)
        for row, col, align in self.sheet.iter_alignments():
            if delta > 0 and col >= index:
                col += delta
            elif delta < 0:
                if index <= col < index - delta:
                    continue
                if col >= index - delta:
                    col += delta
            new_sheet.set_alignment(row, col, align, manual=self.sheet.is_alignment_manual(row, col))
        for row, col in self.sheet.iter_protected():
            if delta > 0 and col >= index:
                col += delta
            elif delta < 0:
                if index <= col < index - delta:
                    continue
                if col >= index - delta:
                    col += delta
            new_sheet.protect(row, col)
        self.sheet = new_sheet
        self.evaluator = Evaluator(self.sheet)

    def _move_rows(self, start: int, destination: int, count: int) -> None:
        order = list(range(self.sheet.rows))
        chunk = order[start : start + count]
        del order[start : start + count]
        destination = max(0, min(len(order), destination))
        for offset, row_index in enumerate(chunk):
            order.insert(destination + offset, row_index)
        self._reorder_rows(order)

    def _move_cols(self, start: int, destination: int, count: int) -> None:
        order = list(range(self.sheet.cols))
        chunk = order[start : start + count]
        del order[start : start + count]
        destination = max(0, min(len(order), destination))
        for offset, col_index in enumerate(chunk):
            order.insert(destination + offset, col_index)
        self._reorder_cols(order)

    def _reorder_rows(self, order: list[int]) -> None:
        new_sheet = Spreadsheet(
            rows=len(order),
            cols=self.sheet.cols,
            column_width=self.sheet.column_width,
            column_widths=self.sheet.column_widths.copy(),
            row_backgrounds=self.sheet.row_backgrounds.copy(),
            hidden_cols=self.sheet.hidden_cols.copy(),
            title_rows=self.sheet.title_rows,
            title_cols=self.sheet.title_cols,
            theme_name=self.sheet.theme_name,
            date_format=self.sheet.date_format,
            active_cell_color=self.sheet.active_cell_color,
            formula_coloration=self.sheet.formula_coloration,
            formula_foreground_color=self.sheet.formula_foreground_color,
            language=self.sheet.language,
            protected_foreground_color=self.sheet.protected_foreground_color,
            protected_background_color=self.sheet.protected_background_color,
            named_ranges=self.sheet.named_ranges.copy(),
        )
        new_sheet.hidden_rows = {new_row for new_row, old_row in enumerate(order) if old_row in self.sheet.hidden_rows}
        for new_row, old_row in enumerate(order):
            row_background = self.sheet.get_row_background(old_row)
            if row_background:
                new_sheet.set_row_background(new_row, row_background)
            for col in range(self.sheet.cols):
                raw = self.sheet.get_raw(old_row, col)
                if raw:
                    new_sheet.set_raw(new_row, col, raw)
                style = self.sheet.get_format(old_row, col)
                if style:
                    new_sheet.set_format(new_row, col, style)
                for text_style in self.sheet.get_text_styles(old_row, col):
                    new_sheet.set_text_style(new_row, col, text_style, enabled=True)
                background = self.sheet.get_background(old_row, col)
                if background:
                    new_sheet.set_background(new_row, col, background)
                border = self.sheet.get_border(old_row, col)
                if border:
                    new_sheet.set_border(new_row, col, border)
                align = self.sheet.get_alignment(old_row, col)
                if align:
                    new_sheet.set_alignment(new_row, col, align, manual=self.sheet.is_alignment_manual(old_row, col))
                if self.sheet.is_protected(old_row, col):
                    new_sheet.protect(new_row, col)
        self.sheet = new_sheet
        self.evaluator = Evaluator(self.sheet)

    def _reorder_cols(self, order: list[int]) -> None:
        new_sheet = Spreadsheet(
            rows=self.sheet.rows,
            cols=len(order),
            column_width=self.sheet.column_width,
            hidden_rows=self.sheet.hidden_rows.copy(),
            row_backgrounds=self.sheet.row_backgrounds.copy(),
            title_rows=self.sheet.title_rows,
            title_cols=self.sheet.title_cols,
            theme_name=self.sheet.theme_name,
            date_format=self.sheet.date_format,
            active_cell_color=self.sheet.active_cell_color,
            formula_coloration=self.sheet.formula_coloration,
            formula_foreground_color=self.sheet.formula_foreground_color,
            language=self.sheet.language,
            protected_foreground_color=self.sheet.protected_foreground_color,
            protected_background_color=self.sheet.protected_background_color,
            named_ranges=self.sheet.named_ranges.copy(),
        )
        new_sheet.hidden_cols = {new_col for new_col, old_col in enumerate(order) if old_col in self.sheet.hidden_cols}
        for new_col, old_col in enumerate(order):
            width = self.sheet.column_widths.get(str(old_col))
            if width is not None:
                new_sheet.column_widths[str(new_col)] = width
        for row in range(self.sheet.rows):
            for new_col, old_col in enumerate(order):
                raw = self.sheet.get_raw(row, old_col)
                if raw:
                    new_sheet.set_raw(row, new_col, raw)
                style = self.sheet.get_format(row, old_col)
                if style:
                    new_sheet.set_format(row, new_col, style)
                for text_style in self.sheet.get_text_styles(row, old_col):
                    new_sheet.set_text_style(row, new_col, text_style, enabled=True)
                background = self.sheet.get_background(row, old_col)
                if background:
                    new_sheet.set_background(row, new_col, background)
                border = self.sheet.get_border(row, old_col)
                if border:
                    new_sheet.set_border(row, new_col, border)
                align = self.sheet.get_alignment(row, old_col)
                if align:
                    new_sheet.set_alignment(row, new_col, align, manual=self.sheet.is_alignment_manual(row, old_col))
                if self.sheet.is_protected(row, old_col):
                    new_sheet.protect(row, new_col)
        self.sheet = new_sheet
        self.evaluator = Evaluator(self.sheet)


def _run_multiple(stdscr, paths: list[Path], settings_path: Path) -> int:
    app = SpreadsheetApp(stdscr, path=None, settings_path=settings_path)
    loadable_suffixes = {".tss", ".csv", ".tsv"}
    loaded_any = False
    skipped: list[str] = []
    for path in paths:
        if not path.exists():
            skipped.append(str(path))
            continue
        if path.suffix.lower() not in loadable_suffixes:
            skipped.append(str(path))
            continue
        app._add_loaded_tab(path, switch=True)
        loaded_any = True
    if loaded_any and skipped:
        app.message = f"Loaded {len(app.tabs)} tab(s); skipped {len(skipped)} non-sheet path(s)."
    elif loaded_any:
        app.message = f"Loaded {len(app.tabs)} tab(s)."
    elif skipped:
        app.message = f"Skipped {len(skipped)} non-sheet path(s)."
    return app.run()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="A modular curses spreadsheet.")
    parser.add_argument("path", nargs="*", help="Optional .tss/.csv/.tsv file(s) to open")
    args = parser.parse_args(argv)
    paths = [Path(item).expanduser() for item in args.path]
    settings_path = DEFAULT_SETTINGS_PATH
    return curses.wrapper(lambda stdscr: _run_multiple(stdscr, paths, settings_path))


if __name__ == "__main__":
    raise SystemExit(main())
