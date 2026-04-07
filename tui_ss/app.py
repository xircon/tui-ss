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
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from .commands import (
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
    is_formula_text,
    parse_date_text,
    normalize_date_text,
    shift_formula_references,
)
from .model import Spreadsheet, column_label, parse_cell_reference
from .storage import load_app_settings, load_sheet, save_app_settings, save_pdf_text, save_sheet

APP_NAME = "tui-ss"
DEFAULT_PATH = Path.home() / "scripts" / "tui-ss" / "sheets" / "autosave.tss"
DEFAULT_SETTINGS_PATH = Path.home() / "scripts" / "tui-ss" / "tui-ss-settings.toml"
THEMES = ["blue", "cyan", "magenta", "purple", "white", "yellow"]
ACTIVE_CELL_COLORS = ["yellow", "pink", "orange", "white", "lightblue", "cornflower", "lightgrey"]
PROTECTED_COLOR_OPTIONS = ["black", "white", "yellow", "pink", "palepink", "orange", "lightblue", "cornflower", "lightgrey", "blue", "cyan", "green", "magenta", "red"]
FORMULA_COLOR_OPTIONS = ["green", "yellow", "cyan", "magenta", "orange", "lightblue", "cornflower", "white", "red", "blue"]
FORMULA_COLOR_SETTING_OPTIONS = ["off"] + FORMULA_COLOR_OPTIONS
FORMAT_STYLES = ["accounting", "background", "bold", "clear-format", "currency", "date", "fixed", "int", "italic", "negative", "percent", "sci", "text", "underline"]
CURRENCY_SYMBOLS = ["£", "€", "$", "¥", "₹"]
DATE_FORMATS = ["european", "us", "ansi"]
BACKGROUND_COLORS = ["blue", "cyan", "green", "magenta", "none", "red", "white", "yellow"]
JUSTIFY_OPTIONS = ["left", "centre", "right"]
FILE_BROWSER_SORT_OPTIONS = ["name", "time", "type"]
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
CLIPBOARD_MARKER = "TUI-SS-CLIP:"
RECENT_FILES_LIMIT = 10
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
    "MIN": "MIN(range)",
    "MOD": "MOD(value, divisor)",
    "MONTH": "MONTH(date)",
    "NOT": "NOT(value)",
    "OR": "OR(value1, value2, ...)",
    "RIGHT": "RIGHT(text, count)",
    "ROUND": "ROUND(value, digits)",
    "SIN": "SIN(value)",
    "SQRT": "SQRT(value)",
    "SUM": "SUM(range)",
    "TAN": "TAN(value)",
    "TEXT": 'TEXT(value, "format")',
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
    "MIN": ["range"],
    "MOD": ["value", "divisor"],
    "MONTH": ["date"],
    "NOT": ["value"],
    "OR": ["value1", "value2", "..."],
    "RIGHT": ["text", "count"],
    "ROUND": ["value", "digits"],
    "SIN": ["value"],
    "SQRT": ["value"],
    "SUM": ["range"],
    "TAN": ["value"],
    "TEXT": ["value", "format"],
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
        self.next_dynamic_pair = 20
        self.selection_anchor: tuple[int, int] | None = None
        self.selection_range: tuple[int, int, int, int] | None = None
        self.mouse_dragging = False
        self.dirty = False
        self.clipboard_cells: list[tuple[int, int, str, str, str, str, str, bool]] = []
        self.clipboard_size: tuple[int, int] = (0, 0)
        self.clipboard_origin: tuple[int, int] = (0, 0)
        self.undo_stack: list[dict[str, object]] = []
        self.redo_stack: list[dict[str, object]] = []
        self.max_history = 100
        self.tabs: list[TabState] = []
        self.current_tab_index = 0
        self.recent_files: list[str] = []
        self.command_hint_visible = True
        self._load_global_settings()
        self.tabs.append(self._capture_tab_state())

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
        self.next_dynamic_pair = 20
        text_color = self._theme_text_color()
        curses.init_pair(COLOR_PAIR_TEXT, text_color, -1)
        curses.init_pair(COLOR_PAIR_FORMULA, self._formula_foreground_color(), -1)
        curses.init_pair(COLOR_PAIR_HEADER, text_color, curses.COLOR_BLACK)
        curses.init_pair(COLOR_PAIR_BAR, text_color, -1)
        curses.init_pair(COLOR_PAIR_GRID, curses.COLOR_BLACK, -1)
        curses.init_pair(COLOR_PAIR_GRID_ROW, text_color, -1)
        curses.init_pair(COLOR_PAIR_MENU_SELECTED, self._selection_foreground_color(), self._selection_background_color())
        curses.init_pair(COLOR_PAIR_SELECTION, self._selection_foreground_color(), self._selection_background_color())
        curses.init_pair(COLOR_PAIR_ROW_HEADER, text_color, curses.COLOR_BLACK)
        curses.init_pair(COLOR_PAIR_NEGATIVE, curses.COLOR_RED, -1)

    def _settings_payload(self) -> dict[str, str]:
        return {
            "theme_name": self.sheet.theme_name,
            "date_format": self.sheet.date_format,
            "active_cell_color": self.sheet.active_cell_color,
            "formula_coloration": "on" if self.sheet.formula_coloration else "off",
            "formula_foreground_color": self.sheet.formula_foreground_color,
            "language": self.sheet.language,
            "protected_foreground_color": self.sheet.protected_foreground_color,
            "protected_background_color": self.sheet.protected_background_color,
            "recent_files_json": json.dumps(self.recent_files),
        }

    def _apply_settings_payload(self, settings: dict[str, str]) -> None:
        theme_name = settings.get("theme_name")
        if theme_name in THEMES:
            self.sheet.theme_name = theme_name
        raw_date_format = settings.get("date_format")
        if raw_date_format:
            if raw_date_format in DATE_FORMATS:
                self.sheet.date_format = f"date:{raw_date_format}"
            elif raw_date_format.startswith("date:"):
                self.sheet.date_format = raw_date_format
        active_cell_color = settings.get("active_cell_color")
        if active_cell_color in ACTIVE_CELL_COLORS:
            self.sheet.active_cell_color = active_cell_color
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

    def _theme_text_color(self) -> int:
        if self.sheet.theme_name == "purple":
            custom_purple = self._custom_purple_color()
            if custom_purple is not None:
                return custom_purple
        return THEME_COLOR_MAP.get(self.sheet.theme_name, curses.COLOR_WHITE)

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
        return default

    def draw(self) -> None:
        self.stdscr.erase()
        height, width = self.stdscr.getmaxyx()
        top_grid_row, grid_height, row_header_width, visible_columns = self._grid_layout(height, width)
        self.stdscr.addnstr(0, 0, self._tabs_line(width), width - 1, self._bar_attr(bold=True))
        self.stdscr.addnstr(1, 0, self._title_line(width), width - 1, self._bar_attr(bold=True))
        self.stdscr.addnstr(2, 0, self._top_formula_line(width), width - 1, self._bar_attr())
        self._draw_grid(top_grid_row - 1, grid_height, row_header_width, visible_columns)
        self._draw_column_headers(top_grid_row - 1, visible_columns)

        for screen_row in range(grid_height):
            row = self.row_offset + screen_row
            if row >= self.sheet.rows:
                break
            y = top_grid_row + screen_row
            self.stdscr.addnstr(y, 0, f"{row + 1:>5}", row_header_width - 1, self._row_header_attr())
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

        self._draw_command_hint(height, width)
        self.stdscr.addnstr(height - 1, 0, self.message.ljust(width - 1), width - 1, self._bar_attr(bold=True))
        self._draw_settings_cog(height, width)
        self.stdscr.refresh()

    def _draw_command_hint(self, height: int, width: int) -> None:
        if not self.command_hint_visible:
            return
        text = "Press / to start"
        y = max(0, height - 2)
        x = max(0, (width - len(text)) // 2)
        attr = self._help_attr()
        self.stdscr.addnstr(y, x, text, min(len(text), width - x - 1), attr)

    def _settings_label(self) -> str:
        return "[⚙]"

    def _draw_settings_cog(self, height: int, width: int) -> None:
        label = self._settings_label()
        x = max(0, width - len(label) - 1)
        attr = self._bar_attr(bold=True)
        self.stdscr.addnstr(height - 1, x, label, len(label), attr)

    def _draw_column_headers(self, y: int, visible_columns: list[tuple[int, int, int]]) -> None:
        for col, x, col_width in visible_columns:
            label = column_label(col)
            if col < self.sheet.title_cols:
                label = f"*{label}"
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

    def handle_key(self, key: int) -> None:
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
        self.current_row = max(0, min(self.sheet.rows - 1, self.current_row + row_delta))
        self.current_col = max(0, min(self.sheet.cols - 1, self.current_col + col_delta))
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
        self.current_row = max(0, min(self.sheet.rows - 1, self.current_row + row_delta))
        self.current_col = max(0, min(self.sheet.cols - 1, self.current_col + col_delta))
        self.selection_range = self._normalize_range(self.selection_anchor, (self.current_row, self.current_col))
        self._scroll_into_view()
        self.message = f"Selected {self._selection_label()}"

    def _capture_tab_state(self) -> TabState:
        return TabState(
            sheet=self.sheet,
            path=self.path,
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
        if not self.tabs:
            self.tabs.append(self._capture_tab_state())
            self.current_tab_index = 0
            return
        self.tabs[self.current_tab_index] = self._capture_tab_state()

    def _restore_tab_state(self, tab: TabState) -> None:
        self.sheet = tab.sheet
        self.evaluator = Evaluator(self.sheet)
        self.path = tab.path
        self.dirty = tab.dirty
        self.current_row = tab.current_row
        self.current_col = tab.current_col
        self.row_offset = tab.row_offset
        self.col_offset = tab.col_offset
        self.selection_anchor = tab.selection_anchor
        self.selection_range = tab.selection_range
        self.mouse_dragging = tab.mouse_dragging
        self.undo_stack = list(tab.undo_stack)
        self.redo_stack = list(tab.redo_stack)
        self._refresh_theme_colors()

    def _tab_label(self, index: int, tab: TabState) -> str:
        base = tab.path.name if tab.path else f"untitled-{index + 1}"
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
        new_tab = TabState(
            sheet=loaded_sheet,
            path=target,
            dirty=False,
            current_row=0,
            current_col=0,
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
        else:
            self._store_current_tab_state()
            self.tabs.append(new_tab)
            if switch:
                self.current_tab_index = len(self.tabs) - 1
                self._restore_tab_state(new_tab)
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
        top_grid_row, _grid_height, _row_header_width, _visible_columns = self._grid_layout(height, width)
        visible_rows = max(1, height - top_grid_row - 1)
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
            row = self.row_offset + (y - top_grid_row)
            if row < self.sheet.rows:
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
            ["freeze", "insert before", "delete", "width"],
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
            ["freeze", "insert above", "delete"],
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
        if self.sheet.is_protected(self.current_row, self.current_col):
            self.message = "Cell is protected."
            return
        row = self.current_row
        col = self.current_col
        cell_ref = f"{column_label(col)}{row + 1}"
        self._save_undo_state()
        self.sheet.clear(row, col)
        if not self.sheet.is_alignment_manual(row, col):
            self.sheet.set_alignment(row, col, "", manual=False)
        self.dirty = True
        self.message = f"Cleared {cell_ref}"

    def copy_selection_to_clipboard(self) -> None:
        row_lo, col_lo, row_hi, col_hi = self._target_range(None)
        cells: list[tuple[int, int, str, str, str, str, str, bool]] = []
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
                        self.sheet.get_alignment(row, col),
                        self.sheet.is_alignment_manual(row, col),
                    )
                )
        self.clipboard_cells = cells
        self.clipboard_size = (row_hi - row_lo + 1, col_hi - col_lo + 1)
        self.clipboard_origin = (row_lo, col_lo)
        self._export_clipboard_to_terminal()
        self._export_clipboard_to_system()
        self.message = f"Copied {self._range_label(row_lo, col_lo, row_hi, col_hi)}"

    def paste_clipboard(self) -> None:
        if self._paste_from_system_clipboard():
            return
        if not self.clipboard_cells:
            self.message = "Clipboard is empty."
            return
        self._save_undo_state()
        cells_by_offset = {
            (row_offset, col_offset): (raw, style, text_styles, background, align, align_manual)
            for row_offset, col_offset, raw, style, text_styles, background, align, align_manual in self.clipboard_cells
        }
        clip_height, clip_width = self.clipboard_size
        if self.selection_range is not None:
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
                raw, style, text_styles, background, align, align_manual = cells_by_offset[(row_offset, col_offset)]
                src_row = self.clipboard_origin[0] + row_offset
                src_col = self.clipboard_origin[1] + col_offset
                shifted_raw = shift_formula_references(raw, row - src_row, col - src_col)
                self.sheet.set_raw(row, col, shifted_raw)
                self.sheet.set_format(row, col, style)
                self.sheet.clear_text_styles(row, col)
                for text_style in [item for item in text_styles.split(",") if item]:
                    self.sheet.set_text_style(row, col, text_style, enabled=True)
                self.sheet.set_background(row, col, background)
                self.sheet.set_alignment(row, col, align, manual=align_manual)
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

    def _export_clipboard_to_terminal(self) -> None:
        if not self.clipboard_cells:
            return
        encoded = base64.b64encode(self._clipboard_payload().encode("utf-8")).decode("ascii")
        sys.stdout.write(f"\x1b]52;c;{encoded}\x07")
        sys.stdout.flush()

    def _export_clipboard_to_system(self) -> None:
        if not self.clipboard_cells:
            return
        row_lo, col_lo, row_hi, col_hi = self._target_range(None)
        lines: list[str] = []
        for row in range(row_lo, row_hi + 1):
            row_values = [self.sheet.get_raw(row, col) for col in range(col_lo, col_hi + 1)]
            lines.append("\t".join(row_values))
        text = "\n".join(lines)
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
        if self._load_clipboard_payload(text):
            return True
        rows = list(csv.reader(text.splitlines(), delimiter="\t"))
        if not rows:
            return False
        start_row, start_col, end_row, end_col = self._target_range(None)
        if len(rows) == 1 and len(rows[0]) == 1 and self.selection_range is None:
            self.edit_current_cell(initial_text=rows[0][0], replace=True)
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
                str(align),
                bool(manual),
            )
            for row_offset, col_offset, raw, style, text_styles, background, align, manual in cells
        ]
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

    def _set_language(self, language_name: str) -> None:
        self._save_undo_state()
        self.sheet.language = LANGUAGE_CODES.get(language_name, "en")
        self._save_global_settings()
        self.dirty = True
        self.message = language_name

    def show_settings_screen(self) -> None:
        options = ["theme", "date format", "active cell", "formula color", "protected fg", "protected bg", "language"]
        selected = 0
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
                (self._tr("theme"), self.sheet.theme_name),
                (self._tr("date_format"), self.sheet.date_format.split(":", 1)[1]),
                (self._tr("active_cell"), self.sheet.active_cell_color),
                ("Formula Colour", self.sheet.formula_foreground_color if self.sheet.formula_coloration else "off"),
                ("Protected FG", self.sheet.protected_foreground_color),
                ("Protected BG", self.sheet.protected_background_color),
                (self._tr("language"), current_language),
            ]
            label_width = min(24, max(len(label) for label, _value in rows) + 2)
            value_x = 3 + label_width
            first_row_y = 6
            row_gap = 2
            for index, (label, value) in enumerate(rows):
                y = first_row_y + (index * row_gap)
                if y >= height - 1:
                    break
                attr = self._menu_selected_attr() if index == selected else self._help_attr()
                self.stdscr.addnstr(y, 0, (" " * (width - 1)), width - 1, attr)
                self.stdscr.addnstr(y, 2, label.ljust(label_width), max(0, width - 3), attr)
                if value_x < width - 1:
                    self.stdscr.addnstr(y, value_x, str(value), width - 1 - value_x, attr)
            self.stdscr.addnstr(height - 1, 0, f" {self._tr('settings_help_3')} ".ljust(width - 1), width - 1, self._bar_attr())
            self.stdscr.refresh()
            key = self.stdscr.getch()
            if key == 27:
                self.message = self._tr("settings_closed")
                return
            if key in (curses.KEY_UP, ord("k")):
                selected = max(0, selected - 1)
                continue
            if key in (curses.KEY_DOWN, ord("j"), 9):
                selected = min(len(options) - 1, selected + 1)
                continue
            if key in (curses.KEY_LEFT, ord("h"), curses.KEY_RIGHT, ord("l"), 10, 13):
                direction = -1 if key in (curses.KEY_LEFT, ord("h")) else 1
                if key in (10, 13):
                    direction = 1
                if selected == 0:
                    current = THEMES.index(self.sheet.theme_name) if self.sheet.theme_name in THEMES else 0
                    self._set_theme(THEMES[(current + direction) % len(THEMES)])
                elif selected == 1:
                    current_style = self.sheet.date_format.split(":", 1)[1]
                    current = DATE_FORMATS.index(current_style) if current_style in DATE_FORMATS else 0
                    self._set_sheet_date_format(DATE_FORMATS[(current + direction) % len(DATE_FORMATS)])
                elif selected == 2:
                    current = ACTIVE_CELL_COLORS.index(self.sheet.active_cell_color) if self.sheet.active_cell_color in ACTIVE_CELL_COLORS else 0
                    self._set_active_cell_color(ACTIVE_CELL_COLORS[(current + direction) % len(ACTIVE_CELL_COLORS)])
                elif selected == 3:
                    current_value = self.sheet.formula_foreground_color if self.sheet.formula_coloration else "off"
                    current = FORMULA_COLOR_SETTING_OPTIONS.index(current_value) if current_value in FORMULA_COLOR_SETTING_OPTIONS else 0
                    self._set_formula_color_setting(FORMULA_COLOR_SETTING_OPTIONS[(current + direction) % len(FORMULA_COLOR_SETTING_OPTIONS)])
                elif selected == 4:
                    current = PROTECTED_COLOR_OPTIONS.index(self.sheet.protected_foreground_color) if self.sheet.protected_foreground_color in PROTECTED_COLOR_OPTIONS else 0
                    self._set_protected_colors(foreground_name=PROTECTED_COLOR_OPTIONS[(current + direction) % len(PROTECTED_COLOR_OPTIONS)])
                elif selected == 5:
                    current = PROTECTED_COLOR_OPTIONS.index(self.sheet.protected_background_color) if self.sheet.protected_background_color in PROTECTED_COLOR_OPTIONS else 0
                    self._set_protected_colors(background_name=PROTECTED_COLOR_OPTIONS[(current + direction) % len(PROTECTED_COLOR_OPTIONS)])
                else:
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
                self.current_row, self.current_col = parse_cell_reference(args[0])
                self.sheet.ensure_size(self.current_row, self.current_col)
                self._scroll_into_view()
                self.message = f"Jumped to {args[0].upper()}"
            elif name == "find":
                self._command_find(args)
            elif name == "edit":
                self._command_edit(args)
            elif name in {"blank", "protect", "unprotect"}:
                self._command_range_flag(name, args)
            elif name in {"copy", "replicate"}:
                self._command_copy(args)
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
                initial = str(self.path or DEFAULT_PATH)
                target_text = args[0] if args else self.prompt("Save as: ", initial)
                if target_text is None or not target_text.strip():
                    self.message = self._tr("save_as_cancelled")
                    return
                target = Path(target_text).expanduser()
            else:
                target = Path(args[0]).expanduser() if args else self.path or DEFAULT_PATH
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

    def _launch_menu_command(self, name: str) -> None:
        if name in {"format", "justify", "save", "help", "redo", "load"}:
            self.execute_command(name, [])
            return
        if name == "edit":
            self.show_settings_screen()
            return
        if name == "quit":
            self.execute_command("quit", [])
            return
        prompt_map = {
            "arrange": ("Arrange range [col] [desc]: ", "A1:C10 0"),
            "blank": ("Blank range (empty=current/selection): ", ""),
            "copy": ("Copy src dst: ", "A1:B3 D1"),
            "replicate": ("Replicate src dst: ", "A1:B3 D1"),
            "delete": ("Delete row|col index [n]: ", "row 1 1"),
            "execute": ("Execute file: ", ""),
            "find": ("Find text [range]: ", ""),
            "global": ("Global width n or width COL n: ", "width 14"),
            "goto": ("Goto cell: ", "A1"),
            "move": ("Move row|col a b [n]: ", "row 1 2 1"),
            "output": ("Output screen or file PATH: ", "screen"),
            "protect": ("Protect range (empty=current/selection): ", ""),
            "replace": ("Replace old new [range]: ", ""),
            "title": ("Title rows [cols]: ", "1 0"),
            "unprotect": ("Unprotect range (empty=current/selection): ", ""),
            "zap": ("Type YES to clear workspace: ", "NO"),
        }
        if name not in prompt_map:
            self.execute_command(name, [])
            return
        label, initial = prompt_map[name]
        text = self.prompt(label, initial)
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
                        if not self.sheet.is_alignment_manual(row, col):
                            self.sheet.set_alignment(row, col, "", manual=False)
                elif name == "protect":
                    self.sheet.protect(row, col)
                else:
                    self.sheet.unprotect(row, col)
        self.dirty = True
        self.message = f"{name.title()} applied to {self._range_label(row_lo, col_lo, row_hi, col_hi)}"

    def _command_copy(self, args: list[str]) -> None:
        if len(args) < 2:
            raise ValueError("copy needs source and destination ranges")
        src_lo_r, src_lo_c, src_hi_r, src_hi_c = self._parse_range_spec(args[0])
        dst_lo_r, dst_lo_c, dst_hi_r, dst_hi_c = self._parse_range_spec(args[1])
        self._save_undo_state()
        src_height = src_hi_r - src_lo_r + 1
        src_width = src_hi_c - src_lo_c + 1
        for dst_row in range(dst_lo_r, dst_hi_r + 1):
            for dst_col in range(dst_lo_c, dst_hi_c + 1):
                row_offset = (dst_row - dst_lo_r) % src_height
                col_offset = (dst_col - dst_lo_c) % src_width
                src_row = src_lo_r + row_offset
                src_col = src_lo_c + col_offset
                if self.sheet.is_protected(dst_row, dst_col):
                    continue
                raw = self.sheet.get_raw(src_row, src_col)
                shifted_raw = shift_formula_references(raw, dst_row - src_row, dst_col - src_col)
                self.sheet.set_raw(dst_row, dst_col, shifted_raw)
                self.sheet.set_format(dst_row, dst_col, self.sheet.get_format(src_row, src_col))
                self.sheet.clear_text_styles(dst_row, dst_col)
                for text_style in self.sheet.get_text_styles(src_row, src_col):
                    self.sheet.set_text_style(dst_row, dst_col, text_style, enabled=True)
                self.sheet.set_background(dst_row, dst_col, self.sheet.get_background(src_row, src_col))
                align = self.sheet.get_alignment(src_row, src_col)
                self.sheet.set_alignment(
                    dst_row,
                    dst_col,
                    align,
                    manual=self.sheet.is_alignment_manual(src_row, src_col),
                )
                if self.sheet.is_protected(src_row, src_col):
                    self.sheet.protect(dst_row, dst_col)
                else:
                    self.sheet.unprotect(dst_row, dst_col)
        self.dirty = True
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
        axis = args[0].lower()
        index = int(args[1]) - 1
        count = int(args[2]) if len(args) > 2 else 1
        self._save_undo_state()
        if axis.startswith("r"):
            self._rebuild_rows(index, -count)
            self.message = f"Deleted {count} row(s) at {index + 1}"
        else:
            self._rebuild_cols(index, -count)
            self.message = f"Deleted {count} column(s) at {index + 1}"
        self.dirty = True

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
        if style in {"clear", "clear-format", "remove-format", "none"}:
            row_lo, col_lo, row_hi, col_hi = self._target_range(args[1] if len(args) > 1 else None)
            self._save_undo_state()
            for row in range(row_lo, row_hi + 1):
                for col in range(col_lo, col_hi + 1):
                    self.sheet.set_format(row, col, "")
                    self.sheet.clear_text_styles(row, col)
                    self.sheet.set_background(row, col, "")
            self.dirty = True
            self.message = f"Formatting cleared on {self._range_label(row_lo, col_lo, row_hi, col_hi)}"
            return
        if style not in {"text", "bold", "underline", "italic", "currency", "fixed", "percent", "int", "negative", "accounting", "sci", "scientific"}:
            raise ValueError("format must be clear-format, text, bold, underline, italic, currency, date, fixed, percent, int, negative, accounting, sci, or b")
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
    ) -> str | None:
        text = list(initial)
        position = len(text)
        curses.curs_set(1)
        original_current = (self.current_row, self.current_col)
        ref_row, ref_col = formula_origin if formula_origin is not None else original_current
        inserted_ref: tuple[int, int] | None = None
        while True:
            height, width = self.stdscr.getmaxyx()
            if formula_origin is not None:
                self.current_row = ref_row
                self.current_col = ref_col
                self._scroll_into_view()
            if formula_origin is not None or help_lines:
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
            prompt_attr = self._bar_attr(bold=True)
            self.stdscr.addnstr(height - 1, 0, (" " * (width - 1)), width - 1, prompt_attr)
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
                return result
            if key == 27:
                curses.curs_set(0)
                if formula_origin is not None:
                    self.current_row, self.current_col = original_current
                return None
            if formula_origin is not None and key in (curses.KEY_UP, curses.KEY_DOWN, curses.KEY_LEFT, curses.KEY_RIGHT):
                row_delta, col_delta = {
                    curses.KEY_UP: (-1, 0),
                    curses.KEY_DOWN: (1, 0),
                    curses.KEY_LEFT: (0, -1),
                    curses.KEY_RIGHT: (0, 1),
                }[key]
                if self._formula_reference_context("".join(text), position) or inserted_ref is not None:
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
            if function_name:
                signature = FORMULA_SIGNATURES.get(function_name, f"{function_name}(...)")
                arguments = FORMULA_ARGUMENT_NAMES.get(function_name, [])
                if arguments:
                    argument_index = min(max(0, argument_count - 1), len(arguments) - 1)
                    arg_text = f"   arg {argument_count}: {arguments[argument_index]}"
                else:
                    arg_text = ""
                text = f" Fx {ref}: {raw}   => {value}   {signature}{arg_text}{error_suffix}"
            else:
                text = f" Fx {ref}: {raw}   => {value}{error_suffix}"
        else:
            display = self._display_value(self.current_row, self.current_col)
            text = f" Cell {ref}: {raw or display or '(empty)'}"
        return text[: width - 1].ljust(width - 1)

    def _display_value(self, row: int, col: int) -> str:
        try:
            text = self.evaluator.display_value(row, col)
        except FormulaError as exc:
            return f"#ERR {exc}"
        return self._apply_format(text, self.sheet.get_format(row, col))

    def _cell_text(self, row: int, col: int, width: int) -> str:
        text = self._display_value(row, col)
        text = text[:width]
        align = self.sheet.get_alignment(row, col)
        if align == "right":
            return text.rjust(width)
        if align == "center":
            return text.center(width)
        return text.ljust(width)

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
            if self.sheet.get_background(row, next_col):
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
        text_style_attr = self._text_style_attr(self.sheet.get_text_styles(row, col))
        background_name = self.sheet.get_background(row, col)
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
            background_color = BACKGROUND_COLOR_MAP.get(background_name, -1)
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
        text_style_attr = self._text_style_attr(self.sheet.get_text_styles(row, col))
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

    def _active_cell_attr(self, row: int, col: int) -> int:
        attr = self._selection_cell_attr(row, col)
        return attr | curses.A_BOLD

    def _text_style_attr(self, styles: set[str]) -> int:
        attr = curses.A_NORMAL
        if "bold" in styles:
            attr |= curses.A_BOLD
        if "underline" in styles:
            attr |= curses.A_UNDERLINE
        if "italic" in styles:
            attr |= getattr(curses, "A_ITALIC", curses.A_DIM)
        return attr

    def _selection_foreground_color(self) -> int:
        if self.sheet.active_cell_color in {"white", "yellow", "lightblue", "lightgrey"}:
            return curses.COLOR_BLACK
        return curses.COLOR_WHITE

    def _formula_foreground_color(self) -> int:
        return self._named_color(self.sheet.formula_foreground_color, curses.COLOR_GREEN)

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

    def _background_foreground_color(self, background_name: str, is_formula: bool, row: int, col: int, style: str) -> int:
        numeric_value = self._cell_numeric_value(row, col)
        if style == "negative" and numeric_value is not None and numeric_value < 0:
            return curses.COLOR_RED
        if is_formula and background_name != "green":
            return curses.COLOR_GREEN
        if background_name in {"cyan", "green", "white", "yellow"}:
            return curses.COLOR_BLACK
        return curses.COLOR_WHITE

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

    def _bar_attr(self, bold: bool = False) -> int:
        attr = curses.A_NORMAL
        if bold:
            attr |= curses.A_BOLD
        if self.colors_ready:
            attr |= curses.color_pair(COLOR_PAIR_BAR)
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
        return stamp.strftime("%y/%m/%d %H:%M")

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
            col_width = self.sheet.get_column_width(col)
            if x + col_width > total_width:
                break
            columns.append((col, x, col_width))
            x += col_width
            col += 1
        return columns

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
        row = self.row_offset + (y - top_grid_row)
        if row >= self.sheet.rows:
            return None
        if x < row_header_width:
            return None
        for col, col_x, col_width in visible_columns:
            if col_x <= x < col_x + col_width - 1:
                return row, col
        return None

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

    def _parse_range_spec(self, spec: str) -> tuple[int, int, int, int]:
        return parse_range_spec(spec, self.current_row, self.current_col, self.sheet.rows, self.sheet.cols)

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
        selected = options.index(default_option) if default_option in options else 0
        typed = ""
        curses.curs_set(0)
        while True:
            self.draw()
            query = f" [{typed}]" if typed else ""
            title_text = f"{title}{query}: "
            hint = footer_hint or " arrows/type/Enter/Esc "
            rows: list[list[tuple[int, str]]] = [[]]
            row_widths = [len(title_text)]
            for index, option in enumerate(options):
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
            description = descriptions.get(options[selected], "") if descriptions else ""
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
                    attr = self._help_attr() if item_index == -1 else (self._menu_selected_attr() if item_index == selected else self._bar_attr())
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
                return options[selected]
            if key == curses.KEY_LEFT:
                selected = (selected - 1) % len(options)
                typed = ""
                continue
            if key == curses.KEY_RIGHT:
                selected = (selected + 1) % len(options)
                typed = ""
                continue
            if key in (curses.KEY_BACKSPACE, 127):
                typed = typed[:-1]
                continue
            if 32 <= key <= 126:
                typed += chr(key).lower()
                for index, option in enumerate(options):
                    if option.startswith(typed):
                        selected = index
                        break

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
                f"  s sort:{sort_mode}  p path  n mkdir  type=filter  Esc cancel "
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

            footer = " Enter=open/select   s=sort   p=path   n=mkdir   .=hidden   type to filter   Backspace=clear   Esc=cancel "
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
                if self._load_clipboard_payload(pasted):
                    return True
                if pasted:
                    self.edit_current_cell(initial_text=pasted, replace=True)
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
        if self.sheet.date_format.startswith("date") and parse_date_text(text, self.sheet.date_format) is not None:
            try:
                return format_date_text(text, self.sheet.date_format)
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
        top_grid_row, _grid_height, _row_header_width, _visible_columns = self._grid_layout(height, width)
        visible_rows = max(1, height - top_grid_row - 1)
        if self.current_row < self.row_offset:
            self.row_offset = self.current_row
        elif self.current_row >= self.row_offset + visible_rows:
            self.row_offset = self.current_row - visible_rows + 1
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
            title_rows=self.sheet.title_rows,
            title_cols=self.sheet.title_cols,
            theme_name=self.sheet.theme_name,
            date_format=self.sheet.date_format,
            active_cell_color=self.sheet.active_cell_color,
            language=self.sheet.language,
        )
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
            title_rows=self.sheet.title_rows,
            title_cols=self.sheet.title_cols,
            theme_name=self.sheet.theme_name,
            date_format=self.sheet.date_format,
            active_cell_color=self.sheet.active_cell_color,
            language=self.sheet.language,
        )
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
            title_rows=self.sheet.title_rows,
            title_cols=self.sheet.title_cols,
            theme_name=self.sheet.theme_name,
            date_format=self.sheet.date_format,
            active_cell_color=self.sheet.active_cell_color,
            language=self.sheet.language,
        )
        for new_row, old_row in enumerate(order):
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
            title_rows=self.sheet.title_rows,
            title_cols=self.sheet.title_cols,
            theme_name=self.sheet.theme_name,
            date_format=self.sheet.date_format,
            active_cell_color=self.sheet.active_cell_color,
            language=self.sheet.language,
        )
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
    executable_dir = Path(sys.argv[0]).resolve().parent
    settings_path = executable_dir / "tui-ss-settings.toml"
    return curses.wrapper(lambda stdscr: _run_multiple(stdscr, paths, settings_path))


if __name__ == "__main__":
    raise SystemExit(main())
