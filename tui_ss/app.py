#!/usr/bin/env python3
"""Curses spreadsheet application."""

from __future__ import annotations

import argparse
import curses
import shlex
from datetime import datetime
from pathlib import Path

from .commands import (
    COMMAND_HELP_LINES,
    COMMAND_MENU_OPTIONS,
    FORMULA_HELP_LINES,
    HELP_TOPICS,
    KEY_HELP_LINES,
    parse_command,
)
from .formulas import Evaluator, FormulaError
from .model import Spreadsheet, column_label, parse_cell_reference
from .storage import load_sheet, save_sheet

APP_NAME = "tui-ss"
DEFAULT_PATH = Path.home() / "scripts" / "tui-ss" / "sheets" / "autosave.tss"
THEMES = ["white", "cyan", "yellow", "magenta", "blue", "purple"]
FORMAT_STYLES = ["text", "currency", "background", "fixed", "percent", "int", "sci"]
CURRENCY_SYMBOLS = ["£", "€", "$", "¥", "₹"]
BACKGROUND_COLORS = ["none", "blue", "cyan", "green", "magenta", "red", "yellow", "white"]
JUSTIFY_OPTIONS = ["l", "c", "r"]
COLOR_PAIR_TEXT = 1
COLOR_PAIR_FORMULA = 2
COLOR_PAIR_HEADER = 3
COLOR_PAIR_BAR = 4
COLOR_PAIR_GRID = 5
COLOR_PAIR_MENU_SELECTED = 6
COLOR_PAIR_GRID_ROW = 7
COLOR_PAIR_SELECTION = 8
COLOR_PAIR_ROW_HEADER = 9
THEME_COLOR_MAP = {
    "white": curses.COLOR_WHITE,
    "cyan": curses.COLOR_CYAN,
    "yellow": curses.COLOR_YELLOW,
    "magenta": curses.COLOR_MAGENTA,
    "blue": curses.COLOR_BLUE,
    "purple": curses.COLOR_MAGENTA,
}
CUSTOM_PURPLE_COLOR_ID = 16
BACKGROUND_COLOR_MAP = {
    "blue": curses.COLOR_BLUE,
    "cyan": curses.COLOR_CYAN,
    "green": curses.COLOR_GREEN,
    "magenta": curses.COLOR_MAGENTA,
    "red": curses.COLOR_RED,
    "yellow": curses.COLOR_YELLOW,
    "white": curses.COLOR_WHITE,
}


def build_stamp() -> str:
    try:
        return datetime.fromtimestamp(Path(__file__).stat().st_mtime).strftime("%y%m%d-%H:%M")
    except OSError:
        return datetime.now().strftime("%y%m%d-%H:%M")


def parse_cell_or_current(token: str | None, row: int, col: int) -> tuple[int, int]:
    if not token:
        return row, col
    return parse_cell_reference(token)


def parse_range_spec(spec: str, current_row: int, current_col: int) -> tuple[int, int, int, int]:
    token = spec.strip().upper() if spec else ""
    if not token or token == ".":
        return current_row, current_col, current_row, current_col
    if ":" not in token:
        row, col = parse_cell_reference(token)
        return row, col, row, col
    start_text, end_text = token.split(":", 1)
    start_row, start_col = parse_cell_reference(start_text)
    end_row, end_col = parse_cell_reference(end_text)
    row_lo, row_hi = sorted((start_row, end_row))
    col_lo, col_hi = sorted((start_col, end_col))
    return row_lo, col_lo, row_hi, col_hi


def should_auto_right_align(raw: str) -> bool:
    text = raw.strip()
    if not text:
        return False
    if text.startswith("="):
        return True
    try:
        float(text.replace(",", ""))
        return True
    except ValueError:
        return False


class SpreadsheetApp:
    def __init__(self, stdscr, path: Path | None = None) -> None:
        self.stdscr = stdscr
        self.sheet = Spreadsheet(rows=100, cols=52)
        self.evaluator = Evaluator(self.sheet)
        self.current_row = 0
        self.current_col = 0
        self.row_offset = 0
        self.col_offset = 0
        self.message = "Press / for SuperCalc-style commands, Enter to edit, Tab to move right."
        self.path = path
        self.running = True
        self.colors_ready = False
        self.dynamic_color_pairs: dict[tuple[int, int], int] = {}
        self.next_dynamic_pair = 20
        self.selection_anchor: tuple[int, int] | None = None
        self.selection_range: tuple[int, int, int, int] | None = None
        self.mouse_dragging = False
        self.dirty = False

    def run(self) -> int:
        curses.curs_set(0)
        self.stdscr.keypad(True)
        mouse_events = curses.BUTTON1_PRESSED | curses.BUTTON1_CLICKED
        mouse_events |= getattr(curses, "BUTTON1_RELEASED", 0)
        mouse_events |= getattr(curses, "REPORT_MOUSE_POSITION", 0)
        curses.mousemask(mouse_events)
        curses.mouseinterval(0)
        self._setup_colors()
        while self.running:
            self.draw()
            key = self.stdscr.getch()
            self.handle_key(key)
        return 0

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
        curses.init_pair(COLOR_PAIR_FORMULA, curses.COLOR_GREEN, -1)
        curses.init_pair(COLOR_PAIR_HEADER, text_color, curses.COLOR_BLACK)
        curses.init_pair(COLOR_PAIR_BAR, text_color, -1)
        curses.init_pair(COLOR_PAIR_GRID, curses.COLOR_BLACK, -1)
        curses.init_pair(COLOR_PAIR_GRID_ROW, text_color, -1)
        curses.init_pair(COLOR_PAIR_MENU_SELECTED, curses.COLOR_BLACK, text_color)
        curses.init_pair(COLOR_PAIR_SELECTION, curses.COLOR_BLACK, text_color)
        curses.init_pair(COLOR_PAIR_ROW_HEADER, text_color, curses.COLOR_BLACK)

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

    def draw(self) -> None:
        self.stdscr.erase()
        height, width = self.stdscr.getmaxyx()
        top_grid_row, grid_height, row_header_width, visible_columns = self._grid_layout(height, width)
        self.stdscr.addnstr(0, 0, self._title_line(width), width - 1, self._bar_attr(bold=True))
        self._draw_grid(top_grid_row - 1, grid_height, row_header_width, visible_columns)
        self._draw_column_headers(top_grid_row - 1, visible_columns)

        for screen_row in range(grid_height):
            row = self.row_offset + screen_row
            if row >= self.sheet.rows:
                break
            y = top_grid_row + screen_row
            self.stdscr.addnstr(y, 0, f"{row + 1:>5}", row_header_width - 1, self._row_header_attr())
            for col, x, col_width in visible_columns:
                text = self._cell_text(row, col, col_width - 1)
                in_selection = self._cell_in_selection(row, col)
                attr = self._cell_attr(row, col)
                if in_selection:
                    attr = self._selection_cell_attr(row, col)
                if (row, col) == (self.current_row, self.current_col):
                    if in_selection:
                        attr |= curses.A_BOLD
                    else:
                        attr |= curses.A_REVERSE
                if row < self.sheet.title_rows or col < self.sheet.title_cols:
                    attr |= curses.A_BOLD
                self.stdscr.addnstr(y, x, text, col_width - 1, attr)

        self.stdscr.addnstr(height - 2, 0, self._formula_bar(width), width - 1, self._bar_attr())
        self.stdscr.addnstr(height - 1, 0, self.message.ljust(width - 1), width - 1, self._bar_attr(bold=True))
        self.stdscr.refresh()

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
        elif key in (curses.KEY_RIGHT, ord("l"), 9):
            self.move(0, 1)
        elif key in (getattr(curses, "KEY_SR", -1),):
            self.extend_selection(-1, 0)
        elif key in (getattr(curses, "KEY_SF", -1),):
            self.extend_selection(1, 0)
        elif key in (getattr(curses, "KEY_SLEFT", -1),):
            self.extend_selection(0, -1)
        elif key in (getattr(curses, "KEY_SRIGHT", -1),):
            self.extend_selection(0, 1)
        elif key in (10, 13):
            self.edit_current_cell()
        elif key == curses.KEY_DC:
            self.clear_current_cell()
        elif key == curses.KEY_MOUSE:
            self._handle_mouse()
        elif key == ord("/"):
            self.run_command_prompt()
        elif key == 27:
            if not self._handle_alt_sequence():
                self.message = "Ready."
        elif key == curses.KEY_NPAGE:
            self.move(10, 0)
        elif key == curses.KEY_PPAGE:
            self.move(-10, 0)
        elif 32 <= key <= 126:
            self.edit_current_cell(initial_text=chr(key), replace=True)

    def move(self, row_delta: int, col_delta: int) -> None:
        self.selection_anchor = None
        self.selection_range = None
        self.mouse_dragging = False
        self.current_row = max(0, min(self.sheet.rows - 1, self.current_row + row_delta))
        self.current_col = max(0, min(self.sheet.cols - 1, self.current_col + col_delta))
        self._scroll_into_view()
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

    def _handle_mouse(self) -> None:
        try:
            _id, mouse_x, mouse_y, _z, state = curses.getmouse()
        except curses.error:
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

    def edit_current_cell(self, initial_text: str = "", replace: bool = False) -> None:
        if self.sheet.is_protected(self.current_row, self.current_col):
            self.message = "Cell is protected."
            return
        raw = "" if replace else self.sheet.get_raw(self.current_row, self.current_col)
        edited = self.prompt(
            f"Edit {column_label(self.current_col)}{self.current_row + 1}: ",
            initial_text if replace else (initial_text or raw),
        )
        if edited is None:
            self.message = "Edit cancelled."
            return
        stored_ref = f"{column_label(self.current_col)}{self.current_row + 1}"
        self.sheet.set_raw(self.current_row, self.current_col, edited)
        self._apply_default_alignment(self.current_row, self.current_col, edited)
        self.dirty = True
        if self.current_row >= self.sheet.rows - 1:
            self.sheet.ensure_size(self.current_row + 1, self.current_col)
        self.current_row += 1
        self._scroll_into_view()
        self.message = f"Stored {stored_ref}; ready for {column_label(self.current_col)}{self.current_row + 1}"

    def clear_current_cell(self) -> None:
        if self.sheet.is_protected(self.current_row, self.current_col):
            self.message = "Cell is protected."
            return
        row = self.current_row
        col = self.current_col
        cell_ref = f"{column_label(col)}{row + 1}"
        self.sheet.clear(row, col)
        if not self.sheet.is_alignment_manual(row, col):
            self.sheet.set_alignment(row, col, "", manual=False)
        self.dirty = True
        self.message = f"Cleared {cell_ref}"

    def run_command_prompt(self) -> None:
        selected = self._choose_from_menu("Slash", COMMAND_MENU_OPTIONS, default_option="edit")
        if selected is None:
            self.message = "Command cancelled."
            return
        self._launch_menu_command(selected)

    def _confirm_quit(self) -> None:
        if not self.dirty:
            self.running = False
            self.message = "Bye."
            return
        choice = self._choose_from_menu("Unsaved changes. Quit?", ["yes", "no"], default_option="yes")
        if choice is None or choice == "no":
            self.message = "Quit cancelled."
            return
        self.running = False
        self.message = "Exited without saving."

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
            elif name == "theme":
                self._command_theme(args)
            elif name == "global":
                self._command_global(args)
            elif name == "title":
                self._command_title(args)
            elif name == "output":
                self._command_output(args)
            elif name == "execute":
                self._command_execute(args)
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
            self.message = "Save cancelled."
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
                    self.message = "Save as cancelled."
                    return
                target = Path(target_text).expanduser()
            else:
                target = Path(args[0]).expanduser() if args else self.path or DEFAULT_PATH
            save_sheet(self.sheet, target)
            self.path = target
            self.dirty = False
            self.message = f"Saved {target}"
            return
        if not args:
            raise ValueError("load needs a path")
        target = Path(args[0]).expanduser()
        self.sheet = load_sheet(target)
        self.evaluator = Evaluator(self.sheet)
        self._refresh_theme_colors()
        self.path = target
        self.current_row = 0
        self.current_col = 0
        self.row_offset = 0
        self.col_offset = 0
        self.dirty = False
        self.message = f"Loaded {target}"

    def _launch_menu_command(self, name: str) -> None:
        if name in {"format", "save", "theme", "help"}:
            self.execute_command(name, [])
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
            "edit": ("Edit [cell] value: ", ""),
            "execute": ("Execute file: ", ""),
            "global": ("Global width n or width COL n: ", "width 14"),
            "goto": ("Goto cell: ", "A1"),
            "insert": ("Insert row|col index [n]: ", "row 1 1"),
            "justify": ("Justify l|c|r [range]: ", "r"),
            "load": ("Load file: ", str(self.path or DEFAULT_PATH)),
            "move": ("Move row|col a b [n]: ", "row 1 2 1"),
            "output": ("Output screen or file PATH: ", "screen"),
            "protect": ("Protect range (empty=current/selection): ", ""),
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

    def _command_help(self, args: list[str]) -> None:
        if args:
            topic = args[0].lower()
        else:
            topic = self._choose_from_menu("Help", HELP_TOPICS, default_option="commands")
            if topic is None:
                self.message = "Help cancelled."
                return
        topic_map = {
            "commands": ("Commands", COMMAND_HELP_LINES),
            "keys": ("Keys", KEY_HELP_LINES),
            "formulas": ("Formulas", FORMULA_HELP_LINES),
            "formula": ("Formulas", FORMULA_HELP_LINES),
        }
        if topic not in topic_map:
            raise ValueError("help topics: commands, keys, formulas")
        title, lines = topic_map[topic]
        self._show_text_page(title, lines)
        self.message = f"Help: {title.lower()}"

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

    def _command_edit(self, args: list[str]) -> None:
        if not args:
            self.edit_current_cell()
            return
        row, col = self.current_row, self.current_col
        value_args = args
        if args[0][0].isalpha() and any(char.isdigit() for char in args[0]):
            row, col = parse_cell_reference(args[0])
            value_args = args[1:]
        if self.sheet.is_protected(row, col):
            raise ValueError("cell is protected")
        value = " ".join(value_args)
        self.sheet.set_raw(row, col, value)
        self._apply_default_alignment(row, col, value)
        self.dirty = True
        self.message = f"Stored {column_label(col)}{row + 1}"

    def _command_range_flag(self, name: str, args: list[str]) -> None:
        row_lo, col_lo, row_hi, col_hi = self._target_range(args[0] if args else None)
        for row in range(row_lo, row_hi + 1):
            for col in range(col_lo, col_hi + 1):
                if name == "blank":
                    if not self.sheet.is_protected(row, col):
                        self.sheet.clear(row, col)
                elif name == "protect":
                    self.sheet.protect(row, col)
                else:
                    self.sheet.unprotect(row, col)
        self.dirty = True
        self.message = f"{name.title()} applied to {self._range_label(row_lo, col_lo, row_hi, col_hi)}"

    def _command_copy(self, args: list[str]) -> None:
        if len(args) < 2:
            raise ValueError("copy needs source and destination ranges")
        src_lo_r, src_lo_c, src_hi_r, src_hi_c = parse_range_spec(args[0], self.current_row, self.current_col)
        dst_r, dst_c, _, _ = parse_range_spec(args[1], self.current_row, self.current_col)
        for row_offset in range(src_hi_r - src_lo_r + 1):
            for col_offset in range(src_hi_c - src_lo_c + 1):
                src_row = src_lo_r + row_offset
                src_col = src_lo_c + col_offset
                dst_row = dst_r + row_offset
                dst_col = dst_c + col_offset
                if self.sheet.is_protected(dst_row, dst_col):
                    continue
                raw = self.sheet.get_raw(src_row, src_col)
                self.sheet.set_raw(dst_row, dst_col, raw)
                self.sheet.set_format(dst_row, dst_col, self.sheet.get_format(src_row, src_col))
                self.sheet.set_background(dst_row, dst_col, self.sheet.get_background(src_row, src_col))
                if self.sheet.is_protected(src_row, src_col):
                    self.sheet.protect(dst_row, dst_col)
                else:
                    self.sheet.unprotect(dst_row, dst_col)
        self.dirty = True
        self.message = f"Copied {args[0].upper()} to {args[1].upper()}"

    def _command_arrange(self, args: list[str]) -> None:
        if not args:
            raise ValueError("arrange needs a range")
        row_lo, col_lo, row_hi, col_hi = parse_range_spec(args[0], self.current_row, self.current_col)
        sort_offset = int(args[1]) if len(args) > 1 else 0
        descending = len(args) > 2 and args[2].lower().startswith("d")
        sort_col = col_lo + sort_offset
        records: list[list[tuple[str, str, str, bool]]] = []
        for row in range(row_lo, row_hi + 1):
            record = []
            for col in range(col_lo, col_hi + 1):
                record.append(
                    (
                        self.sheet.get_raw(row, col),
                        self.sheet.get_format(row, col),
                        self.sheet.get_background(row, col),
                        self.sheet.is_protected(row, col),
                    )
                )
            records.append(record)
        def sort_key(record: list[tuple[str, str, str, bool]]) -> str:
            index = max(0, min(len(record) - 1, sort_col - col_lo))
            return record[index][0]
        records.sort(key=sort_key, reverse=descending)
        for row_index, record in enumerate(records, start=row_lo):
            for col_index, (raw, style, background, protected) in enumerate(record, start=col_lo):
                self.sheet.set_raw(row_index, col_index, raw)
                self.sheet.set_format(row_index, col_index, style)
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
        if axis.startswith("r"):
            self._rebuild_rows(index, -count)
            self.message = f"Deleted {count} row(s) at {index + 1}"
        else:
            self._rebuild_cols(index, -count)
            self.message = f"Deleted {count} column(s) at {index + 1}"
        self.dirty = True

    def _command_insert(self, args: list[str]) -> None:
        axis = args[0].lower()
        index = int(args[1]) - 1
        count = int(args[2]) if len(args) > 2 else 1
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
        if style not in {"text", "currency", "fixed", "percent", "int", "sci", "scientific"}:
            raise ValueError("format must be text, currency, fixed, percent, int, sci, or b")
        format_value = "" if style == "text" else style
        range_arg = "."
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
            raise ValueError("justify must be l, c, or r")
        resolved = align_map[align]
        row_lo, col_lo, row_hi, col_hi = self._target_range(args[1] if len(args) > 1 else None)
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
            self.sheet.theme_name = requested
        else:
            selected = self._choose_from_menu("Theme", THEMES, default_option=self.sheet.theme_name)
            if selected is None:
                self.message = "Theme cancelled."
                return
            self.sheet.theme_name = selected
        self._refresh_theme_colors()
        self.dirty = True
        self.message = f"Theme set to {self.sheet.theme_name}"

    def _command_global(self, args: list[str]) -> None:
        if not args or args[0].lower() != "width":
            raise ValueError("global supports: width N or width COL N")
        if len(args) >= 3:
            _, col = parse_cell_reference(f"{args[1]}1")
            width = max(8, int(args[2]))
            self.sheet.set_column_width(col, width)
            self.dirty = True
            self.message = f"Width for {column_label(col)} set to {width}"
            return
        self.sheet.column_width = max(8, int(args[1]))
        self.dirty = True
        self.message = f"Default column width set to {self.sheet.column_width}"

    def _command_title(self, args: list[str]) -> None:
        self.sheet.title_rows = max(0, int(args[0])) if args else 0
        self.sheet.title_cols = max(0, int(args[1])) if len(args) > 1 else 0
        self.dirty = True
        self.message = f"Title freeze rows={self.sheet.title_rows} cols={self.sheet.title_cols}"

    def _command_output(self, args: list[str]) -> None:
        if not args or args[0].lower() == "screen":
            snapshot = self.render_text_snapshot().splitlines() or ["[No populated cells]"]
            self._show_text_page("Output", snapshot)
            self.message = "Output shown on screen."
            return
        if args[0].lower() != "file" or len(args) < 2:
            raise ValueError("output needs: screen or file PATH")
        target = Path(args[1]).expanduser()
        if target.suffix.lower() == ".csv":
            save_sheet(self.sheet, target)
        else:
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(self.render_text_snapshot(), encoding="utf-8")
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

    def prompt(self, label: str, initial: str, help_lines: list[str] | None = None) -> str | None:
        height, width = self.stdscr.getmaxyx()
        text = list(initial)
        position = len(text)
        curses.curs_set(1)
        while True:
            if help_lines:
                self.draw()
                panel_height = min(len(help_lines) + 2, max(4, height - 4))
                self._draw_help_panel(2, panel_height, width)
            display = "".join(text)
            prompt_attr = self._bar_attr(bold=True)
            self.stdscr.addnstr(height - 1, 0, (" " * (width - 1)), width - 1, prompt_attr)
            self.stdscr.addnstr(height - 1, 0, f"{label}{display}", width - 1, prompt_attr)
            cursor_x = min(width - 2, len(label) + position)
            self.stdscr.move(height - 1, cursor_x)
            self.stdscr.refresh()
            key = self.stdscr.getch()
            if key in (10, 13):
                curses.curs_set(0)
                return "".join(text)
            if key == 27:
                curses.curs_set(0)
                return None
            if key in (curses.KEY_BACKSPACE, 127):
                if position > 0:
                    del text[position - 1]
                    position -= 1
                continue
            if key == curses.KEY_DC and position < len(text):
                del text[position]
                continue
            if key == curses.KEY_LEFT and position > 0:
                position -= 1
                continue
            if key == curses.KEY_RIGHT and position < len(text):
                position += 1
                continue
            if 32 <= key <= 126:
                text.insert(position, chr(key))
                position += 1

    def render_text_snapshot(self) -> str:
        lines = []
        for row in range(self.sheet.rows):
            row_values = []
            for col in range(self.sheet.cols):
                value = self._display_value(row, col)
                if value:
                    row_values.append(f"{column_label(col)}{row + 1}={value}")
            if row_values:
                lines.append(" | ".join(row_values))
        return "\n".join(lines) + ("\n" if lines else "")

    def _title_line(self, width: int) -> str:
        target = str(self.path) if self.path else "[unsaved]"
        current_width = self.sheet.get_column_width(self.current_col)
        dirty_flag = "*" if self.dirty else ""
        title = f" {APP_NAME}{dirty_flag}  {target}  defw={self.sheet.column_width}  {column_label(self.current_col)}w={current_width}  build={build_stamp()} "
        return title.ljust(width - 1)

    def _options_line(self, width: int) -> str:
        current_format = self.sheet.get_format(self.current_row, self.current_col) or "text"
        current_align = self.sheet.get_alignment(self.current_row, self.current_col) or "left"
        text = (
            " Formats: /F text /F currency /F fixed /F percent /F int /F sci"
            f"  |  Theme: /V [Enter]=cycle /V white|cyan|yellow|magenta|blue"
            f"  |  Justify: /J l /J c /J r"
            f"  |  Width: /G width 14  or  /G width {column_label(self.current_col)} 18"
            f"  |  Cell format={current_format} align={current_align} "
        )
        return text[: width - 1].ljust(width - 1)

    def _formula_bar(self, width: int) -> str:
        ref = f"{column_label(self.current_col)}{self.current_row + 1}"
        raw = self.sheet.get_raw(self.current_row, self.current_col)
        value = self._display_value(self.current_row, self.current_col)
        flags = []
        if self.sheet.is_protected(self.current_row, self.current_col):
            flags.append("PROT")
        cell_format = self.sheet.get_format(self.current_row, self.current_col)
        if cell_format:
            flags.append(cell_format.upper())
        cell_background = self.sheet.get_background(self.current_row, self.current_col)
        if cell_background:
            flags.append(f"BG={cell_background.upper()}")
        meta = f"[{' '.join(flags)}]" if flags else ""
        selection = f" sel={self._selection_label()}" if self.selection_range else ""
        text = f" {ref} raw={raw or ' '} value={value or ' '} {meta}{selection}"
        return text[: width - 1].ljust(width - 1)

    def _display_value(self, row: int, col: int) -> str:
        try:
            text = self.evaluator.display_value(row, col)
        except FormulaError as exc:
            return f"#ERR {exc}"
        return self._apply_format(text, self.sheet.get_format(row, col))

    def _cell_text(self, row: int, col: int, width: int) -> str:
        text = self._display_value(row, col)
        if self.sheet.is_protected(row, col) and text:
            text = f"!{text}"
        text = text[:width]
        align = self.sheet.get_alignment(row, col)
        if align == "right":
            return text.rjust(width)
        if align == "center":
            return text.center(width)
        return text.ljust(width)

    def _cell_attr(self, row: int, col: int) -> int:
        attr = curses.A_NORMAL
        if not self.colors_ready:
            return attr
        raw = self.sheet.get_raw(row, col)
        background_name = self.sheet.get_background(row, col)
        if background_name:
            background_color = BACKGROUND_COLOR_MAP.get(background_name, -1)
            foreground_color = self._background_foreground_color(background_name, raw.startswith("="))
            pair_number = self._ensure_color_pair(foreground_color, background_color)
            if pair_number is not None:
                attr |= curses.color_pair(pair_number)
                if raw.startswith("="):
                    attr |= curses.A_BOLD
                return attr
        if raw.startswith("="):
            return attr | curses.color_pair(COLOR_PAIR_FORMULA) | curses.A_BOLD
        return attr | curses.color_pair(COLOR_PAIR_TEXT)

    def _selection_cell_attr(self, row: int, col: int) -> int:
        attr = curses.A_NORMAL
        if not self.colors_ready:
            return attr | curses.A_REVERSE
        raw = self.sheet.get_raw(row, col)
        selection_background = self._selection_background_color()
        if raw.startswith("="):
            pair_number = self._ensure_color_pair(curses.COLOR_GREEN, selection_background)
            if pair_number is not None:
                return attr | curses.color_pair(pair_number) | curses.A_BOLD
            return attr | curses.color_pair(COLOR_PAIR_SELECTION) | curses.A_BOLD
        foreground = curses.COLOR_BLACK if selection_background in {
            curses.COLOR_CYAN,
            curses.COLOR_GREEN,
            curses.COLOR_WHITE,
            curses.COLOR_YELLOW,
        } else curses.COLOR_WHITE
        pair_number = self._ensure_color_pair(foreground, selection_background)
        if pair_number is not None:
            return attr | curses.color_pair(pair_number)
        return attr | curses.color_pair(COLOR_PAIR_SELECTION)

    def _selection_background_color(self) -> int:
        background_name = self.sheet.get_background(self.current_row, self.current_col)
        if background_name:
            return BACKGROUND_COLOR_MAP.get(background_name, self._theme_text_color())
        return self._theme_text_color()

    def _cell_in_selection(self, row: int, col: int) -> bool:
        if self.selection_range is None:
            return False
        row_lo, col_lo, row_hi, col_hi = self.selection_range
        return row_lo <= row <= row_hi and col_lo <= col <= col_hi

    def _background_foreground_color(self, background_name: str, is_formula: bool) -> int:
        if is_formula and background_name != "green":
            return curses.COLOR_GREEN
        if background_name in {"cyan", "green", "white", "yellow"}:
            return curses.COLOR_BLACK
        return curses.COLOR_WHITE

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
        bottom_bars = 2
        top_grid_row = 2
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
        return self._range_label(row_lo, col_lo, row_hi, col_hi)

    def _target_range(self, spec: str | None) -> tuple[int, int, int, int]:
        if spec and spec.strip() and spec.strip() != ".":
            return parse_range_spec(spec, self.current_row, self.current_col)
        if self.selection_range is not None:
            return self.selection_range
        return self.current_row, self.current_col, self.current_row, self.current_col

    def _choose_from_menu(self, title: str, options: list[str], default_option: str | None = None) -> str | None:
        height, width = self.stdscr.getmaxyx()
        selected = options.index(default_option) if default_option in options else 0
        typed = ""
        curses.curs_set(0)
        while True:
            self.draw()
            title_text = f"{title}: "
            hint = " arrows/type/Enter/Esc "
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
            start_y = max(0, height - menu_rows)
            for offset in range(menu_rows):
                self.stdscr.addnstr(start_y + offset, 0, (" " * (width - 1)), width - 1, self._bar_attr(bold=True))
            for offset, row_items in enumerate(rows):
                y = start_y + offset
                x = 0
                if offset == 0:
                    self.stdscr.addnstr(y, x, title_text, min(len(title_text), width - 1), self._bar_attr(bold=True))
                    x += len(title_text)
                for item_index, chip in row_items:
                    if x and x < width - 1:
                        self.stdscr.addch(y, x, ord(" "), self._bar_attr())
                        x += 1
                    attr = self._help_attr() if item_index == -1 else (self._menu_selected_attr() if item_index == selected else self._bar_attr())
                    if x < width - 1:
                        self.stdscr.addnstr(y, x, chip, min(len(chip), width - 1 - x), attr)
                    x += len(chip)
            self.stdscr.refresh()
            key = self.stdscr.getch()
            if key in (27,):
                return None
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
            if 32 <= key <= 126:
                typed += chr(key).lower()
                for index, option in enumerate(options):
                    if option.startswith(typed):
                        selected = index
                        break

    def _menu_selected_attr(self) -> int:
        attr = curses.A_BOLD
        if self.colors_ready:
            attr |= curses.color_pair(COLOR_PAIR_MENU_SELECTED)
        else:
            attr |= curses.A_REVERSE
        return attr

    def _handle_alt_sequence(self) -> bool:
        self.stdscr.nodelay(True)
        try:
            next_key = self.stdscr.getch()
        finally:
            self.stdscr.nodelay(False)
        if next_key in (ord("="), ord("+")):
            self._sum_above_into_current_cell()
            return True
        return False

    def _sum_above_into_current_cell(self) -> None:
        if self.sheet.is_protected(self.current_row, self.current_col):
            self.message = "Cell is protected."
            return
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
        if not style or not text:
            return text
        try:
            number = float(text)
        except ValueError:
            return text
        if style == "currency" or style.startswith("currency:"):
            symbol = style.split(":", 1)[1] if ":" in style else "£"
            return f"{symbol}{number:,.2f}"
        if style == "fixed":
            return f"{number:.2f}"
        if style == "percent":
            return f"{number * 100:.2f}%"
        if style == "int":
            return str(int(round(number)))
        if style in {"sci", "scientific"}:
            return f"{number:.2e}"
        return text

    def _scroll_into_view(self) -> None:
        height, width = self.stdscr.getmaxyx()
        top_grid_row = 2
        visible_rows = max(1, height - top_grid_row - 2)
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
        for row, col, background in self.sheet.iter_backgrounds():
            if delta > 0 and col >= index:
                col += delta
            elif delta < 0:
                if index <= col < index - delta:
                    continue
                if col >= index - delta:
                    col += delta
            new_sheet.set_background(row, col, background)
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
        )
        for new_row, old_row in enumerate(order):
            for col in range(self.sheet.cols):
                raw = self.sheet.get_raw(old_row, col)
                if raw:
                    new_sheet.set_raw(new_row, col, raw)
                style = self.sheet.get_format(old_row, col)
                if style:
                    new_sheet.set_format(new_row, col, style)
                background = self.sheet.get_background(old_row, col)
                if background:
                    new_sheet.set_background(new_row, col, background)
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
                background = self.sheet.get_background(row, old_col)
                if background:
                    new_sheet.set_background(row, new_col, background)
                if self.sheet.is_protected(row, old_col):
                    new_sheet.protect(row, new_col)
        self.sheet = new_sheet
        self.evaluator = Evaluator(self.sheet)
def _run(stdscr, path: Path | None) -> int:
    app = SpreadsheetApp(stdscr, path=path)
    if path and path.exists():
        app.sheet = load_sheet(path)
        app.evaluator = Evaluator(app.sheet)
        app.dirty = False
        app.message = f"Loaded {path}"
    return app.run()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="A modular curses spreadsheet.")
    parser.add_argument("path", nargs="?", help="Optional .tss or .csv file to open")
    args = parser.parse_args(argv)
    path = Path(args.path).expanduser() if args.path else None
    return curses.wrapper(lambda stdscr: _run(stdscr, path))


if __name__ == "__main__":
    raise SystemExit(main())
