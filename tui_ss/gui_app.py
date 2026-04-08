#!/usr/bin/env python3
"""PySide GUI front end for tui-ss."""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import QPoint, Qt, QSignalBlocker
from PySide6.QtGui import QColor, QFont, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from .app import (
    ACTIVE_CELL_COLORS,
    APP_NAME,
    BACKGROUND_COLORS,
    CURRENCY_SYMBOLS,
    DATE_FORMATS,
    DEFAULT_SETTINGS_PATH,
    FORMULA_COLOR_OPTIONS,
    FORMULA_COLOR_SETTING_OPTIONS,
    FORMAT_STYLES,
    JUSTIFY_OPTIONS,
    PROTECTED_COLOR_OPTIONS,
    THEMES,
    build_stamp,
)
from .commands import (
    ALIASES,
    ADVANCED_COMMAND_MENU_OPTIONS,
    COMMAND_DESCRIPTIONS,
    COMMAND_MENU_OPTIONS,
    LANGUAGE_CODES,
    LANGUAGE_OPTIONS,
    get_command_help_lines,
    get_formula_help_lines,
    get_key_help_lines,
    parse_command,
    tr,
)
from .formulas import (
    Evaluator,
    FormulaError,
    format_date_text,
    is_formula_text,
    normalize_date_text,
    parse_date_text,
    unescape_literal_text,
)
from .model import Cell, Spreadsheet, TEXT_STYLE_NAMES, column_label, parse_cell_reference
from .storage import load_app_settings, load_sheet, save_app_settings, save_pdf_text, save_sheet

DEFAULT_PATH = Path.home() / "scripts" / "tui-ss" / "sheets" / "autosave.tss"

GUI_BG = "#161126"
GUI_PANEL = "#201733"
GUI_PANEL_2 = "#261b3c"
GUI_ACCENT = "#3ba5ff"
GUI_TEXT = "#d7e9ff"
GUI_MUTED = "#8aa4c9"
GUI_GRID = "#3a2957"
GUI_GREEN = "#3cff71"
GUI_RED = "#ff5d5d"
GUI_BLACK = "#08060f"

COLOR_HEX = {
    "black": "#111111",
    "white": "#f4f4f7",
    "yellow": "#f6ff4d",
    "pink": "#ff6fcf",
    "palepink": "#f1d1df",
    "orange": "#ff7a00",
    "lightblue": "#8fd3ff",
    "cornflower": "#6a8cff",
    "lightgrey": "#dddddd",
    "blue": "#3ba5ff",
    "cyan": "#44f0ff",
    "green": "#3cff71",
    "magenta": "#d858ff",
    "red": "#ff5d5d",
    "purple": "#6600ff",
}

ALIGN_MAP = {"left": Qt.AlignLeft | Qt.AlignVCenter, "center": Qt.AlignHCenter | Qt.AlignVCenter, "right": Qt.AlignRight | Qt.AlignVCenter}
CELL_REF_RE = re.compile(r"^[A-Za-z]+\d+$")


def color_hex(name: str, fallback: str = GUI_TEXT) -> str:
    return COLOR_HEX.get(name.lower(), fallback)


def ensure_sheet_defaults(sheet: Spreadsheet, defaults: dict[str, str]) -> Spreadsheet:
    if defaults.get("theme_name"):
        sheet.theme_name = defaults.get("theme_name", sheet.theme_name)
    raw_date = defaults.get("date_format", sheet.date_format)
    sheet.date_format = raw_date if raw_date.startswith("date:") else f"date:{raw_date}"
    if defaults.get("active_cell_color"):
        sheet.active_cell_color = defaults["active_cell_color"]
    formula_color = defaults.get("formula_coloration", "on")
    if formula_color == "off":
        sheet.formula_coloration = False
    elif formula_color:
        sheet.formula_coloration = True
        sheet.formula_foreground_color = formula_color
    if defaults.get("language"):
        sheet.language = defaults["language"]
    if defaults.get("protected_foreground_color"):
        sheet.protected_foreground_color = defaults["protected_foreground_color"]
    if defaults.get("protected_background_color"):
        sheet.protected_background_color = defaults["protected_background_color"]
    return sheet


def sheet_to_text_lines(sheet: Spreadsheet, evaluator: Evaluator, display_fn) -> list[str]:
    rows: list[int] = []
    cols: list[int] = []
    for row, col, _raw in sheet.iter_cells():
        rows.append(row)
        cols.append(col)
    if not rows or not cols:
        return [""]
    max_row = max(rows)
    max_col = max(cols)
    rendered: list[list[str]] = []
    widths = [0] * (max_col + 1)
    for row in range(max_row + 1):
        row_values = []
        for col in range(max_col + 1):
            value = display_fn(row, col)
            row_values.append(value)
            widths[col] = max(widths[col], len(value))
        rendered.append(row_values)
    return [" | ".join(value.ljust(widths[index]) for index, value in enumerate(row_values)).rstrip() for row_values in rendered]


@dataclass
class GuiTabState:
    widget: "SheetView"
    path: Path | None


class CommandPaletteDialog(QDialog):
    def __init__(self, parent: QWidget, title: str, options: list[str], descriptions: dict[str, str], prefix: str = "/") -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.resize(760, 420)
        self.options = options
        self.descriptions = descriptions
        self.prefix = prefix
        self.result_text: str | None = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)
        title_label = QLabel(f"{prefix} commands", self)
        title_label.setStyleSheet(f"color:{GUI_ACCENT}; font-size:18px; font-weight:700;")
        self.prompt = QLineEdit(self)
        self.prompt.setText(prefix)
        self.listing = QListWidget(self)
        self.description = QLabel(self)
        self.description.setWordWrap(True)
        self.description.setStyleSheet(f"color:{GUI_MUTED};")
        hint = QLabel("type to filter, Enter to run, Esc to cancel", self)
        hint.setStyleSheet(f"color:{GUI_MUTED};")
        layout.addWidget(title_label)
        layout.addWidget(self.prompt)
        layout.addWidget(self.listing, 1)
        layout.addWidget(self.description)
        layout.addWidget(hint)
        self.prompt.textChanged.connect(self._refresh)
        self.prompt.returnPressed.connect(self._accept_current)
        self.listing.itemDoubleClicked.connect(lambda _item: self._accept_current())
        self.listing.currentItemChanged.connect(lambda current, _previous: self._update_description(current.text() if current else ""))
        self._refresh()
        self.prompt.setFocus()
        self.setStyleSheet(
            f"""
            QDialog {{ background:{GUI_BLACK}; color:{GUI_TEXT}; border:1px solid {GUI_GRID}; }}
            QLineEdit {{ background:{GUI_PANEL}; color:{GUI_TEXT}; border:1px solid {GUI_GRID}; padding:10px 12px; font-size:18px; font-family:monospace; }}
            QListWidget {{ background:{GUI_PANEL}; color:{GUI_TEXT}; border:1px solid {GUI_GRID}; font-family:monospace; font-size:16px; outline:none; }}
            QListWidget::item {{ padding:8px 10px; }}
            QListWidget::item:selected {{ background:{color_hex('orange')}; color:#111111; }}
            """
        )

    def _refresh(self) -> None:
        typed = self.prompt.text().strip()
        query = typed[len(self.prefix):].strip().lower() if typed.startswith(self.prefix) else typed.lower()
        self.listing.clear()
        filtered = [option for option in self.options if query in option.lower()]
        for option in filtered or self.options:
            self.listing.addItem(option)
        if self.listing.count():
            self.listing.setCurrentRow(0)
            self._update_description(self.listing.currentItem().text())

    def _update_description(self, option: str) -> None:
        self.description.setText(self.descriptions.get(option, ""))

    def _accept_current(self) -> None:
        typed = self.prompt.text().strip()
        if " " in typed[len(self.prefix):].strip():
            self.result_text = typed
        elif typed.strip() == self.prefix and self.listing.currentItem():
            self.result_text = f"{self.prefix}{self.listing.currentItem().text()}"
        elif typed.startswith(self.prefix) and typed[len(self.prefix):].strip():
            current = typed[len(self.prefix):].strip()
            if current in self.options:
                self.result_text = f"{self.prefix}{current}"
            elif self.listing.currentItem():
                self.result_text = f"{self.prefix}{self.listing.currentItem().text()}"
            else:
                self.result_text = typed
        else:
            self.result_text = typed
        self.accept()


class CommandLineEdit(QLineEdit):
    def __init__(self, owner: "GuiSpreadsheetWindow", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.owner = owner

    def keyPressEvent(self, event) -> None:  # type: ignore[override]
        key = event.key()
        if key == Qt.Key_Up:
            self.owner._move_command_selection(-1)
            event.accept()
            return
        if key == Qt.Key_Down:
            self.owner._move_command_selection(1)
            event.accept()
            return
        if key == Qt.Key_PageUp:
            self.owner._move_command_selection(-5)
            event.accept()
            return
        if key == Qt.Key_PageDown:
            self.owner._move_command_selection(5)
            event.accept()
            return
        if key == Qt.Key_Escape:
            self.owner.close_command_palette()
            event.accept()
            return
        if key in {Qt.Key_Return, Qt.Key_Enter}:
            self.owner.accept_command_palette()
            event.accept()
            return
        super().keyPressEvent(event)


class HelpDialog(QDialog):
    def __init__(self, parent: QWidget, title: str, lines: list[str]) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(920, 640)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)
        title_label = QLabel(title, self)
        title_label.setStyleSheet(f"background:{GUI_BLACK}; color:{GUI_ACCENT}; padding:8px 10px; font-family:monospace; font-size:18px; font-weight:700;")
        text = QTextEdit(self)
        text.setReadOnly(True)
        text.setPlainText("\n".join(lines))
        hint = QLabel("Esc closes help", self)
        hint.setStyleSheet(f"color:{GUI_MUTED}; font-family:monospace;")
        layout.addWidget(title_label)
        layout.addWidget(text)
        layout.addWidget(hint)
        self.setStyleSheet(
            f"""
            QDialog {{ background:{GUI_BLACK}; color:{GUI_TEXT}; border:1px solid {GUI_GRID}; }}
            QTextEdit {{ background:{GUI_PANEL}; color:{GUI_TEXT}; border:1px solid {GUI_GRID}; font-family:monospace; font-size:15px; padding:8px; }}
            """
        )


class SettingsDialog(QDialog):
    def __init__(self, parent: QWidget, sheet: Spreadsheet, settings_path: Path, recent_files: list[str]) -> None:
        super().__init__(parent)
        self.setWindowTitle(tr(sheet.language, "settings"))
        self.sheet = sheet
        self.settings_path = settings_path
        self.recent_files = recent_files
        self.resize(760, 520)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(14)
        title_label = QLabel(tr(sheet.language, "settings"), self)
        title_label.setStyleSheet(f"background:{GUI_BLACK}; color:{GUI_ACCENT}; padding:8px 10px; font-family:monospace; font-size:18px; font-weight:700;")
        form = QFormLayout()
        form.setHorizontalSpacing(24)
        form.setVerticalSpacing(16)
        form.setLabelAlignment(Qt.AlignLeft)
        self.theme = QComboBox(self)
        self.theme.addItems(THEMES)
        self.theme.setCurrentText(sheet.theme_name)
        self.date_format = QComboBox(self)
        self.date_format.addItems(DATE_FORMATS)
        self.date_format.setCurrentText(sheet.date_format.replace("date:", ""))
        self.active_cell = QComboBox(self)
        self.active_cell.addItems(ACTIVE_CELL_COLORS)
        self.active_cell.setCurrentText(sheet.active_cell_color)
        self.formula_color = QComboBox(self)
        self.formula_color.addItems(FORMULA_COLOR_SETTING_OPTIONS)
        self.formula_color.setCurrentText(sheet.formula_foreground_color if sheet.formula_coloration else "off")
        self.language = QComboBox(self)
        self.language.addItems(LANGUAGE_OPTIONS)
        current_language = next((label for label, code in LANGUAGE_CODES.items() if code == sheet.language), "english")
        self.language.setCurrentText(current_language)
        self.protected_fg = QComboBox(self)
        self.protected_fg.addItems(PROTECTED_COLOR_OPTIONS)
        self.protected_fg.setCurrentText(sheet.protected_foreground_color)
        self.protected_bg = QComboBox(self)
        self.protected_bg.addItems(PROTECTED_COLOR_OPTIONS)
        self.protected_bg.setCurrentText(sheet.protected_background_color)
        form.addRow(tr(sheet.language, "theme"), self.theme)
        form.addRow(tr(sheet.language, "date_format"), self.date_format)
        form.addRow(tr(sheet.language, "active_cell"), self.active_cell)
        form.addRow("Formula Colour", self.formula_color)
        form.addRow(tr(sheet.language, "language"), self.language)
        form.addRow("Protected FG", self.protected_fg)
        form.addRow("Protected BG", self.protected_bg)
        layout.addWidget(title_label)
        layout.addLayout(form)
        buttons = QHBoxLayout()
        buttons.addStretch(1)
        save = QPushButton("Save", self)
        save.setStyleSheet("background:#1e8f39;color:white;padding:8px 16px;")
        close = QPushButton("Quit", self)
        close.setStyleSheet("background:#c62828;color:white;padding:8px 16px;")
        save.clicked.connect(self.accept)
        close.clicked.connect(self.reject)
        buttons.addWidget(save)
        buttons.addWidget(close)
        layout.addStretch(1)
        layout.addLayout(buttons)
        self.setStyleSheet(
            f"""
            QDialog {{ background:{GUI_BLACK}; color:{GUI_TEXT}; border:1px solid {GUI_GRID}; }}
            QLabel {{ color:{GUI_TEXT}; font-family:monospace; font-size:16px; }}
            QComboBox {{
                background:{GUI_PANEL};
                color:{GUI_TEXT};
                border:1px solid {GUI_GRID};
                padding:8px 10px;
                min-width:220px;
                font-family:monospace;
                font-size:15px;
            }}
            QComboBox QAbstractItemView {{
                background:{GUI_PANEL};
                color:{GUI_TEXT};
                selection-background-color:{color_hex(sheet.active_cell_color, color_hex('orange'))};
                selection-color:#111111;
                font-family:monospace;
            }}
            """
        )

    def apply(self) -> None:
        self.sheet.theme_name = self.theme.currentText()
        self.sheet.date_format = f"date:{self.date_format.currentText()}"
        self.sheet.active_cell_color = self.active_cell.currentText()
        formula_value = self.formula_color.currentText()
        self.sheet.formula_coloration = formula_value != "off"
        if formula_value != "off":
            self.sheet.formula_foreground_color = formula_value
        self.sheet.language = LANGUAGE_CODES[self.language.currentText()]
        self.sheet.protected_foreground_color = self.protected_fg.currentText()
        self.sheet.protected_background_color = self.protected_bg.currentText()
        save_app_settings(
            self.settings_path,
            {
                "theme_name": self.sheet.theme_name,
                "date_format": self.sheet.date_format,
                "active_cell_color": self.sheet.active_cell_color,
                "formula_coloration": formula_value,
                "formula_foreground_color": self.sheet.formula_foreground_color,
                "language": self.sheet.language,
                "protected_foreground_color": self.sheet.protected_foreground_color,
                "protected_background_color": self.sheet.protected_background_color,
                "recent_files_json": __import__("json").dumps(self.recent_files),
            },
        )


class SheetView(QWidget):
    def __init__(self, sheet: Spreadsheet, path: Path | None) -> None:
        super().__init__()
        self.sheet = sheet
        self.path = path
        self.dirty = False
        self.evaluator = Evaluator(self.sheet)
        self.updating = False

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)
        self.title_line = QLabel(self)
        self.formula_line = QLineEdit(self)
        self.formula_line.returnPressed.connect(self._commit_formula_line)
        self.info_line = QLabel(self)
        self.info_line.setStyleSheet(f"color:{GUI_MUTED};")
        layout.addWidget(self.title_line)
        formula_box = QHBoxLayout()
        formula_box.setContentsMargins(0, 0, 0, 0)
        formula_box.setSpacing(8)
        self.ref_label = QLabel("A1", self)
        self.ref_label.setMinimumWidth(60)
        formula_box.addWidget(self.ref_label)
        formula_box.addWidget(self.formula_line, 1)
        formula_box.addWidget(self.info_line, 1)
        layout.addLayout(formula_box)

        self.table = QTableWidget(self.sheet.rows, self.sheet.cols, self)
        self.table.itemChanged.connect(self._item_changed)
        self.table.currentCellChanged.connect(self._current_changed)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.setSelectionBehavior(QTableWidget.SelectItems)
        self.table.horizontalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.verticalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.horizontalHeader().customContextMenuRequested.connect(self._show_column_menu)
        self.table.verticalHeader().customContextMenuRequested.connect(self._show_row_menu)
        self.table.horizontalHeader().sectionClicked.connect(self._select_column_header)
        self.table.verticalHeader().sectionClicked.connect(self._select_row_header)
        self.table.horizontalHeader().sectionDoubleClicked.connect(self._set_column_width)
        self.table.verticalHeader().sectionDoubleClicked.connect(self._set_row_height)
        layout.addWidget(self.table, 1)
        self._populate()

    def _populate(self) -> None:
        self.updating = True
        self.table.setRowCount(self.sheet.rows)
        self.table.setColumnCount(self.sheet.cols)
        self.table.setHorizontalHeaderLabels([column_label(index) for index in range(self.sheet.cols)])
        for row in range(self.sheet.rows):
            self.table.setVerticalHeaderItem(row, QTableWidgetItem(str(row + 1)))
            self.table.setRowHeight(row, 30)
            for col in range(self.sheet.cols):
                self._apply_cell(row, col)
        for col in range(self.sheet.cols):
            self.table.setColumnWidth(col, self.sheet.get_column_width(col) * 11)
        self._apply_theme()
        self.updating = False
        self._current_changed(self.table.currentRow(), self.table.currentColumn(), -1, -1)

    def _apply_theme(self) -> None:
        active_bg = color_hex(self.sheet.active_cell_color, color_hex("orange"))
        active_fg = "#111111" if self.sheet.active_cell_color not in {"blue", "cornflower"} else "#ffffff"
        self.table.setStyleSheet(
            f"""
            QTableWidget {{
                background:{GUI_BG};
                color:{GUI_TEXT};
                gridline-color:{GUI_GRID};
                selection-background-color:{active_bg};
                selection-color:{active_fg};
                font-size:16px;
                border:none;
                alternate-background-color:{GUI_PANEL};
            }}
            QTableWidget::item:selected {{
                background:{active_bg};
                color:{active_fg};
            }}
            QTableWidget::item:selected:active {{
                background:{active_bg};
                color:{active_fg};
            }}
            QTableWidget::item:selected:!active {{
                background:{active_bg};
                color:{active_fg};
            }}
            QHeaderView::section {{
                background:#000000;
                color:{color_hex(self.sheet.theme_name, GUI_ACCENT)};
                border:none;
                border-right:1px solid {GUI_GRID};
                border-bottom:1px solid {GUI_GRID};
                padding:4px 6px;
                font-family:monospace;
                font-size:15px;
            }}
            QScrollBar:vertical, QScrollBar:horizontal {{
                background:{GUI_BLACK};
                border:none;
                margin:0;
            }}
            QScrollBar::handle:vertical, QScrollBar::handle:horizontal {{
                background:{GUI_GRID};
                min-height:22px;
                min-width:22px;
                border-radius:4px;
            }}
            QScrollBar::add-line, QScrollBar::sub-line,
            QScrollBar::add-page, QScrollBar::sub-page {{
                background:none;
                border:none;
            }}
            """
        )
        self.title_line.setStyleSheet(f"background:{GUI_BLACK}; color:{GUI_ACCENT}; padding:6px 8px; font-family:monospace; font-size:16px;")
        self.formula_line.setStyleSheet(f"background:{GUI_BLACK}; color:{GUI_TEXT}; border:1px solid {GUI_GRID}; padding:6px 8px; font-family:monospace; font-size:16px;")
        self.ref_label.setStyleSheet(f"background:{GUI_BLACK}; color:{GUI_ACCENT}; font-weight:700; padding:6px 8px; font-family:monospace;")
        self.info_line.setStyleSheet(f"background:{GUI_BLACK}; color:{GUI_MUTED}; padding:6px 8px; font-family:monospace;")

    def _display_value(self, row: int, col: int) -> str:
        raw = self.sheet.get_raw(row, col)
        if not raw:
            return ""
        try:
            text = self.evaluator.display_value(row, col)
        except FormulaError as exc:
            return f"#ERR {exc}"
        style = self.sheet.get_format(row, col)
        if style:
            return self._apply_format(text, style)
        if self.sheet.date_format.startswith("date") and parse_date_text(text, self.sheet.date_format) is not None:
            try:
                return format_date_text(text, self.sheet.date_format)
            except ValueError:
                return text
        return text

    def _apply_format(self, text: str, style: str) -> str:
        if not text:
            return text
        normalized_style = style.lower()
        if normalized_style.startswith("date"):
            return format_date_text(text, normalized_style)
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
            rendered = f"{abs(number):,.2f}"
            return f"({rendered})" if number < 0 else rendered
        if normalized_style in {"sci", "scientific"}:
            return f"{number:.2e}"
        return text

    def _apply_cell(self, row: int, col: int) -> None:
        raw = self.sheet.get_raw(row, col)
        item = self.table.item(row, col)
        if item is None:
            item = QTableWidgetItem()
            self.table.setItem(row, col, item)
        item.setText(self._display_value(row, col))
        item.setData(Qt.UserRole, raw)
        item.setFlags(item.flags() | Qt.ItemIsEditable | Qt.ItemIsSelectable | Qt.ItemIsEnabled)
        if self.sheet.is_protected(row, col):
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            item.setBackground(QColor(color_hex(self.sheet.protected_background_color, "#dddddd")))
            item.setForeground(QColor(color_hex(self.sheet.protected_foreground_color, "#111111")))
        else:
            background = self.sheet.get_background(row, col)
            item.setBackground(QColor(color_hex(background, GUI_BG)) if background else QColor(GUI_BG))
            if self.sheet.get_format(row, col) == "negative":
                try:
                    number = float(item.text().replace("%", "").replace(",", ""))
                    item.setForeground(QColor(GUI_RED if number < 0 else GUI_TEXT))
                except ValueError:
                    item.setForeground(QColor(GUI_TEXT))
            elif is_formula_text(raw) and self.sheet.formula_coloration:
                item.setForeground(QColor(color_hex(self.sheet.formula_foreground_color, GUI_GREEN)))
            else:
                item.setForeground(QColor(GUI_TEXT))
        styles = self.sheet.get_text_styles(row, col)
        font = QFont("Monospace")
        font.setStyleHint(QFont.TypeWriter)
        font.setBold("bold" in styles)
        font.setUnderline("underline" in styles)
        font.setItalic("italic" in styles)
        font_size = self.sheet.get_font_size(row, col)
        if font_size:
            font.setPointSize(font_size)
            self.table.setRowHeight(row, max(self.table.rowHeight(row), int(font_size * 1.9)))
        else:
            font.setPointSize(14)
        item.setFont(font)
        align = self.sheet.get_alignment(row, col)
        item.setTextAlignment(ALIGN_MAP.get(align, Qt.AlignLeft | Qt.AlignVCenter))

    def _item_changed(self, item: QTableWidgetItem) -> None:
        if self.updating:
            return
        row = item.row()
        col = item.column()
        if self.sheet.is_protected(row, col):
            self.updating = True
            with QSignalBlocker(self.table):
                item.setText(self._display_value(row, col))
            self.updating = False
            return
        text = item.text()
        if self.sheet.get_format(row, col).startswith("date"):
            try:
                text = normalize_date_text(text, self.sheet.get_format(row, col))
            except ValueError:
                pass
        elif self.sheet.date_format.startswith("date") and parse_date_text(text, self.sheet.date_format) is not None:
            text = normalize_date_text(text, self.sheet.date_format)
        self.sheet.set_raw(row, col, text)
        self.updating = True
        with QSignalBlocker(self.table):
            self._apply_cell(row, col)
        self.updating = False
        self.dirty = True
        self._current_changed(row, col, row, col)

    def _current_changed(self, current_row: int, current_col: int, _previous_row: int, _previous_col: int) -> None:
        if current_row < 0 or current_col < 0:
            return
        ref = f"{column_label(current_col)}{current_row + 1}"
        raw = self.sheet.get_raw(current_row, current_col)
        path_text = str(self.path) if self.path else "untitled"
        self.title_line.setText(f"{APP_NAME}  {path_text}  defw={self.sheet.column_width}  {column_label(current_col)}w={self.sheet.get_column_width(current_col)}  build={build_stamp()}")
        self.ref_label.setText(ref)
        with QSignalBlocker(self.formula_line):
            self.formula_line.setText(raw)
        try:
            value = self.evaluator.display_value(current_row, current_col)
            self.info_line.setText(f"raw={raw or ''}   value={value or ''}")
        except FormulaError as exc:
            self.info_line.setText(f"raw={raw or ''}   #ERR {exc}")

    def _commit_formula_line(self) -> None:
        row = self.table.currentRow()
        col = self.table.currentColumn()
        if row < 0 or col < 0:
            return
        self.sheet.set_raw(row, col, self.formula_line.text())
        self.updating = True
        with QSignalBlocker(self.table):
            self._apply_cell(row, col)
        self.updating = False
        self.dirty = True
        self._current_changed(row, col, row, col)

    def selected_range(self) -> tuple[int, int, int, int]:
        ranges = self.table.selectedRanges()
        if ranges:
            area = ranges[0]
            return area.topRow(), area.leftColumn(), area.bottomRow(), area.rightColumn()
        row = max(0, self.table.currentRow())
        col = max(0, self.table.currentColumn())
        return row, col, row, col

    def selected_rows(self) -> list[int]:
        model = self.table.selectionModel()
        return sorted({index.row() for index in model.selectedRows()})

    def selected_columns(self) -> list[int]:
        model = self.table.selectionModel()
        return sorted({index.column() for index in model.selectedColumns()})

    def copy_selection_to_clipboard(self) -> None:
        row_lo, col_lo, row_hi, col_hi = self.selected_range()
        lines = []
        for row in range(row_lo, row_hi + 1):
            lines.append("\t".join(self.sheet.get_raw(row, col) for col in range(col_lo, col_hi + 1)))
        QApplication.clipboard().setText("\n".join(lines))

    def paste_from_clipboard(self) -> None:
        start_row = max(0, self.table.currentRow())
        start_col = max(0, self.table.currentColumn())
        text = QApplication.clipboard().text()
        if not text:
            return
        rows = [line.split("\t") for line in text.splitlines()]
        self.updating = True
        with QSignalBlocker(self.table):
            for row_offset, values in enumerate(rows):
                for col_offset, raw in enumerate(values):
                    row = start_row + row_offset
                    col = start_col + col_offset
                    self.sheet.ensure_size(row, col)
                    self.sheet.set_raw(row, col, raw)
                    self._apply_cell(row, col)
        self.updating = False
        self.dirty = True

    def clear_selection(self) -> None:
        row_lo, col_lo, row_hi, col_hi = self.selected_range()
        self.updating = True
        with QSignalBlocker(self.table):
            for row in range(row_lo, row_hi + 1):
                for col in range(col_lo, col_hi + 1):
                    self.sheet.set_raw(row, col, "")
                    self.sheet.set_format(row, col, "")
                    self.sheet.clear_text_styles(row, col)
                    self.sheet.set_background(row, col, "")
                    self.sheet.set_font_size(row, col, 0)
                    self._apply_cell(row, col)
        self.updating = False
        self.dirty = True

    def set_protection(self, enabled: bool) -> None:
        row_lo, col_lo, row_hi, col_hi = self.selected_range()
        self.updating = True
        with QSignalBlocker(self.table):
            for row in range(row_lo, row_hi + 1):
                for col in range(col_lo, col_hi + 1):
                    if enabled:
                        self.sheet.protect(row, col)
                    else:
                        self.sheet.unprotect(row, col)
                    self._apply_cell(row, col)
        self.updating = False
        self.dirty = True

    def insert_rows(self, index: int, count: int = 1) -> None:
        for _ in range(count):
            self.table.insertRow(index)
            self.sheet.rows += 1
            new_cells: dict[str, object] = {}
            for row, col, raw in list(self.sheet.iter_cells()):
                target_row = row + 1 if row >= index else row
                new_cells[f"{target_row}:{col}"] = type(next(iter(self.sheet.cells.values()), None))(raw=raw) if False else None
            self._shift_row_metadata(index, 1)
        self._rebuild_from_sheet()

    def delete_rows(self, index: int, count: int = 1) -> None:
        if self.sheet.rows <= 1:
            return
        self._shift_row_metadata(index, -count)
        self.sheet.rows = max(1, self.sheet.rows - count)
        self._rebuild_from_sheet()

    def insert_columns(self, index: int, count: int = 1) -> None:
        self._shift_col_metadata(index, count)
        self.sheet.cols += count
        self._rebuild_from_sheet()

    def delete_columns(self, index: int, count: int = 1) -> None:
        if self.sheet.cols <= 1:
            return
        self._shift_col_metadata(index, -count)
        self.sheet.cols = max(1, self.sheet.cols - count)
        self._rebuild_from_sheet()

    def _shift_row_metadata(self, index: int, delta: int) -> None:
        self._shift_dimension(index, delta, is_row=True)

    def _shift_col_metadata(self, index: int, delta: int) -> None:
        self._shift_dimension(index, delta, is_row=False)

    def _shift_dimension(self, index: int, delta: int, is_row: bool) -> None:
        def shift_key_map(source: dict[str, str]) -> dict[str, str]:
            result: dict[str, str] = {}
            for key, value in source.items():
                row_text, col_text = key.split(":", 1)
                row = int(row_text)
                col = int(col_text)
                pivot = row if is_row else col
                if delta < 0 and index <= pivot < index + abs(delta):
                    continue
                if pivot >= index:
                    if is_row:
                        row = max(index, row + delta)
                    else:
                        col = max(index, col + delta)
                result[f"{row}:{col}"] = value
            return result

        def shift_key_set(source: set[str]) -> set[str]:
            result: set[str] = set()
            for key in source:
                row_text, col_text = key.split(":", 1)
                row = int(row_text)
                col = int(col_text)
                pivot = row if is_row else col
                if delta < 0 and index <= pivot < index + abs(delta):
                    continue
                if pivot >= index:
                    if is_row:
                        row = max(index, row + delta)
                    else:
                        col = max(index, col + delta)
                result.add(f"{row}:{col}")
            return result

        rebuilt_cells: dict[str, Cell] = {}
        for row, col, raw in list(self.sheet.iter_cells()):
            pivot = row if is_row else col
            if delta < 0 and index <= pivot < index + abs(delta):
                continue
            if pivot >= index:
                if is_row:
                    row = max(index, row + delta)
                else:
                    col = max(index, col + delta)
            rebuilt_cells[f"{row}:{col}"] = Cell(raw=raw)
        self.sheet.cells = rebuilt_cells
        self.sheet.formats = shift_key_map(self.sheet.formats)
        self.sheet.text_styles = shift_key_map(self.sheet.text_styles)
        self.sheet.backgrounds = shift_key_map(self.sheet.backgrounds)
        self.sheet.font_sizes = {
            key: int(value)
            for key, value in shift_key_map({k: str(v) for k, v in self.sheet.font_sizes.items()}).items()
        }
        self.sheet.alignments = shift_key_map(self.sheet.alignments)
        self.sheet.manual_alignments = shift_key_set(self.sheet.manual_alignments)
        self.sheet.protected = shift_key_set(self.sheet.protected)

    def _rebuild_from_sheet(self) -> None:
        self.evaluator = Evaluator(self.sheet)
        self._populate()
        self.dirty = True

    def _select_row_header(self, row: int) -> None:
        self.table.selectRow(row)

    def _select_column_header(self, col: int) -> None:
        self.table.selectColumn(col)

    def _show_row_menu(self, point: QPoint) -> None:
        row = self.table.verticalHeader().logicalIndexAt(point)
        if row < 0:
            return
        menu = QMenu(self)
        freeze = menu.addAction("Freeze through row")
        insert_above = menu.addAction("Insert above")
        delete = menu.addAction("Delete row")
        chosen = menu.exec(self.table.verticalHeader().mapToGlobal(point))
        if chosen == freeze:
            self.sheet.title_rows = row + 1
            self.dirty = True
        elif chosen == insert_above:
            self.insert_rows(row, 1)
        elif chosen == delete:
            self.delete_rows(row, 1)

    def _show_column_menu(self, point: QPoint) -> None:
        col = self.table.horizontalHeader().logicalIndexAt(point)
        if col < 0:
            return
        menu = QMenu(self)
        freeze = menu.addAction("Freeze through column")
        insert_before = menu.addAction("Insert before")
        delete = menu.addAction("Delete column")
        width = menu.addAction("Width")
        chosen = menu.exec(self.table.horizontalHeader().mapToGlobal(point))
        if chosen == freeze:
            self.sheet.title_cols = col + 1
            self.dirty = True
        elif chosen == insert_before:
            self.insert_columns(col, 1)
        elif chosen == delete:
            self.delete_columns(col, 1)
        elif chosen == width:
            self._set_column_width(col)

    def apply_format(self, style: str, extra: str | None = None) -> None:
        row_lo, col_lo, row_hi, col_hi = self.selected_range()
        for row in range(row_lo, row_hi + 1):
            for col in range(col_lo, col_hi + 1):
                if style in TEXT_STYLE_NAMES:
                    enabled = not self.sheet.has_text_style(row, col, style)
                    self.sheet.set_text_style(row, col, style, enabled=enabled)
                elif style == "background":
                    self.sheet.set_background(row, col, "" if extra == "none" else (extra or "blue"))
                elif style == "font-size":
                    self.sheet.set_font_size(row, col, int(extra or 14))
                elif style == "clear-format":
                    self.sheet.set_format(row, col, "")
                    self.sheet.clear_text_styles(row, col)
                    self.sheet.set_background(row, col, "")
                    self.sheet.set_font_size(row, col, 0)
                else:
                    if style == "date":
                        self.sheet.date_format = f"date:{extra or 'european'}"
                    elif style == "currency":
                        self.sheet.set_format(row, col, f"currency:{extra or '£'}")
                    else:
                        self.sheet.set_format(row, col, style)
                self._apply_cell(row, col)
        self._apply_theme()
        self.dirty = True

    def set_alignment(self, align: str) -> None:
        row_lo, col_lo, row_hi, col_hi = self.selected_range()
        resolved = {"left": "left", "centre": "center", "center": "center", "right": "right"}[align]
        for row in range(row_lo, row_hi + 1):
            for col in range(col_lo, col_hi + 1):
                self.sheet.set_alignment(row, col, resolved)
                self._apply_cell(row, col)
        self.dirty = True

    def _set_column_width(self, col: int) -> None:
        width, ok = QInputDialog.getInt(self, "Column Width", f"Width for {column_label(col)}", self.sheet.get_column_width(col), 8, 80)
        if not ok:
            return
        self.sheet.set_column_width(col, width)
        self.table.setColumnWidth(col, width * 11)
        self.dirty = True

    def _set_row_height(self, row: int) -> None:
        height, ok = QInputDialog.getInt(self, "Row Height", f"Height for row {row + 1}", self.table.rowHeight(row), 20, 200)
        if not ok:
            return
        self.table.setRowHeight(row, height)
        self.dirty = True


class GuiSpreadsheetWindow(QMainWindow):
    def __init__(self, paths: list[Path] | None = None, settings_path: Path | None = None) -> None:
        super().__init__()
        self.settings_path = settings_path or DEFAULT_SETTINGS_PATH
        self.defaults = load_app_settings(self.settings_path)
        self.recent_files = self._parse_recent_files(self.defaults.get("recent_files_json", "[]"))
        self.setWindowTitle("gui-ss")
        self.resize(1480, 960)
        self.tabs = QTabWidget(self)
        self.tabs.setDocumentMode(True)
        self.tabs.setMovable(True)
        self.tabs.setTabsClosable(True)
        self.tabs.tabCloseRequested.connect(self.close_tab)
        self.tabs.currentChanged.connect(lambda _index: self._sync_title())
        central = QWidget(self)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(4)
        layout.addWidget(self.tabs, 1)

        self.command_panel = QWidget(self)
        self.command_panel.hide()
        command_layout = QVBoxLayout(self.command_panel)
        command_layout.setContentsMargins(0, 0, 0, 0)
        command_layout.setSpacing(4)
        self.command_line = CommandLineEdit(self, self.command_panel)
        self.command_line.textChanged.connect(self._refresh_command_palette)
        self.command_choices = QListWidget(self.command_panel)
        self.command_choices.setMaximumHeight(156)
        self.command_choices.itemDoubleClicked.connect(lambda _item: self.accept_command_palette())
        self.command_choices.currentItemChanged.connect(lambda current, _prev: self._update_command_description(current.text() if current else ""))
        self.command_description = QLabel(self.command_panel)
        self.command_description.setWordWrap(True)
        self.command_description.setStyleSheet(f"color:{GUI_MUTED}; font-family:monospace; padding:4px 8px;")
        command_layout.addWidget(self.command_line)
        command_layout.addWidget(self.command_choices)
        command_layout.addWidget(self.command_description)
        layout.addWidget(self.command_panel)

        footer = QHBoxLayout()
        footer.setContentsMargins(0, 0, 0, 0)
        self.hint = QLabel('Press "/" to start', self)
        self.hint.setAlignment(Qt.AlignCenter)
        self.hint.setStyleSheet(f"color:{GUI_MUTED}; font-size:15px; font-family:monospace;")
        footer.addStretch(1)
        footer.addWidget(self.hint, 1)
        footer.addStretch(1)
        self.quit_button = QPushButton("Quit", self)
        self.quit_button.setFixedWidth(96)
        self.quit_button.setStyleSheet("background:#c62828;color:white;padding:6px 12px;border:none;")
        self.quit_button.clicked.connect(self.quit_current)
        footer.addWidget(self.quit_button)
        layout.addLayout(footer)
        self.setCentralWidget(central)

        self.setStyleSheet(
            f"""
            QMainWindow {{ background:{GUI_PANEL_2}; color:{GUI_TEXT}; }}
            QTabWidget::pane {{ border:1px solid {GUI_GRID}; top:-1px; }}
            QTabBar::tab {{
                background:{GUI_BLACK};
                color:{GUI_MUTED};
                padding:6px 12px;
                margin-right:2px;
                font-family:monospace;
                border:1px solid {GUI_GRID};
                border-bottom:none;
            }}
            QTabBar::tab:selected {{
                color:{GUI_ACCENT};
                background:{GUI_PANEL};
            }}
            QMenu {{
                background:{GUI_BLACK};
                color:{GUI_TEXT};
                border:1px solid {GUI_GRID};
                font-family:monospace;
                padding:4px;
            }}
            QMenu::item {{
                padding:6px 24px 6px 10px;
            }}
            QMenu::item:selected {{
                background:{color_hex('orange')};
                color:#111111;
            }}
            QLineEdit#commandLine {{
                background:{GUI_BLACK};
                color:{GUI_TEXT};
                border:1px solid {GUI_GRID};
                padding:8px 10px;
                font-family:monospace;
                font-size:16px;
            }}
            QListWidget#commandChoices {{
                background:{GUI_BLACK};
                color:{GUI_TEXT};
                border:1px solid {GUI_GRID};
                font-family:monospace;
                font-size:15px;
                outline:none;
            }}
            QListWidget#commandChoices::item {{
                padding:6px 10px;
            }}
            QListWidget#commandChoices::item:selected {{
                background:{color_hex('orange')};
                color:#111111;
            }}
            QInputDialog, QMessageBox {{
                background:{GUI_BLACK};
                color:{GUI_TEXT};
            }}
            """
        )
        self.command_line.setObjectName("commandLine")
        self.command_choices.setObjectName("commandChoices")
        QShortcut(QKeySequence(Qt.Key_Slash), self, activated=self.open_command_palette)
        QShortcut(QKeySequence.Save, self, activated=self.save_current)
        QShortcut(QKeySequence.Quit, self, activated=self.quit_current)
        QShortcut(QKeySequence.Copy, self, activated=self.copy_current_selection)
        QShortcut(QKeySequence.Paste, self, activated=self.paste_into_current)
        QShortcut(QKeySequence("Ctrl+Y"), self, activated=self.paste_into_current)
        QShortcut(QKeySequence("Ctrl+B"), self, activated=lambda: self._toggle_style_shortcut("bold"))
        QShortcut(QKeySequence("Ctrl+U"), self, activated=lambda: self._toggle_style_shortcut("underline"))
        QShortcut(QKeySequence("Ctrl+I"), self, activated=lambda: self._toggle_style_shortcut("italic"))
        QShortcut(QKeySequence("Shift+Space"), self, activated=self.select_current_row)
        QShortcut(QKeySequence("Ctrl+Space"), self, activated=self.select_current_column)

        files = paths or []
        if files:
            for path in files:
                self.open_sheet(path)
        else:
            self.new_sheet()
        self._sync_title()

    def _command_prefix(self) -> str:
        text = self.command_line.text()
        return "//" if text.startswith("//") else "/"

    def _command_options(self) -> list[str]:
        return ADVANCED_COMMAND_MENU_OPTIONS if self._command_prefix() == "//" else COMMAND_MENU_OPTIONS

    def open_command_palette(self) -> None:
        self.hint.hide()
        if self.command_panel.isVisible():
            current = self.command_line.text().strip()
            if current == "/":
                self.command_line.setText("//")
            elif current == "//":
                self.command_line.setText("/")
            else:
                self.command_line.insert("/")
            self.command_line.setFocus()
            return
        self.command_panel.show()
        self.command_line.setText("/")
        self.command_line.setFocus()
        self._refresh_command_palette()

    def close_command_palette(self) -> None:
        self.command_panel.hide()
        self.command_line.clear()
        self.command_choices.clear()
        self.command_description.clear()
        current = self._current_view()
        if current:
            current.table.setFocus()

    def _refresh_command_palette(self) -> None:
        if not self.command_panel.isVisible():
            return
        typed = self.command_line.text().strip()
        prefix = "//" if typed.startswith("//") else "/"
        query = typed[len(prefix):].strip().lower() if typed.startswith(prefix) else typed.lower()
        options = self._command_options()
        if not query:
            ranked = list(options)
        else:
            alias_target = ALIASES.get(query)
            starts = [option for option in options if option.lower().startswith(query)]
            contains = [option for option in options if query in option.lower() and option not in starts]
            ranked = []
            if alias_target and alias_target in options:
                ranked.append(alias_target)
            for bucket in (starts, contains):
                for option in bucket:
                    if option not in ranked:
                        ranked.append(option)
        self.command_choices.clear()
        for option in ranked or options:
            self.command_choices.addItem(f"{prefix}{option}")
        if self.command_choices.count():
            self.command_choices.setCurrentRow(0)
            self._update_command_description(self.command_choices.currentItem().text())
        else:
            self.command_description.clear()

    def _update_command_description(self, option_text: str) -> None:
        option = option_text.lstrip("/").strip()
        self.command_description.setText(COMMAND_DESCRIPTIONS.get(option, ""))

    def _move_command_selection(self, delta: int) -> None:
        if not self.command_panel.isVisible() or self.command_choices.count() == 0:
            return
        row = max(0, min(self.command_choices.count() - 1, self.command_choices.currentRow() + delta))
        self.command_choices.setCurrentRow(row)

    def accept_command_palette(self) -> None:
        if not self.command_panel.isVisible():
            return
        typed = self.command_line.text().strip()
        result: str
        if " " in typed.lstrip("/"):
            result = typed
        elif self.command_choices.currentItem():
            result = self.command_choices.currentItem().text()
        else:
            result = typed
        self.close_command_palette()
        if result:
            self.execute_command(result)

    def copy_current_selection(self) -> None:
        view = self._current_view()
        if view:
            view.copy_selection_to_clipboard()

    def paste_into_current(self) -> None:
        view = self._current_view()
        if view:
            view.paste_from_clipboard()
            self._sync_title()

    def _toggle_style_shortcut(self, style: str) -> None:
        view = self._current_view()
        if view:
            view.apply_format(style)
            self._sync_title()

    def select_current_row(self) -> None:
        view = self._current_view()
        if view and view.table.currentRow() >= 0:
            view.table.selectRow(view.table.currentRow())

    def select_current_column(self) -> None:
        view = self._current_view()
        if view and view.table.currentColumn() >= 0:
            view.table.selectColumn(view.table.currentColumn())

    def _parse_recent_files(self, raw: str) -> list[str]:
        try:
            import json
            loaded = json.loads(raw)
            if isinstance(loaded, list):
                return [str(item) for item in loaded]
        except Exception:
            pass
        return []

    def _current_view(self) -> SheetView | None:
        widget = self.tabs.currentWidget()
        return widget if isinstance(widget, SheetView) else None

    def _sync_title(self) -> None:
        view = self._current_view()
        if not view:
            self.setWindowTitle("gui-ss")
            return
        name = view.path.name if view.path else "Untitled"
        star = "*" if view.dirty else ""
        self.setWindowTitle(f"gui-ss{star} - {name}")

    def new_sheet(self) -> None:
        sheet = ensure_sheet_defaults(Spreadsheet(rows=100, cols=52), self.defaults)
        self._add_tab(sheet, None, "Untitled")

    def _add_tab(self, sheet: Spreadsheet, path: Path | None, label: str) -> None:
        view = SheetView(sheet, path)
        view.table.itemChanged.connect(lambda _item, self=self: self._sync_title())
        index = self.tabs.addTab(view, label)
        self.tabs.setCurrentIndex(index)
        self._sync_title()

    def open_sheet(self, path: Path) -> None:
        sheet = load_sheet(path, defaults=self.defaults)
        self._add_tab(sheet, path, path.name)
        self._remember_recent(path)

    def save_current(self) -> None:
        view = self._current_view()
        if not view:
            return
        target = view.path
        if target is None:
            self.save_current_as()
            return
        save_sheet(view.sheet, target)
        view.dirty = False
        self._sync_title()

    def save_current_as(self) -> None:
        view = self._current_view()
        if not view:
            return
        file_name, _selected = QFileDialog.getSaveFileName(self, "Save Sheet", str(view.path or DEFAULT_PATH), "Sheets (*.tss *.csv *.tsv)")
        if not file_name:
            return
        target = Path(file_name).expanduser()
        save_sheet(view.sheet, target)
        view.path = target
        view.dirty = False
        self.tabs.setTabText(self.tabs.currentIndex(), target.name)
        self._remember_recent(target)
        self._sync_title()

    def close_tab(self, index: int) -> None:
        widget = self.tabs.widget(index)
        if not isinstance(widget, SheetView):
            return
        if widget.dirty:
            answer = QMessageBox.question(self, "Unsaved changes", f"Save changes to {widget.path.name if widget.path else 'Untitled'}?")
            if answer == QMessageBox.StandardButton.Yes:
                self.tabs.setCurrentIndex(index)
                self.save_current()
        self.tabs.removeTab(index)
        if self.tabs.count() == 0:
            self.close()
        else:
            self._sync_title()

    def quit_current(self) -> None:
        index = self.tabs.currentIndex()
        if index >= 0:
            self.close_tab(index)

    def _remember_recent(self, path: Path) -> None:
        value = str(path.expanduser())
        self.recent_files = [item for item in self.recent_files if item != value]
        self.recent_files.insert(0, value)
        self.recent_files = self.recent_files[:10]
        save_app_settings(
            self.settings_path,
            {
                **self.defaults,
                "recent_files_json": __import__("json").dumps(self.recent_files),
            },
        )

    def execute_command(self, text: str) -> None:
        command = parse_command(text)
        view = self._current_view()
        if command.name == "load":
            if command.args:
                self.open_sheet(Path(command.args[0]).expanduser())
            else:
                start_dir = Path(self.recent_files[0]).expanduser().parent if self.recent_files else Path.cwd()
                file_name, _selected = QFileDialog.getOpenFileName(self, "Open Sheet", str(start_dir), "Sheets (*.tss *.csv *.tsv)")
                if file_name:
                    self.open_sheet(Path(file_name))
            return
        if command.name == "save":
            if command.args:
                if view:
                    view.path = Path(command.args[0]).expanduser()
                    self.tabs.setTabText(self.tabs.currentIndex(), view.path.name)
                    self.save_current()
            else:
                self.save_current()
            return
        if command.name == "saveas":
            if command.args and view:
                view.path = Path(command.args[0]).expanduser()
                self.tabs.setTabText(self.tabs.currentIndex(), view.path.name)
                self.save_current()
            else:
                self.save_current_as()
            return
        if command.name == "quit":
            self.quit_current()
            return
        if command.name == "help":
            language = view.sheet.language if view else "en"
            topic = command.args[0].lower() if command.args else "commands"
            if topic == "formulas":
                lines = get_formula_help_lines(language)
            elif topic == "keys":
                lines = get_key_help_lines(language)
            else:
                lines = get_command_help_lines(language)
            HelpDialog(self, topic.title(), lines).exec()
            return
        if command.name == "edit":
            if not view:
                return
            dialog = SettingsDialog(self, view.sheet, self.settings_path, self.recent_files)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                dialog.apply()
                view._apply_theme()
                view._populate()
                self.defaults = load_app_settings(self.settings_path)
            return
        if command.name == "goto" and view and command.args:
            if CELL_REF_RE.match(command.args[0]):
                row, col = parse_cell_reference(command.args[0])
                view.table.setCurrentCell(row, col)
                view.table.scrollToItem(view.table.item(row, col))
            return
        if command.name == "blank" and view:
            view.clear_selection()
            self._sync_title()
            return
        if command.name == "justify" and view:
            align = command.args[0].lower() if command.args else ""
            if not align:
                align, ok = QInputDialog.getItem(self, "Justify", "Alignment", JUSTIFY_OPTIONS, 0, False)
                if not ok:
                    return
            view.set_alignment(align)
            self._sync_title()
            return
        if command.name == "protect" and view:
            view.set_protection(True)
            self._sync_title()
            return
        if command.name == "unprotect" and view:
            view.set_protection(False)
            self._sync_title()
            return
        if command.name == "insert" and view:
            axis = command.args[0].lower() if command.args else ""
            if not axis:
                axis, ok = QInputDialog.getItem(self, "Insert", "Insert", ["column", "row"], 0, False)
                if not ok:
                    return
            if axis in {"row", "rows"}:
                row = max(0, view.table.currentRow())
                count = int(command.args[2]) if len(command.args) > 2 and command.args[2].isdigit() else 1
                view.insert_rows(row, count)
            else:
                col = max(0, view.table.currentColumn())
                count = int(command.args[2]) if len(command.args) > 2 and command.args[2].isdigit() else 1
                view.insert_columns(col, count)
            self._sync_title()
            return
        if command.name == "delete" and view:
            axis = command.args[0].lower() if command.args else ""
            if not axis:
                axis, ok = QInputDialog.getItem(self, "Delete", "Delete", ["column", "row"], 0, False)
                if not ok:
                    return
            if axis in {"row", "rows"}:
                row = max(0, view.table.currentRow())
                count = int(command.args[2]) if len(command.args) > 2 and command.args[2].isdigit() else 1
                view.delete_rows(row, count)
            else:
                col = max(0, view.table.currentColumn())
                count = int(command.args[2]) if len(command.args) > 2 and command.args[2].isdigit() else 1
                view.delete_columns(col, count)
            self._sync_title()
            return
        if command.name == "global" and view:
            if len(command.args) >= 2 and command.args[0].lower() == "width":
                if len(command.args) == 2:
                    width = int(command.args[1])
                    view.sheet.column_width = max(8, width)
                elif len(command.args) >= 3 and CELL_REF_RE.match(f"{command.args[1].upper()}1"):
                    col = parse_cell_reference(f"{command.args[1].upper()}1")[1]
                    view.sheet.set_column_width(col, int(command.args[2]))
                view._populate()
                self._sync_title()
            return
        if command.name == "output" and view:
            self._output_current(view, command.args)
            return
        if command.name == "format" and view:
            self._execute_format_command(view, command.args)
            return

    def _output_current(self, view: SheetView, args: list[str]) -> None:
        mode = args[0].lower() if args else "file"
        if mode == "screen":
            lines = sheet_to_text_lines(view.sheet, view.evaluator, view._display_value)
            HelpDialog(self, "Output", lines).exec()
            return
        target: Path | None = Path(args[-1]).expanduser() if len(args) >= 1 and mode != "screen" and len(args) >= 2 else None
        if target is None:
            file_name, _selected = QFileDialog.getSaveFileName(self, "Output Sheet", str(view.path or DEFAULT_PATH), "Text (*.txt);;CSV (*.csv);;TSV (*.tsv);;PDF (*.pdf)")
            if not file_name:
                return
            target = Path(file_name).expanduser()
        if target.suffix.lower() == ".pdf":
            save_pdf_text(sheet_to_text_lines(view.sheet, view.evaluator, view._display_value), target, title=target.name)
        elif target.suffix.lower() in {".csv", ".tsv"}:
            save_sheet(view.sheet, target)
        else:
            target.write_text("\n".join(sheet_to_text_lines(view.sheet, view.evaluator, view._display_value)) + "\n", encoding="utf-8")

    def _execute_format_command(self, view: SheetView, args: list[str]) -> None:
        if not args:
            style, ok = QInputDialog.getItem(self, "Format", "Choose format", sorted(FORMAT_STYLES + ["font-size"]), 0, False)
            if not ok:
                return
            args = [style]
        style = args[0].lower()
        if style == "date":
            date_style = args[1] if len(args) > 1 else None
            if not date_style:
                date_style, ok = QInputDialog.getItem(self, "Date Format", "Date format", DATE_FORMATS, 0, False)
                if not ok:
                    return
            view.apply_format("date", date_style)
            return
        if style == "currency":
            symbol = args[1] if len(args) > 1 else None
            if not symbol:
                symbol, ok = QInputDialog.getItem(self, "Currency", "Symbol", CURRENCY_SYMBOLS, 0, False)
                if not ok:
                    return
            view.apply_format("currency", symbol)
            return
        if style in {"background", "b"}:
            color = args[1] if len(args) > 1 else None
            if not color:
                color, ok = QInputDialog.getItem(self, "Background", "Background", BACKGROUND_COLORS, 0, False)
                if not ok:
                    return
            view.apply_format("background", color)
            return
        if style == "font-size":
            if len(args) > 1:
                size = int(args[1])
            else:
                size, ok = QInputDialog.getInt(self, "Font Size", "Point size", 14, 6, 72)
                if not ok:
                    return
            view.apply_format("font-size", str(size))
            return
        view.apply_format(style)


def gui_main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="gui-ss", description="GUI spreadsheet front end for tui-ss")
    parser.add_argument("paths", nargs="*", help="Optional sheets to open")
    parser.add_argument("--settings", default=str(DEFAULT_SETTINGS_PATH), help="Path to the settings TOML file")
    args = parser.parse_args(argv)

    app = QApplication(sys.argv[:1] + (argv if argv is not None else sys.argv[1:]))
    app.setApplicationName("gui-ss")
    app.setOrganizationName("OpenAI")
    app.setDesktopFileName("org.ai.accounts")
    paths = [Path(item).expanduser() for item in args.paths]
    window = GuiSpreadsheetWindow(paths=paths, settings_path=Path(args.settings).expanduser())
    window.show()
    return app.exec()
