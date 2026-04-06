#!/usr/bin/env python3
"""Slash command parsing."""

from __future__ import annotations

import shlex
from dataclasses import dataclass


@dataclass(slots=True)
class Command:
    name: str
    args: list[str]


ALIASES = {
    "a": "arrange",
    "b": "blank",
    "c": "copy",
    "d": "delete",
    "find": "find",
    "h": "help",
    "q": "quit",
    "redo": "redo",
    "replace": "replace",
    "e": "edit",
    "f": "format",
    "g": "global",
    "i": "insert",
    "j": "justify",
    "l": "load",
    "m": "move",
    "o": "output",
    "p": "protect",
    "r": "replicate",
    "s": "save",
    "saveas": "saveas",
    "t": "title",
    "u": "unprotect",
    "v": "theme",
    "w": "window",
    "x": "execute",
    "z": "zap",
    "go": "goto",
    "open": "load",
    "clear": "blank",
    "?": "help",
}

COMMAND_MENU_OPTIONS = [
    "arrange",
    "blank",
    "copy",
    "delete",
    "edit",
    "execute",
    "find",
    "format",
    "global",
    "goto",
    "help",
    "insert",
    "justify",
    "load",
    "move",
    "output",
    "protect",
    "quit",
    "redo",
    "replicate",
    "replace",
    "save",
    "theme",
    "title",
    "unprotect",
    "zap",
]

HELP_TOPICS = ["commands", "formulas", "keys"]

COMMAND_HELP_LINES = [
    "/A range [col] [desc]  Sort rows in a range by column offset.",
    "/B [range]             Blank a cell or range.",
    "/C src dst             Copy a range to a destination.",
    "/D row|col index [n]   Delete rows or columns.",
    "/E [cell] value        Edit current or named cell.",
    "/F                     Open the format menu.",
    "/F DATE                Set the whole-sheet date format.",
    "/FIND text [range]     Find next matching cell.",
    "/G width n             Set global column width.",
    "/G width col n         Set one column width, for example /G width B 18.",
    "/H                     Open help topics.",
    "/I row|col index [n]   Insert rows or columns.",
    "/J left|centre|right [range]  Justify left, centre, or right.",
    "/L file                Load a .tss/.csv/.tsv sheet in a new tab.",
    "/M row|col a b [n]     Move rows or columns.",
    "/O screen|file path    Output to screen, .txt, .csv, .tsv, or .pdf.",
    "/P [range]             Protect cells from editing.",
    "/Q                     Quit.",
    "/REDO                  Redo the last undone action.",
    "/R src dst             Replicate formulas/contents to destination.",
    "/REPLACE old new [range] Replace raw cell text.",
    "/S [file]              Save file, or open save/save-as/save-quit menu.",
    "/SAVEAS file           Save sheet to a new file.",
    "/T rows [cols]         Freeze title rows and columns.",
    "/U [range]             Remove protection.",
    "/V [theme]             Open theme menu, or set white/cyan/yellow/magenta/blue/purple.",
    "/W                     Open command help.",
    "/X file                Execute commands from a file.",
    "/Z                     Clear the whole workspace.",
    "/GO cell               Jump to a cell.",
]

KEY_HELP_LINES = [
    "Arrow keys or hjkl     Move the active cell.",
    "Enter                  Edit current cell, store, then move down.",
    "Tab                    Move right.",
    "Delete                 Clear current cell contents.",
    "Ctrl+Q                 Quit.",
    "Ctrl+E or F2           Edit current cell in the formula bar.",
    "Ctrl+R                 Redo the last undone action.",
    "/                      Open the command menu.",
    "/H                     Open help topics.",
    "Esc                    Cancel prompt or menu.",
    "Alt+=                  Insert SUM of numeric cells above.",
    "Mouse click            Move active cell.",
    "Mouse drag             Select a rectangular range.",
    "Mouse click tab        Switch open files.",
    "Click row/col header   Freeze through that row or column.",
]

FORMULA_HELP_LINES = [
    "Math/Stats: ABS, AVERAGE, AVG, COS, COUNT, INT, MAX, MIN, ROUND, SUM",
    "Logical: IF(condition, then_value, else_value)",
    "Data: LOOKUP, VLOOKUP, HLOOKUP",
    "Dates: DATE, TODAY, YEAR, MONTH, DAY, DATEDIFF, WEEKDAY",
    "Sheet date display/input: /F DATE then choose european, us, or ansi",
    "",
    "Examples:",
    "=SUM(A1:A10)",
    "=AVERAGE(B1:B5)",
    "=ROUND(C1/7, 2)",
    "=ABS(D1)",
    "=COS(0)",
    "=IF(A1=10, 1, 0)",
    "=IF(B2<>0, B1/B2, 0)",
    '=LOOKUP("Fred", A2:A10, B2:B10)',
    '=VLOOKUP("Fred", A2:C10, 2)',
    '=HLOOKUP("Q2", A1:D4, 3)',
    "=DATE(2026, 4, 5)",
    "=TODAY()",
    "=YEAR(A1)",
    "=MONTH(A1)",
    "=DAY(A1)",
    "=DATEDIFF(A1, B1)",
    "=WEEKDAY(A1)",
    "",
    "Absolute refs: $A$1  $A1  A$1",
]


def parse_command(text: str) -> Command:
    cleaned = text.strip()
    if cleaned.startswith("/"):
        cleaned = cleaned[1:]
    parts = shlex.split(cleaned)
    if not parts:
        return Command(name="help", args=[])
    name = ALIASES.get(parts[0].lower(), parts[0].lower())
    return Command(name=name, args=parts[1:])
