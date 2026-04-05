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
    "h": "help",
    "q": "quit",
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
    "replicate",
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
    "/G width n             Set global column width.",
    "/G width col n         Set one column width, for example /G width B 18.",
    "/H                     Open help topics.",
    "/I row|col index [n]   Insert rows or columns.",
    "/J left|centre|right [range]  Justify left, centre, or right.",
    "/L file                Load a .tss or .csv sheet.",
    "/M row|col a b [n]     Move rows or columns.",
    "/O screen|file path    Output to screen, .txt, or .csv.",
    "/P [range]             Protect cells from editing.",
    "/Q                     Quit.",
    "/R src dst             Replicate formulas/contents to destination.",
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
    "/                      Open the command menu.",
    "/H                     Open help topics.",
    "Esc                    Cancel prompt or menu.",
    "Alt+=                  Insert SUM of numeric cells above.",
    "Mouse click            Move active cell.",
    "Mouse drag             Select a rectangular range.",
]

FORMULA_HELP_LINES = [
    "Math/Stats: ABS, AVERAGE, AVG, COS, COUNT, INT, MAX, MIN, ROUND, SUM",
    "Logical: IF(condition, then_value, else_value)",
    "Data: LOOKUP(value, lookup_range, result_range)",
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
