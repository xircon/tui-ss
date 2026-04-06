#!/usr/bin/env markdown
# tui-ss

A modular terminal spreadsheet with a SuperCalc-style slash command workflow.

## Controls

- Arrow keys or `hjkl`: move
- `Shift` + arrows: grow/shrink a rectangular selection
- `Enter`: edit the current cell
- `Tab`: move right
- `Ctrl+S`: save
- `Ctrl+Q`: quit
- `Ctrl+C`: copy the current cell or selection
- `Ctrl+V` or `Ctrl+Y`: paste the copied block
- `Ctrl+Z`: undo
- `Ctrl+R`: redo
- `Ctrl+E` or `F2`: edit the current cell in the formula bar
- `/`: open the slash command menu
- `Esc`: cancel the current prompt
- Click a tab at the top: switch open files

## Slash Help

Pressing `/` now shows the command reference while you type. You can use either the long form command names or the classic one-letter forms.

- `/A range [col] [desc]`: arrange and sort rows in a range
- `/B [range]`: blank the current cell or a range
- `/C src dst`: copy a source range to a destination
- `/D row|col index [n]`: delete rows or columns
- `/E [cell] value`: edit the current or named cell
- `/F style [range]`: format cells as `clear-format`, `text`, `currency`, `fixed`, `percent`, `int`, `negative`, `accounting`, or `sci`
- `/F date`: change the whole-sheet date format with a horizontal `european` / `us` / `ansi` menu
- `/F`: open a horizontal format menu you can drive with arrows or typing
- `/FIND text [range]`: find the next matching cell
- `/G width n`: set the global column width
- `/G width B 18`: set one column width
- `/G width B:D 18`: set a column range width
- `/J left|centre|right [range]`: justify cells left, centre, or right
- `/I row|col index [n]`: insert rows or columns
- `/L file`: load a `.tss`, `.csv`, or `.tsv` file in a new tab
- `/M row|col a b [n]`: move rows or columns
- `/O screen|file path`: output to the screen or a `.csv` / `.tsv` / `.pdf` / text snapshot
- `/P [range]`: protect cells from editing
- `/Q`: quit
- `/REDO`: redo the last undone action
- `/R src dst`: replicate a source range into a destination block
- `/REPLACE old new [range]`: replace raw text in cells
- `/S [file]`: save directly, or with no argument open a `save` / `save-as` / `save-quit` menu
- `/SAVEAS file`: save the sheet to a new file
- `/T rows [cols]`: freeze title rows and columns
- `/U [range]`: remove protection
- `/V`: open the theme menu and save the choice with the file
- `/V cyan|yellow|magenta|blue|white|purple`: set an exact theme
- `/W`: toggle the command/help window summary
- `/X file`: execute commands from a text file
- `/Z`: zap the whole workspace
- `/GO cell`: jump to a cell like `B12`

## Examples

- `/E A1 125`
- `/E B1 =A1*2`
- `/FIND Banana`
- `/REPLACE old new A:A`
- `/C A1:B3 D1`
- `/C A1 B1:B10`
- `/F currency A1:B10`
- `/F date`
- `/F clear-format A1:C10`
- `/F negative B1:B10`
- `/F accounting C1:C10`
- `/F`
- `/J centre A1:C10`
- `/V`
- `/V cyan`
- `/G width B 18`
- `/G width B:D 18`
- `/A A1:C20 1`
- `/I row 5 2`
- `/D col 3 1`
- `/P A1:C3`
- `/U B2`
- `/S ~/sheets/budget.tss`
- `/O file ~/sheets/budget.pdf`
- `/SAVEAS ~/sheets/budget-copy.tss`
- `/L ~/sheets/budget.tss`
- `/X ~/scripts/tui-ss/demo.commands`

## Notes

- `.tss` files store cells, formats, protection, title freeze settings, column width, theme choice, and alignment metadata.
- You can open multiple files at once; each file gets its own tab at the top.
- The sheet has one date format for display and input. Default is European `dd/mm/yyyy`.
- Individual column widths are saved in `.tss` files.
- CSV and TSV load/save are supported.
- Protected cells are marked with a leading `!` in the grid.
- Title rows and columns are emphasized in the display.
- Clicking a row header freezes rows through that row.
- Clicking a column header freezes columns through that column.
- Formula cells are drawn in bright green.
- `negative` format draws negative numbers in red.
- `accounting` format shows negative numbers in brackets.
- `Alt+=` inserts a `SUM(...)` formula for the numeric cells above the current cell.
- Formula functions include `ABS`, `AVERAGE`, `AVG`, `COS`, `COUNT`, `IF`, `INT`, `LOOKUP`, `MAX`, `MIN`, `ROUND`, and `SUM`.
- Date functions include `DATE`, `TODAY`, `YEAR`, `MONTH`, `DAY`, `DATEDIFF`, and `WEEKDAY`.
- With the default European sheet format, entering `05/04/2026` stores the date and displays it as `05/04/2026`.
- Formula entry supports arrow-key pointing after `=` and inside ranges such as `=SUM(` ... `:` ... `Enter`.
- Absolute references are supported: `$A$1`, `$A1`, and `A$1`.
- Examples: `=IF(A1=10,1,0)`, `=AVERAGE(B1:B5)`, `=COS(0)`, `=LOOKUP("Fred",A1:A10,B1:B10)`.
