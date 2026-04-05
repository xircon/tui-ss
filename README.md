#!/usr/bin/env markdown
# tui-ss

A modular terminal spreadsheet with a SuperCalc-style slash command workflow.

## Controls

- Arrow keys or `hjkl`: move
- `Enter`: edit the current cell
- `Tab`: move right
- `/`: open the slash command prompt and show the in-app help panel
- `Esc`: cancel the current prompt

## Slash Help

Pressing `/` now shows the command reference while you type. You can use either the long form command names or the classic one-letter forms.

- `/A range [col] [desc]`: arrange and sort rows in a range
- `/B [range]`: blank the current cell or a range
- `/C src dst`: copy a source range to a destination
- `/D row|col index [n]`: delete rows or columns
- `/E [cell] value`: edit the current or named cell
- `/F style [range]`: format cells as `clear-format`, `text`, `currency`, `fixed`, `percent`, `int`, `negative`, `accounting`, or `sci`
- `/F`: open a horizontal format menu you can drive with arrows or typing
- `/G width n`: set the global column width
- `/G width B 18`: set one column width
- `/J left|centre|right [range]`: justify cells left, centre, or right
- `/I row|col index [n]`: insert rows or columns
- `/L file`: load a `.tss` or `.csv` file
- `/M row|col a b [n]`: move rows or columns
- `/O screen|file path`: output to the screen, a `.txt` snapshot, or `.csv`
- `/P [range]`: protect cells from editing
- `/Q`: quit
- `/R src dst`: replicate a source range into a destination block
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
- `/C A1:B3 D1`
- `/F currency A1:B10`
- `/F clear-format A1:C10`
- `/F negative B1:B10`
- `/F accounting C1:C10`
- `/F`
- `/J centre A1:C10`
- `/V`
- `/V cyan`
- `/G width B 18`
- `/A A1:C20 1`
- `/I row 5 2`
- `/D col 3 1`
- `/P A1:C3`
- `/U B2`
- `/S ~/sheets/budget.tss`
- `/SAVEAS ~/sheets/budget-copy.tss`
- `/L ~/sheets/budget.tss`
- `/X ~/scripts/tui-ss/demo.commands`

## Notes

- `.tss` files store cells, formats, protection, title freeze settings, column width, and theme choice.
- Individual column widths are saved in `.tss` files.
- CSV load and save are supported.
- Protected cells are marked with a leading `!` in the grid.
- Title rows and columns are emphasized in the display.
- Formula cells are drawn in bright green.
- `negative` format draws negative numbers in red.
- `accounting` format shows negative numbers in brackets.
- The second row shows quick format and theme options at all times.
- The grid uses subtle separators so cell boundaries are easier to track.
- `Alt+=` inserts a `SUM(...)` formula for the numeric cells above the current cell.
- Formula functions include `ABS`, `AVERAGE`, `AVG`, `COS`, `COUNT`, `IF`, `INT`, `LOOKUP`, `MAX`, `MIN`, `ROUND`, and `SUM`.
- Examples: `=IF(A1=10,1,0)`, `=AVERAGE(B1:B5)`, `=COS(0)`, `=LOOKUP("Fred",A1:A10,B1:B10)`.
