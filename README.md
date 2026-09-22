# Excel Micro

Excel Micro is a browser-based spreadsheet workspace built from scratch as a modern, lightweight Excel-style application.

## Excel Micro 0.7

The current release is a real working spreadsheet core, not a static spreadsheet mockup.

### Spreadsheet workspace

- 200 × 52 editable sheet grid (A–AZ)
- Click, drag, and Shift-select ranges
- Click row headers to select complete rows
- Click column headers to select complete columns
- Click the top-left corner to select the whole sheet
- Arrow keys, Tab, Enter, F2, Delete, and Backspace navigation
- Editable Name Box
  - Jump directly to a cell such as `B12`
  - Select a range by entering `A1:D20`
- Formula bar and direct in-cell editing
- Selection status with Average, Count, and Sum
- Zoom from 60% to 180%

### Excel-style formulas

Excel Micro now uses its own safe formula parser instead of executing formula text as JavaScript.

Supported examples include:

```text
=A1+B1*2
=A1^2
=A1&" units"
=SUM(A1:A10)
=AVERAGE(B2:B20)
=MIN(A1:A20)
=MAX(A1:A20)
=COUNT(A1:A20)
=COUNTA(A1:A20)
=PRODUCT(A1:A5)
=MEDIAN(A1:A10)
=ROUND(A1*B1,2)
=ROUNDUP(A1,0)
=ROUNDDOWN(A1,0)
=ABS(A1)
=SQRT(A1)
=POWER(A1,2)
=MOD(A1,2)
=IF(SUM(A1:A5)>100,"Over","OK")
=AND(A1>0,B1>0)
=OR(A1="Yes",B1="Yes")
=NOT(A1=0)
=LEN(A1)
=UPPER(A1)
=LOWER(A1)
=TRIM(A1)
=LEFT(A1,3)
=RIGHT(A1,3)
=MID(A1,2,4)
=CONCAT(A1," ",B1)
```

Nested formulas, comparisons, percentages, ranges, strings, booleans, arithmetic operators, exponentiation, and text concatenation are supported.

### Fill behavior

- Fill Down
- Fill Right
- Relative formula references shift as formulas are filled
- Absolute references such as `$A$1` remain fixed
- Keyboard shortcuts:
  - `Ctrl/Cmd + D` Fill Down
  - `Ctrl/Cmd + R` Fill Right

### Formatting

- Bold
- Italic
- Underline
- Strikethrough
- Font sizes
- Left / center / right alignment
- Wrap text
- Text color
- Fill color
- General formatting
- Number formatting
- Currency formatting
- Percent formatting
- Decimal increase / decrease
- Clear values
- Clear formatting

### Rows, columns, and sheet view

- Insert rows above the active selection
- Delete rows
- Insert columns to the left of the active selection
- Delete columns
- Formula references are adjusted when rows or columns move
- Drag a column border to resize it
- Double-click a column border to AutoFit
- Drag a row border to resize it
- Double-click a row border to AutoFit
- Widen / narrow selected columns from the View ribbon
- Increase / decrease selected row heights
- Reset selected column widths and row heights
- Imported Excel column widths are read when available
- Exported Excel workbooks include column widths
- Toggle gridlines
- Freeze top row
- Freeze first column
- Freeze both together

### Worksheets

- Multiple worksheets
- New worksheet
- Rename worksheet
- Duplicate worksheet
- Delete worksheet
- Sheet tabs
- Local workbook persistence
- Undo / redo history

### Excel-style context menus

Right-click directly where you are working:

- Cells: Copy, Clear contents, Clear formatting, Insert row, Insert column
- Row headers: Insert row, Delete row, AutoFit row, Increase row height
- Column headers: Insert column, Delete column, AutoFit column, Increase column width

### Data tools

- Excel-style filter headers for a selected range
- Per-column filter rules: Contains, Equals, Greater than, Less than, Is not blank
- Clear one filter or all filter conditions
- Sort selected data A → Z
- Sort selected data Z → A
- Find within the active sheet
- Copy selected ranges
- Multi-cell paste
- Quick bar, line, and pie charts
- Switch chart type without rebuilding the selection

### Structured tables

- Format any selected dataset as a table
- Named tables
- Green, blue, orange, and gray style presets
- Optional banded rows
- First row becomes a styled table header
- Table creation automatically enables filter headers
- Table ranges move with inserted/deleted rows and columns
- Manage and remove tables from the workbook UI

### Conditional formatting

- Apply rules to any selected range
- Greater than
- Less than
- Equals
- Contains text
- Is not blank
- Custom fill and text colors
- Multiple saved rules per sheet
- Remove individual rules or clear all

### Named ranges

- Name any selected range from the Formulas ribbon
- Use names directly in formulas such as `=SUM(SALES)`
- Jump to named ranges by typing the name in the Name Box
- Named ranges shift automatically with row/column insertion and deletion
- Manage and remove names from the workbook UI

### Data validation and dropdown cells

- Apply dropdown validation to any selected range
- Define allowed values with one item per line or comma-separated
- Optional blank values
- In-cell dropdown picker
- Invalid manually typed values are rejected
- Invalid pasted values are rejected without overwriting valid cells
- Validation ranges shift with inserted/deleted rows and columns
- Clear validation from the selected range

### File compatibility

- Import `.xlsx`
- Import `.xls`
- Export `.xlsx`
- Import `.csv`
- Export `.csv`
- Excel Micro JSON backup and restore

### Local-first autosave

The active workbook is automatically stored in browser local storage. The title bar shows saving status as edits are made.

## Keyboard shortcuts

| Shortcut | Action |
| --- | --- |
| Ctrl/Cmd + Z | Undo |
| Ctrl/Cmd + Y | Redo |
| Ctrl/Cmd + C | Copy |
| Ctrl/Cmd + F | Find |
| Ctrl/Cmd + S | Export workbook |
| Ctrl/Cmd + A | Select all |
| Ctrl/Cmd + B | Bold |
| Ctrl/Cmd + I | Italic |
| Ctrl/Cmd + U | Underline |
| Ctrl/Cmd + D | Fill Down |
| Ctrl/Cmd + R | Fill Right |
| F2 | Edit active cell |
| Enter | Edit / commit |
| Tab | Move horizontally |
| Shift + Arrow | Extend selection |

## Run locally

```bash
npm install
npm run dev
```

## Production

```bash
npm install
npm run build
npm start
```

The production build is emitted to `dist/`. The included Node server serves the compiled app and honors the `PORT` environment variable.

## Railway

The repository includes `railway.json`.

- Build: `npm install --no-audit --no-fund && npm run build`
- Start: `npm run start`
- No secrets are required for the current local-first release.

## Verification

GitHub Actions runs on every push to `main` and on pull requests. CI:

1. installs dependencies
2. runs TypeScript compilation
3. builds the Vite production app
4. starts the production server
5. performs a real HTTP smoke test

## Architecture

- React
- TypeScript
- Vite
- SheetJS 0.20.3
- Custom spreadsheet renderer
- Custom selection model
- Custom workbook state
- Custom safe formula parser
- Browser localStorage persistence
- Lightweight Node production server

## Next major phases

Excel Micro 0.7 is still local-first. The next major product layers are:

1. Accounts and cloud workbooks
2. Share links and permissions
3. Real-time multiplayer editing
4. Comments and version history
5. More Excel-compatible formulas
6. Pivot tables
7. Print and page layout
8. Large-sheet virtualization and worker-based calculations
