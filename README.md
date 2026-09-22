# Excel Micro

Excel Micro is a browser-based spreadsheet workspace built from scratch as a modern, lightweight Excel-style application.

## Excel Micro 0.3

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

### Columns and sheet view

- Drag a column border to resize it
- Double-click a column border to AutoFit
- Widen / narrow selected columns from the View ribbon
- Reset selected column widths
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

### Data tools

- Sort selected data A → Z
- Sort selected data Z → A
- Find within the active sheet
- Copy selected ranges
- Multi-cell paste
- Quick bar charts

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

Excel Micro 0.3 is still local-first. The next major product layers are:

1. Accounts and cloud workbooks
2. Share links and permissions
3. Real-time multiplayer editing
4. Comments and version history
5. Filters and structured table objects
6. Conditional formatting
7. More Excel-compatible formulas
8. Row insertion/deletion and row resizing
9. Named ranges
10. Advanced chart types
11. Pivot tables
12. Print and page layout
13. Large-sheet virtualization and worker-based calculations
