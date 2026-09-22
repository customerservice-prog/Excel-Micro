# Excel Micro

Excel Micro is a browser-based spreadsheet workspace built from scratch as a modern, lightweight alternative for everyday spreadsheet work.

## Current product

- Editable 200 × 52 spreadsheet grid (A–AZ)
- Mouse selection, shift-selection, drag-selection, arrow-key navigation, Tab, Enter, F2, Delete
- Formula bar and in-cell editing
- Formulas and cell references
  - Arithmetic such as `=A1+B1*2`
  - Ranges such as `=SUM(A1:A10)`
  - `SUM`, `AVERAGE` / `AVG`, `MIN`, `MAX`, `COUNT`, `IF`
- Multiple worksheets with add, rename, delete, and tab switching
- Undo / redo history
- Copy and multi-cell paste
- Find in sheet
- Formatting
  - Bold, italic, underline
  - Text alignment
  - Text and fill colors
  - General, number, currency, and percent formats
- Data sorting for selected ranges
- Quick bar chart from a selected range
- Real `.xlsx` import and export
- CSV import/export
- Native Excel Micro JSON backup/restore
- Local browser autosave
- Excel-inspired responsive ribbon interface
- Adjustable zoom
- Selection statistics: average, count, and sum

## Run locally

```bash
npm install
npm run dev
```

## Production build

```bash
npm install
npm run build
npm start
```

The production build is emitted to `dist/`. The included Node static server serves the compiled app and uses the `PORT` environment variable when provided.

## Railway

The repository includes `railway.json`, so it can be connected directly to Railway.

- Build: `npm install --no-audit --no-fund && npm run build`
- Start: `npm run start`
- No application secrets are required for the current local-first release.

## Verification

GitHub Actions runs on every push to `main` and on pull requests. CI installs dependencies, runs the TypeScript/Vite production build, starts the production server, and requests the live app as a smoke test.

## Architecture

- React + TypeScript
- Vite
- SheetJS (`xlsx`) for Excel workbook compatibility
- Custom spreadsheet grid
- Custom formula engine
- Custom selection/workbook state
- Browser `localStorage` persistence
- Lightweight Node production static server

## Next milestones

The current repository is the usable spreadsheet core. The next major product phases are:

1. Accounts and cloud workbooks
2. Real-time collaboration and presence
3. Version history and comments
4. Conditional formatting
5. More Excel-compatible formulas
6. Row/column resizing, insertion, deletion, hiding, and freeze panes
7. Table objects, filters, and named ranges
8. Advanced chart types
9. Pivot tables
10. Print/page layout
11. Permissions and share links
12. Large-sheet virtualization and worker-based calculations
