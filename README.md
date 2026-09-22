# Excel Micro

Excel Micro is a browser-based spreadsheet workspace built from scratch as a modern, lightweight alternative for everyday spreadsheet work.

## What works now

- Editable 200 × 52 spreadsheet grid (A–AZ)
- Mouse selection, shift-selection, drag-selection, arrow-key navigation, Tab, Enter, F2, Delete
- Formula bar and in-cell editing
- Formulas and cell references
  - Arithmetic (`=A1+B1*2`)
  - Ranges (`=SUM(A1:A10)`)
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
- Local autosave in the browser
- Responsive Excel-inspired ribbon UI
- Adjustable zoom
- Selection statistics: average, count, and sum

## Run locally

```bash
npm install
npm run dev
```

Then open the local Vite URL printed in the terminal.

## Production build

```bash
npm run build
npm run preview
```

The production files are emitted to `dist/`.

## Deploy

Excel Micro is a standard Vite app and can be deployed to Railway, Vercel, Netlify, Cloudflare Pages, GitHub Pages, or any static host. Build command: `npm run build`. Output directory: `dist`.

## Architecture

- React + TypeScript
- Vite
- SheetJS (`xlsx`) for Excel workbook file compatibility
- Custom spreadsheet grid, formula engine, selection model, and workbook state
- Browser `localStorage` for automatic local persistence

## Next platform milestones

The current repository is the usable core spreadsheet product. Future phases can add collaborative accounts, server-side workbooks, comments, pivot tables, conditional formatting, advanced charts, formula expansion, table objects, freeze panes, row/column resizing, printable page layout, permissions, version history, and real-time multiplayer editing.
