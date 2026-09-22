# Excel Micro — Microsoft Excel Parity Roadmap

Excel Micro is intentionally being built toward a broad Excel-style feature set, but **full Microsoft Excel parity has not been reached**. This file is the engineering checklist for closing that gap.

Legend:

- ✅ Built and usable
- 🟡 Partial / simplified
- ⬜ Not yet built

## Workbook and worksheet basics

- ✅ Multiple worksheets
- ✅ Add / rename / duplicate / delete sheets
- ✅ Editable cells and formula bar
- ✅ Name Box navigation
- ✅ Whole-row / whole-column selection
- ✅ Insert / delete rows
- ✅ Insert / delete columns
- ✅ Row resizing and AutoFit
- ✅ Column resizing and AutoFit
- ✅ Freeze top row / first column
- ✅ Gridline toggle
- ✅ Local autosave
- ✅ Undo / redo
- ✅ Local workbook version history
- 🟡 Worksheet protection
  - Blocks edits through Excel Micro's mutation engine
  - No password / granular locked-cell permissions yet
- ⬜ Hide / unhide worksheets
- ⬜ Hide / unhide rows and columns
- ⬜ Move / copy worksheets between workbooks
- ⬜ Multiple workbook windows

## Cell editing and formatting

- ✅ Bold / italic / underline / strikethrough
- ✅ Font size
- ✅ Text and fill color
- ✅ Alignment
- ✅ Wrap text
- ✅ Number / currency / percent formatting
- ✅ Decimal controls
- ✅ Clear contents / formatting
- ✅ Copy and multi-cell paste
- ✅ Fill Down / Fill Right with relative formula shifting
- ✅ Absolute cell references during fill
- ⬜ Cell borders
- ⬜ Merge / unmerge cells
- ⬜ Format Painter
- ⬜ Custom number format language
- ⬜ Date/time format gallery
- ⬜ Accounting / scientific / fraction formats
- ⬜ Cell styles gallery
- ⬜ Themes / fonts / workbook colors

## Formulas and calculation engine

### Built

- ✅ Arithmetic and comparisons
- ✅ Cell references and ranges
- ✅ Named ranges
- ✅ Nested formulas
- ✅ Text concatenation with &
- ✅ SUM
- ✅ AVERAGE / AVG
- ✅ MIN / MAX
- ✅ COUNT / COUNTA
- ✅ PRODUCT
- ✅ MEDIAN
- ✅ ROUND / ROUNDUP / ROUNDDOWN
- ✅ ABS / SQRT / POWER / MOD
- ✅ IF / IFERROR
- ✅ AND / OR / NOT
- ✅ LEN / LOWER / UPPER / TRIM
- ✅ LEFT / RIGHT / MID
- ✅ CONCAT / CONCATENATE
- ✅ TEXTJOIN
- ✅ COUNTIF
- ✅ SUMIF
- ✅ AVERAGEIF
- ✅ XLOOKUP
- ✅ MATCH
- ✅ INDEX

### Remaining formula parity

- ⬜ SUMIFS / COUNTIFS / AVERAGEIFS
- ⬜ VLOOKUP / HLOOKUP
- ⬜ XMATCH
- ⬜ OFFSET / INDIRECT
- ⬜ Dynamic arrays: FILTER / SORT / SORTBY / UNIQUE / SEQUENCE
- ⬜ LET / LAMBDA
- ⬜ Date/time family
- ⬜ Financial family
- ⬜ Statistical family
- ⬜ Engineering family
- ⬜ Database family
- ⬜ Cube functions
- ⬜ Information functions
- ⬜ Full Excel error behavior and coercion
- ⬜ Cross-sheet references
- ⬜ Cross-workbook references
- ⬜ Calculation modes / dependency graph / incremental recalculation
- ⬜ Multi-threaded / worker calculation

## Tables, sorting, and filtering

- ✅ Sort selected range A→Z / Z→A
- ✅ AutoFilter range
- ✅ Per-column filter conditions
- ✅ Clear filters
- ✅ Format as Table
- ✅ Named tables
- ✅ Banded rows
- ✅ Multiple table style presets
- ✅ Table filter headers
- ✅ Table ranges move with structural edits
- 🟡 Structured tables
  - Visual and filter/table management exists
  - Structured-reference formulas such as Table1[Amount] are not implemented
- ⬜ Multi-level custom sort
- ⬜ Filter by color
- ⬜ Date/number-specific filter menus
- ⬜ Slicers
- ⬜ Timeline controls
- ⬜ Remove duplicates
- ⬜ Text to Columns
- ⬜ Flash Fill
- ⬜ Consolidate

## Conditional formatting

- ✅ Greater than
- ✅ Less than
- ✅ Equals
- ✅ Contains text
- ✅ Not blank
- ✅ Custom fill/text colors
- ✅ Multiple saved rules
- ✅ Rules move with inserted/deleted rows and columns
- ⬜ Data bars
- ⬜ Color scales
- ⬜ Icon sets
- ⬜ Top/bottom rules
- ⬜ Duplicate/unique rules
- ⬜ Formula-based rules
- ⬜ Rule priority / Stop If True

## Data validation

- ✅ Dropdown/list validation
- ✅ In-cell dropdown picker
- ✅ Allow blank
- ✅ Invalid typed values rejected
- ✅ Invalid pasted values rejected
- ✅ Validation ranges move with structural edits
- ⬜ Whole number / decimal validation
- ⬜ Date / time validation
- ⬜ Text length validation
- ⬜ Custom formula validation
- ⬜ Input messages
- ⬜ Configurable warning/error modes
- ⬜ Source list from named range or sheet range

## Notes, comments, and review

- ✅ Cell notes
- ✅ Note indicator
- ✅ Notes move with row/column structural edits
- ✅ Add/edit/delete note
- ⬜ Threaded comments
- ⬜ @mentions
- ⬜ Resolve / reopen comment threads
- ⬜ Author identity
- ⬜ Cloud comment synchronization
- ⬜ Track changes / show changes

## Charts and visualization

- ✅ Bar charts
- ✅ Line charts
- ✅ Pie / donut-style charts
- ✅ Chart type switcher
- 🟡 Charts are quick views generated from the active range
- ⬜ Persistent chart objects on worksheet canvas
- ⬜ Column charts
- ⬜ Area charts
- ⬜ Scatter / bubble
- ⬜ Combo charts
- ⬜ Histogram / Pareto
- ⬜ Box & whisker
- ⬜ Waterfall
- ⬜ Funnel
- ⬜ Treemap / sunburst
- ⬜ Radar / surface / stock charts
- ⬜ Sparklines
- ⬜ Chart titles / legends / axes editor
- ⬜ Series editor and secondary axes
- ⬜ Trendlines and error bars

## PivotTables and analysis

- ✅ Basic PivotTable-style summary builder
- ✅ Choose row field
- ✅ Choose value field
- ✅ Sum / Count / Average
- ✅ Output to a new worksheet
- 🟡 This is a simplified pivot summary, not full Excel PivotTable behavior
- ⬜ Multiple row / column fields
- ⬜ Filters fields
- ⬜ Multiple values
- ⬜ Grouping
- ⬜ Calculated fields/items
- ⬜ PivotTable styles
- ⬜ Pivot charts
- ⬜ Slicers
- ⬜ Refreshable pivot cache

## Printing and Page Layout

- ✅ Print active sheet used range
- ✅ Portrait / landscape
- ✅ Letter / A4
- ✅ Margin presets
- ✅ Print gridlines
- ✅ Workbook/sheet print title
- ⬜ Print area selection
- ⬜ Print titles / repeating rows
- ⬜ Header and footer editor
- ⬜ Page breaks
- ⬜ Page Break Preview
- ⬜ Scaling / Fit to pages
- ⬜ Center on page
- ⬜ Custom paper sizes

## File compatibility

- ✅ Import XLSX
- ✅ Import XLS
- ✅ Export XLSX
- ✅ Import / export CSV
- ✅ Excel Micro JSON backup / restore
- ✅ Basic formula import/export
- ✅ Column width import/export
- 🟡 Formatting interoperability is limited
- ⬜ Full Excel style round-trip
- ⬜ Charts round-trip
- ⬜ PivotTable round-trip
- ⬜ Data validation round-trip
- ⬜ Conditional formatting round-trip
- ⬜ Comments round-trip
- ⬜ Macro-enabled XLSM preservation
- ⬜ ODS import/export
- ⬜ PDF export

## Objects and insert tools

- ⬜ Images
- ⬜ Shapes
- ⬜ Text boxes
- ⬜ Icons
- ⬜ SmartArt-style diagrams
- ⬜ Hyperlinks
- ⬜ Checkboxes / form controls
- ⬜ Symbols / equations
- ⬜ Embedded objects

## Data import and transformation

- ⬜ Power Query-style query editor
- ⬜ CSV/text query connections
- ⬜ Web/API data imports
- ⬜ JSON/XML transformations
- ⬜ Database connections
- ⬜ Refreshable external data
- ⬜ Connection manager
- ⬜ Data Model / relationships
- ⬜ Power Pivot / DAX

## Automation and extensibility

- ⬜ VBA macro compatibility
- ⬜ Macro recorder
- ⬜ Office Scripts-style automation
- ⬜ Custom functions
- ⬜ Add-in/plugin model
- ⬜ Events / workbook scripting API

## Collaboration and cloud

- ✅ Local-first workbook autosave
- ✅ Local manual version restore points
- ⬜ User accounts
- ⬜ Cloud workbook storage
- ⬜ Share links
- ⬜ Viewer / editor permissions
- ⬜ Real-time multiplayer editing
- ⬜ Presence / cursors
- ⬜ Conflict resolution
- ⬜ Cloud version history
- ⬜ Offline/cloud synchronization

## Scale and performance

- 🟡 Current working grid is intentionally bounded to 200 × 52 for the custom renderer
- ⬜ Excel-scale row / column limits
- ⬜ Virtualized rows and columns
- ⬜ Worker-based formula calculations
- ⬜ Dependency graph recalculation
- ⬜ Efficient large-file streaming/import
- ⬜ Memory-aware workbook paging

## Accessibility and enterprise

- 🟡 Keyboard navigation exists
- ⬜ Complete screen-reader grid semantics
- ⬜ Full keyboard parity with Excel
- ⬜ High-contrast mode tuning
- ⬜ Enterprise identity / SSO
- ⬜ Audit logs
- ⬜ Retention / compliance controls
- ⬜ Admin policies

## Build priority

Recommended order for closing the largest practical gaps:

1. Virtualized, much larger worksheets
2. Cross-sheet references and more formula families
3. Merge/borders/custom formatting
4. Images, links, and worksheet objects
5. Advanced PivotTables and persistent charts
6. Cloud workbook API + accounts + share links
7. Real-time collaboration and cloud history/comments
8. Data import/query system
9. Automation / scripts
10. Deeper XLSX round-trip compatibility
