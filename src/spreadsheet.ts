import type { CellData, Point, Selection, SheetData } from './types'

export const ROWS = 200
export const COLS = 52
export const DEFAULT_COLUMN_WIDTH = 106
export const DEFAULT_ROW_HEIGHT = 25

export const cellKey = (row: number, col: number) => `${row}:${col}`

export function colToName(col: number): string {
  let n = col + 1
  let name = ''
  while (n > 0) {
    const remainder = (n - 1) % 26
    name = String.fromCharCode(65 + remainder) + name
    n = Math.floor((n - 1) / 26)
  }
  return name
}

export function nameToCol(name: string): number {
  let value = 0
  for (const char of name.toUpperCase()) value = value * 26 + (char.charCodeAt(0) - 64)
  return value - 1
}

export function pointToAddress(point: Point): string {
  return `${colToName(point.col)}${point.row + 1}`
}

export function parseAddress(address: string): Point | null {
  const match = /^([A-Z]+)(\d+)$/i.exec(address.trim())
  if (!match) return null
  const row = Number(match[2]) - 1
  const col = nameToCol(match[1])
  if (row < 0 || col < 0 || row >= ROWS || col >= COLS) return null
  return { row, col }
}

export function normalizeSelection(selection: Selection) {
  return {
    top: Math.min(selection.anchor.row, selection.focus.row),
    bottom: Math.max(selection.anchor.row, selection.focus.row),
    left: Math.min(selection.anchor.col, selection.focus.col),
    right: Math.max(selection.anchor.col, selection.focus.col),
  }
}

export function selectionToAddress(selection: Selection): string {
  const range = normalizeSelection(selection)
  const start = pointToAddress({ row: range.top, col: range.left })
  const end = pointToAddress({ row: range.bottom, col: range.right })
  return start === end ? start : `${start}:${end}`
}

export function parseSelectionAddress(value: string): Selection | null {
  const parts = value.trim().toUpperCase().split(':').map((part) => part.trim()).filter(Boolean)
  if (!parts.length || parts.length > 2) return null
  const start = parseAddress(parts[0])
  const end = parseAddress(parts[1] || parts[0])
  if (!start || !end) return null
  return { anchor: start, focus: end }
}

export function getSelectedPoints(selection: Selection): Point[] {
  const { top, bottom, left, right } = normalizeSelection(selection)
  const points: Point[] = []
  for (let row = top; row <= bottom; row += 1) {
    for (let col = left; col <= right; col += 1) points.push({ row, col })
  }
  return points
}

export function getCell(sheet: SheetData, row: number, col: number): CellData {
  return sheet.cells[cellKey(row, col)] ?? { value: '' }
}

export function getColumnWidth(sheet: SheetData, col: number): number {
  return sheet.columnWidths?.[String(col)] ?? DEFAULT_COLUMN_WIDTH
}

export function getRowHeight(sheet: SheetData, row: number): number {
  return sheet.rowHeights?.[String(row)] ?? DEFAULT_ROW_HEIGHT
}

export function shiftFormulaForStructure(
  value: string,
  axis: 'row' | 'col',
  index: number,
  delta: 1 | -1,
): string {
  if (!value.startsWith('=')) return value

  return value.replace(/(\$?)([A-Z]+)(\$?)(\d+)/gi, (_match, absoluteCol: string, letters: string, absoluteRow: string, digits: string) => {
    let col = nameToCol(letters)
    let row = Number(digits) - 1

    if (axis === 'row') {
      if (delta === 1 && row >= index) row += 1
      else if (delta === -1 && row === index) return '#REF!'
      else if (delta === -1 && row > index) row -= 1
    } else {
      if (delta === 1 && col >= index) col += 1
      else if (delta === -1 && col === index) return '#REF!'
      else if (delta === -1 && col > index) col -= 1
    }

    if (row < 0 || row >= ROWS || col < 0 || col >= COLS) return '#REF!'
    return `${absoluteCol}${colToName(col)}${absoluteRow}${row + 1}`
  })
}

export function createBlankSheet(index = 1): SheetData {
  return {
    id: crypto.randomUUID(),
    name: `Sheet${index}`,
    cells: {},
    hidden: false,
    columnWidths: {},
    rowHeights: {},
    showGridlines: true,
    freezeTopRow: false,
    freezeFirstColumn: false,
    filters: {},
    conditionalFormats: [],
    namedRanges: {},
    dataValidations: [],
    tables: [],
    notes: {},
    protected: false,
    pageLayout: {
      orientation: 'portrait',
      paperSize: 'letter',
      margins: 'normal',
      printGridlines: true,
      scalePercent: 100,
      headerText: '',
      footerText: '',
    },
    merges: [],
    hiddenRows: {},
    hiddenColumns: {},
    objects: [],
  }
}

export function shiftFormulaReferences(value: string, rowDelta: number, colDelta: number): string {
  if (!value.startsWith('=')) return value

  return value.replace(/(\$?)([A-Z]+)(\$?)(\d+)/gi, (_match, absoluteCol: string, letters: string, absoluteRow: string, digits: string) => {
    const sourceCol = nameToCol(letters)
    const sourceRow = Number(digits) - 1
    const nextCol = absoluteCol ? sourceCol : Math.max(0, Math.min(COLS - 1, sourceCol + colDelta))
    const nextRow = absoluteRow ? sourceRow : Math.max(0, Math.min(ROWS - 1, sourceRow + rowDelta))
    return `${absoluteCol}${colToName(nextCol)}${absoluteRow}${nextRow + 1}`
  })
}

export function parseCsv(text: string): string[][] {
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let quoted = false

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i]
    if (quoted) {
      if (char === '"' && text[i + 1] === '"') {
        field += '"'
        i += 1
      } else if (char === '"') quoted = false
      else field += char
    } else if (char === '"') quoted = true
    else if (char === ',') {
      row.push(field)
      field = ''
    } else if (char === '\n') {
      row.push(field.replace(/\r$/, ''))
      rows.push(row)
      row = []
      field = ''
    } else field += char
  }

  row.push(field.replace(/\r$/, ''))
  if (row.some(Boolean) || rows.length === 0) rows.push(row)
  return rows
}

export function toCsv(sheet: SheetData): string {
  let maxRow = 0
  let maxCol = 0

  Object.keys(sheet.cells).forEach((key) => {
    const [row, col] = key.split(':').map(Number)
    maxRow = Math.max(maxRow, row)
    maxCol = Math.max(maxCol, col)
  })

  const quote = (value: string) => (/[,"\n\r]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value)
  const lines: string[] = []

  for (let row = 0; row <= maxRow; row += 1) {
    const values: string[] = []
    for (let col = 0; col <= maxCol; col += 1) values.push(quote(getCell(sheet, row, col).value))
    lines.push(values.join(','))
  }

  return lines.join('\n')
}
