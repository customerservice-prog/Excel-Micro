import type { CellData, Point, Selection, SheetData } from './types'

export const ROWS = 200
export const COLS = 52

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

export function createBlankSheet(index = 1): SheetData {
  return {
    id: crypto.randomUUID(),
    name: `Sheet${index}`,
    cells: {},
  }
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
