export type Align = 'left' | 'center' | 'right'
export type NumberFormat = 'general' | 'number' | 'currency' | 'percent'

export interface CellFormat {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strikethrough?: boolean
  align?: Align
  color?: string
  background?: string
  numberFormat?: NumberFormat
  fontSize?: number
  wrap?: boolean
  decimals?: number
}

export interface CellData {
  value: string
  format?: CellFormat
}

export interface SheetData {
  id: string
  name: string
  cells: Record<string, CellData>
  columnWidths?: Record<string, number>
  rowHeights?: Record<string, number>
  showGridlines?: boolean
  freezeTopRow?: boolean
  freezeFirstColumn?: boolean
}

export interface WorkbookData {
  id: string
  title: string
  activeSheetId: string
  sheets: SheetData[]
  updatedAt: number
}

export interface Point {
  row: number
  col: number
}

export interface Selection {
  anchor: Point
  focus: Point
}
