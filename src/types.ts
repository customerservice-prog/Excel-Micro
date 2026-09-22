export type Align = 'left' | 'center' | 'right'
export type NumberFormat = 'general' | 'number' | 'currency' | 'percent'

export interface CellFormat {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  align?: Align
  color?: string
  background?: string
  numberFormat?: NumberFormat
}

export interface CellData {
  value: string
  format?: CellFormat
}

export interface SheetData {
  id: string
  name: string
  cells: Record<string, CellData>
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
