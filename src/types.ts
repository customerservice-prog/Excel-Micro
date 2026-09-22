export type Align = 'left' | 'center' | 'right'
export type NumberFormat = 'general' | 'number' | 'currency' | 'accounting' | 'percent' | 'date' | 'time' | 'datetime' | 'scientific' | 'fraction'

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
  borderTop?: boolean
  borderRight?: boolean
  borderBottom?: boolean
  borderLeft?: boolean
  borderColor?: string
}

export interface CellNote {
  text: string
  updatedAt: number
}

export interface CellData {
  value: string
  format?: CellFormat
  hyperlink?: string
}

export type FilterOperator = 'contains' | 'equals' | 'greaterThan' | 'lessThan' | 'notBlank'
export interface SheetFilterRule {
  operator: FilterOperator
  value?: string
}
export interface SheetFilterRange {
  top: number
  bottom: number
  left: number
  right: number
}
export type ConditionalFormatOperator = 'greaterThan' | 'lessThan' | 'equals' | 'contains' | 'notBlank'
export interface ConditionalFormatRule {
  id: string
  top: number
  bottom: number
  left: number
  right: number
  operator: ConditionalFormatOperator
  value?: string
  background: string
  color: string
}

export type TableStyle = 'green' | 'blue' | 'orange' | 'gray'
export interface SheetTable {
  id: string
  name: string
  top: number
  bottom: number
  left: number
  right: number
  style: TableStyle
  bandedRows: boolean
}

export interface DataValidationRule {
  id: string
  top: number
  bottom: number
  left: number
  right: number
  options: string[]
  allowBlank?: boolean
}

export type PageOrientation = 'portrait' | 'landscape'
export type PagePaperSize = 'letter' | 'a4'
export type PageMargins = 'normal' | 'narrow' | 'wide'
export interface PageLayoutSettings {
  orientation: PageOrientation
  paperSize: PagePaperSize
  margins: PageMargins
  printGridlines: boolean
}

export interface MergeRange {
  id: string
  top: number
  bottom: number
  left: number
  right: number
}

export interface SheetData {
  id: string
  name: string
  cells: Record<string, CellData>
  hidden?: boolean
  columnWidths?: Record<string, number>
  rowHeights?: Record<string, number>
  showGridlines?: boolean
  freezeTopRow?: boolean
  freezeFirstColumn?: boolean
  filterRange?: SheetFilterRange
  filters?: Record<string, SheetFilterRule>
  conditionalFormats?: ConditionalFormatRule[]
  namedRanges?: Record<string, SheetFilterRange>
  dataValidations?: DataValidationRule[]
  tables?: SheetTable[]
  notes?: Record<string, CellNote>
  protected?: boolean
  pageLayout?: PageLayoutSettings
  merges?: MergeRange[]
  hiddenRows?: Record<string, boolean>
  hiddenColumns?: Record<string, boolean>
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
