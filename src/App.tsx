import { useEffect, useMemo, useRef, useState, type CSSProperties, type ReactNode } from 'react'
import * as XLSX from 'xlsx'
import { displayValue } from './formulas'
import {
  COLS,
  DEFAULT_COLUMN_WIDTH,
  DEFAULT_ROW_HEIGHT,
  ROWS,
  cellKey,
  colToName,
  createBlankSheet,
  getCell,
  getColumnWidth,
  getRowHeight,
  getSelectedPoints,
  normalizeSelection,
  parseCsv,
  parseSelectionAddress,
  pointToAddress,
  selectionToAddress,
  shiftFormulaReferences,
  shiftFormulaForStructure,
  toCsv,
} from './spreadsheet'
import type {
  CellData,
  CellFormat,
  ConditionalFormatKind,
  ConditionalFormatOperator,
  DataValidationErrorStyle,
  DataValidationOperator,
  DataValidationRule,
  DataValidationType,
  FilterOperator,
  NumberFormat,
  Point,
  Selection,
  SheetData,
  TableStyle,
  WorkbookData,
} from './types'

const STORAGE_KEY = 'excel-micro-workbook-v1'
const VERSION_STORAGE_KEY = 'excel-micro-version-history-v1'
type Tab = 'Home' | 'Insert' | 'Formulas' | 'Data' | 'Page Layout' | 'Review' | 'View'
type ContextMenuState = { x: number; y: number; kind: 'cell' | 'row' | 'col'; index: number }
type FilterEditorState = { col: number; operator: FilterOperator; value: string }
type ChartType = 'bar' | 'line' | 'pie'
type VersionSnapshot = {
  id: string
  savedAt: number
  label: string
  workbook: WorkbookData
}

function newWorkbook(): WorkbookData {
  const sheet = createBlankSheet(1)
  return {
    id: crypto.randomUUID(),
    title: 'Book 1',
    activeSheetId: sheet.id,
    sheets: [sheet],
    updatedAt: Date.now(),
  }
}

function loadWorkbook(): WorkbookData {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    if (!raw) return newWorkbook()
    const parsed = JSON.parse(raw) as WorkbookData
    return parsed.sheets?.length ? parsed : newWorkbook()
  } catch {
    return newWorkbook()
  }
}

function loadVersions(): VersionSnapshot[] {
  try {
    const raw = localStorage.getItem(VERSION_STORAGE_KEY)
    if (!raw) return []
    const parsed = JSON.parse(raw) as VersionSnapshot[]
    return Array.isArray(parsed) ? parsed.slice(0, 12) : []
  } catch {
    return []
  }
}

function safeName(value: string) {
  return (value.trim() || 'Excel-Micro').replace(/[\\/:*?"<>|]/g, '-')
}

function download(name: string, content: BlobPart, type: string) {
  const url = URL.createObjectURL(new Blob([content], { type }))
  const a = document.createElement('a')
  a.href = url
  a.download = name
  a.click()
  URL.revokeObjectURL(url)
}

function formatted(cell: CellData, sheet: SheetData, sheets: SheetData[] = [sheet]) {
  const value = displayValue(cell.value, sheet, sheets)
  const n = Number(value)
  const format = cell.format?.numberFormat || 'general'
  const decimals = cell.format?.decimals

  if (!Number.isFinite(n) || value === '') return value

  if (format === 'currency' || format === 'accounting') {
    if (format === 'accounting' && n === 0) return '–'
    return n.toLocaleString(undefined, {
      style: 'currency',
      currency: 'USD',
      minimumFractionDigits: decimals ?? 2,
      maximumFractionDigits: decimals ?? 2,
    })
  }

  if (format === 'percent') {
    return n.toLocaleString(undefined, {
      style: 'percent',
      minimumFractionDigits: decimals ?? 0,
      maximumFractionDigits: decimals ?? 2,
    })
  }

  if (format === 'date' || format === 'time' || format === 'datetime') {
    const date = new Date(Date.UTC(1899, 11, 30) + n * 86_400_000)
    if (Number.isNaN(date.getTime())) return value
    if (format === 'date') return date.toLocaleDateString()
    if (format === 'time') {
      return date.toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit', second: '2-digit' })
    }
    return date.toLocaleString()
  }

  if (format === 'scientific') return n.toExponential(decimals ?? 2)

  if (format === 'fraction') {
    const sign = n < 0 ? '-' : ''
    const absolute = Math.abs(n)
    const whole = Math.floor(absolute)
    const fraction = absolute - whole
    if (fraction < 1e-10) return sign + whole
    let bestNumerator = 0
    let bestDenominator = 1
    let bestError = Number.POSITIVE_INFINITY
    for (let denominator = 1; denominator <= 64; denominator += 1) {
      const numerator = Math.round(fraction * denominator)
      const error = Math.abs(fraction - numerator / denominator)
      if (error < bestError) {
        bestNumerator = numerator
        bestDenominator = denominator
        bestError = error
      }
    }
    const prefix = whole ? whole + ' ' : ''
    return sign + prefix + bestNumerator + '/' + bestDenominator
  }

  if (format === 'number') {
    return n.toLocaleString(undefined, {
      minimumFractionDigits: decimals ?? 0,
      maximumFractionDigits: decimals ?? 4,
    })
  }

  return value
}

function matchesCondition(
  text: string,
  operator: FilterOperator | ConditionalFormatOperator,
  expected = '',
) {
  const source = text.trim()
  const sourceLower = source.toLowerCase()
  const expectedLower = expected.trim().toLowerCase()
  const sourceNumber = Number(source.replace(/[$,%]/g, ''))
  const expectedNumber = Number(expected.replace(/[$,%]/g, ''))

  if (operator === 'notBlank') return source !== ''
  if (operator === 'contains') return sourceLower.includes(expectedLower)
  if (operator === 'equals') {
    if (Number.isFinite(sourceNumber) && Number.isFinite(expectedNumber) && expected.trim() !== '') {
      return sourceNumber === expectedNumber
    }
    return sourceLower === expectedLower
  }
  if (operator === 'greaterThan') {
    return Number.isFinite(sourceNumber) && Number.isFinite(expectedNumber) && sourceNumber > expectedNumber
  }
  if (operator === 'lessThan') {
    return Number.isFinite(sourceNumber) && Number.isFinite(expectedNumber) && sourceNumber < expectedNumber
  }
  return true
}

type ConditionalMetric = {
  min?: number
  max?: number
  counts?: Map<string, number>
}

function hexToRgb(hex: string) {
  const normalized = hex.replace('#', '')
  const source = normalized.length === 3
    ? normalized.split('').map((char) => char + char).join('')
    : normalized.padEnd(6, '0').slice(0, 6)
  return {
    r: parseInt(source.slice(0, 2), 16) || 0,
    g: parseInt(source.slice(2, 4), 16) || 0,
    b: parseInt(source.slice(4, 6), 16) || 0,
  }
}

function mixHexColors(from: string, to: string, ratio: number) {
  const a = hexToRgb(from)
  const b = hexToRgb(to)
  const t = Math.max(0, Math.min(1, ratio))
  const channel = (start: number, end: number) => Math.round(start + (end - start) * t)
  return `rgb(${channel(a.r, b.r)}, ${channel(a.g, b.g)}, ${channel(a.b, b.b)})`
}

function conditionalStyleForCell(
  sheet: SheetData,
  row: number,
  col: number,
  value: string,
  metrics: Record<string, ConditionalMetric> = {},
) {
  let background: string | undefined
  let color: string | undefined
  let dataBarPercent: number | undefined
  let dataBarColor: string | undefined

  for (const rule of sheet.conditionalFormats || []) {
    if (row < rule.top || row > rule.bottom || col < rule.left || col > rule.right) continue
    const kind = rule.kind || 'cell'
    const metric = metrics[rule.id]

    if (kind === 'dataBar') {
      const numeric = Number(value.replace(/[$,%]/g, ''))
      if (!Number.isFinite(numeric)) continue
      const min = metric?.min ?? 0
      const max = metric?.max ?? 0
      dataBarPercent = max === min ? 100 : ((numeric - min) / (max - min)) * 100
      dataBarColor = rule.dataBarColor || '#63be7b'
      continue
    }

    if (kind === 'colorScale') {
      const numeric = Number(value.replace(/[$,%]/g, ''))
      if (!Number.isFinite(numeric)) continue
      const min = metric?.min ?? 0
      const max = metric?.max ?? 0
      const ratio = max === min ? 0.5 : (numeric - min) / (max - min)
      background = mixHexColors(rule.minColor || '#f8696b', rule.maxColor || '#63be7b', ratio)
      continue
    }

    if (kind === 'duplicate' || kind === 'unique') {
      const count = metric?.counts?.get(value.trim().toLowerCase()) || 0
      const matches = kind === 'duplicate' ? count > 1 : count === 1
      if (!matches) continue
      background = rule.background
      color = rule.color
      continue
    }

    if (!matchesCondition(value, rule.operator, rule.value || '')) continue
    background = rule.background
    color = rule.color
  }

  return { background, color, dataBarPercent, dataBarColor }
}

const TABLE_PALETTES: Record<TableStyle, { header: string; band: string; text: string }> = {
  green: { header: '#107c41', band: '#eaf4ee', text: '#ffffff' },
  blue: { header: '#2b579a', band: '#eaf0f8', text: '#ffffff' },
  orange: { header: '#c65911', band: '#fbefe6', text: '#ffffff' },
  gray: { header: '#5b5b5b', band: '#f1f1f1', text: '#ffffff' },
}

function tableStyleForCell(sheet: SheetData, row: number, col: number) {
  const tables = sheet.tables || []

  for (let index = tables.length - 1; index >= 0; index -= 1) {
    const table = tables[index]
    if (row < table.top || row > table.bottom || col < table.left || col > table.right) continue

    const palette = TABLE_PALETTES[table.style]
    if (row === table.top) {
      return {
        background: palette.header,
        color: palette.text,
        bold: true,
        header: true,
      }
    }

    const banded = table.bandedRows && (row - table.top) % 2 === 0
    return {
      background: banded ? palette.band : undefined,
      color: undefined,
      bold: false,
      header: false,
    }
  }

  return {
    background: undefined,
    color: undefined,
    bold: false,
    header: false,
  }
}

function mergeForCell(sheet: SheetData, row: number, col: number) {
  return (sheet.merges || []).find((merge) => (
    row >= merge.top &&
    row <= merge.bottom &&
    col >= merge.left &&
    col <= merge.right
  ))
}

function dataValidationForCell(sheet: SheetData, row: number, col: number) {
  const rules = sheet.dataValidations || []
  for (let index = rules.length - 1; index >= 0; index -= 1) {
    const rule = rules[index]
    if (row >= rule.top && row <= rule.bottom && col >= rule.left && col <= rule.right) {
      return rule
    }
  }
  return undefined
}

function validationAllowsValue(rule: DataValidationRule | undefined, value: string) {
  if (!rule) return true
  const trimmed = value.trim()
  if (rule.allowBlank && trimmed === '') return true

  const type = rule.type || 'list'
  if (type === 'list') {
    return (rule.options || []).some((option) => option === value)
  }

  const operator = rule.operator || 'between'
  let candidate: number
  let minimum: number
  let maximum: number

  if (type === 'date') {
    const toDateNumber = (source: string) => {
      if (source.trim() === '') return Number.NaN
      if (Number.isFinite(Number(source))) {
        return Date.UTC(1899, 11, 30) + Number(source) * 86_400_000
      }
      return Date.parse(source)
    }
    candidate = toDateNumber(trimmed)
    minimum = toDateNumber(rule.minimum || '')
    maximum = toDateNumber(rule.maximum || '')
  } else if (type === 'textLength') {
    candidate = value.length
    minimum = Number(rule.minimum)
    maximum = Number(rule.maximum)
  } else {
    candidate = Number(trimmed)
    minimum = Number(rule.minimum)
    maximum = Number(rule.maximum)
    if (!Number.isFinite(candidate)) return false
    if (type === 'whole' && !Number.isInteger(candidate)) return false
  }

  if (!Number.isFinite(candidate)) return false
  if (operator === 'equalTo') return candidate === minimum
  if (operator === 'greaterThan') return candidate > minimum
  if (operator === 'lessThan') return candidate < minimum
  return candidate >= minimum && candidate <= maximum
}

function validationDescription(rule: DataValidationRule) {
  const type = rule.type || 'list'
  if (type === 'list') return 'Choose one of: ' + (rule.options || []).join(', ')
  const operator = rule.operator || 'between'
  const label = type === 'whole'
    ? 'whole number'
    : type === 'decimal'
      ? 'number'
      : type === 'date'
        ? 'date'
        : 'text length'

  if (operator === 'between') return 'Enter a ' + label + ' between ' + (rule.minimum || '') + ' and ' + (rule.maximum || '')
  if (operator === 'equalTo') return 'Enter a ' + label + ' equal to ' + (rule.minimum || '')
  if (operator === 'greaterThan') return 'Enter a ' + label + ' greater than ' + (rule.minimum || '')
  return 'Enter a ' + label + ' less than ' + (rule.minimum || '')
}

type TrackMetrics = {
  sizes: number[]
  offsets: number[]
  total: number
}

function buildTrackMetrics(sizes: number[], leading: number): TrackMetrics {
  const offsets: number[] = []
  let cursor = leading
  sizes.forEach((size, index) => {
    offsets[index] = cursor
    cursor += size
  })
  return { sizes, offsets, total: cursor }
}

function visibleTrackIndexes(
  metrics: TrackMetrics,
  viewportStart: number,
  viewportSize: number,
  overscan: number,
) {
  const { offsets, sizes } = metrics
  if (!sizes.length) return []

  const targetStart = Math.max(0, viewportStart)
  const targetEnd = targetStart + Math.max(1, viewportSize)

  let low = 0
  let high = sizes.length - 1
  let first = sizes.length - 1
  while (low <= high) {
    const middle = Math.floor((low + high) / 2)
    if (offsets[middle] + sizes[middle] >= targetStart) {
      first = middle
      high = middle - 1
    } else {
      low = middle + 1
    }
  }

  low = first
  high = sizes.length - 1
  let last = first
  while (low <= high) {
    const middle = Math.floor((low + high) / 2)
    if (offsets[middle] <= targetEnd) {
      last = middle
      low = middle + 1
    } else {
      high = middle - 1
    }
  }

  const start = Math.max(0, first - overscan)
  const end = Math.min(sizes.length - 1, last + overscan)
  const result: number[] = []
  for (let index = start; index <= end; index += 1) {
    if (sizes[index] > 0) result.push(index)
  }
  return result
}

function chartDataForRange(
  sheet: SheetData,
  selected: { top: number; bottom: number; left: number; right: number },
  sheets: SheetData[],
) {
  const data: { label: string; value: number }[] = []

  for (let row = selected.top; row <= selected.bottom; row += 1) {
    const label = selected.right > selected.left
      ? formatted(getCell(sheet, row, selected.left), sheet, sheets)
      : String(row + 1)
    const valueCol = selected.right > selected.left ? selected.left + 1 : selected.left
    const value = Number(displayValue(getCell(sheet, row, valueCol).value, sheet, sheets))
    if (Number.isFinite(value)) data.push({ label: label || String(row + 1), value })
  }

  return data.slice(0, 100)
}

function uniqueSheetName(sheets: SheetData[], base: string) {
  const existing = new Set(sheets.map((sheet) => sheet.name.toLowerCase()))
  let candidate = base
  let index = 2

  while (existing.has(candidate.toLowerCase())) {
    candidate = `${base} (${index})`
    index += 1
  }

  return candidate.slice(0, 31)
}

export default function App() {
  const [book, setBook] = useState<WorkbookData>(loadWorkbook)
  const [selection, setSelection] = useState<Selection>({
    anchor: { row: 0, col: 0 },
    focus: { row: 0, col: 0 },
  })
  const [tab, setTab] = useState<Tab>('Home')
  const [editing, setEditing] = useState(false)
  const [draft, setDraft] = useState('')
  const [dragging, setDragging] = useState(false)
  const [history, setHistory] = useState<WorkbookData[]>([])
  const [future, setFuture] = useState<WorkbookData[]>([])
  const [versions, setVersions] = useState<VersionSnapshot[]>(loadVersions)
  const [versionOpen, setVersionOpen] = useState(false)
  const [find, setFind] = useState('')
  const [replaceText, setReplaceText] = useState('')
  const [formatPainter, setFormatPainter] = useState<CellFormat | null>(null)
  const [findOpen, setFindOpen] = useState(false)
  const [chartOpen, setChartOpen] = useState(false)
  const [chartType, setChartType] = useState<ChartType>('bar')
  const [zoom, setZoom] = useState(100)
  const [nameBox, setNameBox] = useState('A1')
  const [saveStatus, setSaveStatus] = useState('Saved')
  const [notice, setNotice] = useState('')
  const [contextMenu, setContextMenu] = useState<ContextMenuState | null>(null)
  const [filterEditor, setFilterEditor] = useState<FilterEditorState | null>(null)
  const [conditionalOpen, setConditionalOpen] = useState(false)
  const [conditionalOperator, setConditionalOperator] = useState<ConditionalFormatOperator>('greaterThan')
  const [conditionalValue, setConditionalValue] = useState('')
  const [conditionalBackground, setConditionalBackground] = useState('#fff2cc')
  const [conditionalColor, setConditionalColor] = useState('#7f6000')
  const [namedRangeOpen, setNamedRangeOpen] = useState(false)
  const [namedRangeName, setNamedRangeName] = useState('')
  const [validationOpen, setValidationOpen] = useState(false)
  const [validationOptions, setValidationOptions] = useState('Pending\nPaid\nCanceled')
  const [validationAllowBlank, setValidationAllowBlank] = useState(true)
  const [validationType, setValidationType] = useState<DataValidationType>('list')
  const [validationOperator, setValidationOperator] = useState<DataValidationOperator>('between')
  const [validationMinimum, setValidationMinimum] = useState('0')
  const [validationMaximum, setValidationMaximum] = useState('100')
  const [validationInputTitle, setValidationInputTitle] = useState('')
  const [validationInputMessage, setValidationInputMessage] = useState('')
  const [validationErrorStyle, setValidationErrorStyle] = useState<DataValidationErrorStyle>('stop')
  const [validationPicker, setValidationPicker] = useState<{
    row: number
    col: number
    x: number
    y: number
    ruleId: string
  } | null>(null)
  const [tableOpen, setTableOpen] = useState(false)
  const [tableName, setTableName] = useState('Table1')
  const [tableStyle, setTableStyle] = useState<TableStyle>('green')
  const [tableBandedRows, setTableBandedRows] = useState(true)
  const [noteOpen, setNoteOpen] = useState(false)
  const [noteDraft, setNoteDraft] = useState('')
  const [noteTarget, setNoteTarget] = useState<Point>({ row: 0, col: 0 })
  const [pivotOpen, setPivotOpen] = useState(false)
  const [pivotSource, setPivotSource] = useState({ top: 0, bottom: 1, left: 0, right: 1 })
  const [pivotRowField, setPivotRowField] = useState(0)
  const [pivotValueField, setPivotValueField] = useState(1)
  const [pivotAggregator, setPivotAggregator] = useState<'sum' | 'count' | 'average'>('sum')
  const [hyperlinkOpen, setHyperlinkOpen] = useState(false)
  const [hyperlinkDraft, setHyperlinkDraft] = useState('')
  const [hyperlinkTarget, setHyperlinkTarget] = useState<Point>({ row: 0, col: 0 })
  const [viewport, setViewport] = useState({
    scrollTop: 0,
    scrollLeft: 0,
    width: 1200,
    height: 700,
  })
  const fileRef = useRef<HTMLInputElement>(null)
  const imageRef = useRef<HTMLInputElement>(null)
  const gridRef = useRef<HTMLDivElement>(null)
  const columnResizeRef = useRef<{ col: number; startX: number; startWidth: number } | null>(null)
  const rowResizeRef = useRef<{ row: number; startY: number; startHeight: number } | null>(null)

  const sheet = useMemo(
    () => book.sheets.find((item) => item.id === book.activeSheetId) || book.sheets[0],
    [book],
  )
  const point = selection.focus
  const cell = getCell(sheet, point.row, point.col)
  const range = normalizeSelection(selection)

  const hiddenRows = useMemo(() => {
    const hidden = new Set<number>()
    const filterRange = sheet.filterRange
    const filters = sheet.filters || {}
    if (!filterRange || !Object.keys(filters).length) return hidden

    for (let row = filterRange.top + 1; row <= filterRange.bottom; row += 1) {
      const visible = Object.entries(filters).every(([colText, rule]) => {
        const col = Number(colText)
        const value = formatted(getCell(sheet, row, col), sheet, book.sheets)
        return matchesCondition(value, rule.operator, rule.value || '')
      })

      if (!visible) hidden.add(row)
    }

    return hidden
  }, [sheet])

  const conditionalMetrics = useMemo(() => {
    const result: Record<string, ConditionalMetric> = {}

    for (const rule of sheet.conditionalFormats || []) {
      const kind = rule.kind || 'cell'
      if (kind === 'cell') continue

      const numbers: number[] = []
      const counts = new Map<string, number>()

      for (let row = rule.top; row <= rule.bottom; row += 1) {
        for (let col = rule.left; col <= rule.right; col += 1) {
          const value = displayValue(getCell(sheet, row, col).value, sheet, book.sheets)
          const normalized = value.trim().toLowerCase()
          counts.set(normalized, (counts.get(normalized) || 0) + 1)

          const numeric = Number(value.replace(/[$,%]/g, ''))
          if (Number.isFinite(numeric)) numbers.push(numeric)
        }
      }

      result[rule.id] = {
        min: numbers.length ? Math.min(...numbers) : 0,
        max: numbers.length ? Math.max(...numbers) : 0,
        counts,
      }
    }

    return result
  }, [book.sheets, sheet])

  useEffect(() => {
    setSaveStatus('Saving…')
    const timer = window.setTimeout(() => {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(book))
      setSaveStatus('Saved')
    }, 220)

    return () => window.clearTimeout(timer)
  }, [book])

  useEffect(() => {
    setNameBox(selectionToAddress(selection))
  }, [selection])

  useEffect(() => {
    try {
      localStorage.setItem(VERSION_STORAGE_KEY, JSON.stringify(versions.slice(0, 12)))
    } catch {
      window.setTimeout(() => setNotice('Version history storage is full in this browser.'), 0)
    }
  }, [versions])

  useEffect(() => {
    const stop = () => setDragging(false)
    window.addEventListener('mouseup', stop)
    return () => window.removeEventListener('mouseup', stop)
  }, [])

  useEffect(() => {
    const node = gridRef.current
    if (!node) return

    const update = () => {
      setViewport((current) => ({
        ...current,
        scrollTop: node.scrollTop,
        scrollLeft: node.scrollLeft,
        width: node.clientWidth || current.width,
        height: node.clientHeight || current.height,
      }))
    }

    update()
    const observer = new ResizeObserver(update)
    observer.observe(node)
    window.addEventListener('resize', update)

    return () => {
      observer.disconnect()
      window.removeEventListener('resize', update)
    }
  }, [])

  useEffect(() => {
    const onMove = (event: MouseEvent) => {
      const columnResize = columnResizeRef.current
      const rowResize = rowResizeRef.current

      if (columnResize) {
        const width = Math.max(
          48,
          Math.min(360, columnResize.startWidth + event.clientX - columnResize.startX),
        )

        setBook((current) => {
          const next = structuredClone(current)
          const target = next.sheets.find((item) => item.id === next.activeSheetId) || next.sheets[0]
          target.columnWidths ||= {}
          target.columnWidths[String(columnResize.col)] = Math.round(width)
          next.updatedAt = Date.now()
          return next
        })
      }

      if (rowResize) {
        const height = Math.max(
          20,
          Math.min(180, rowResize.startHeight + event.clientY - rowResize.startY),
        )

        setBook((current) => {
          const next = structuredClone(current)
          const target = next.sheets.find((item) => item.id === next.activeSheetId) || next.sheets[0]
          target.rowHeights ||= {}
          target.rowHeights[String(rowResize.row)] = Math.round(height)
          next.updatedAt = Date.now()
          return next
        })
      }
    }

    const onUp = () => {
      if (columnResizeRef.current) setNotice('Column width updated')
      if (rowResizeRef.current) setNotice('Row height updated')
      columnResizeRef.current = null
      rowResizeRef.current = null
      document.body.classList.remove('is-column-resizing')
      document.body.classList.remove('is-row-resizing')
    }

    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)

    return () => {
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
    }
  }, [])

  useEffect(() => {
    if (!notice) return
    const timer = window.setTimeout(() => setNotice(''), 2200)
    return () => window.clearTimeout(timer)
  }, [notice])

  useEffect(() => {
    const close = () => {
      setContextMenu(null)
      setValidationPicker(null)
    }
    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setContextMenu(null)
    }

    window.addEventListener('mousedown', close)
    window.addEventListener('resize', close)
    window.addEventListener('blur', close)
    window.addEventListener('keydown', closeOnEscape)

    return () => {
      window.removeEventListener('mousedown', close)
      window.removeEventListener('resize', close)
      window.removeEventListener('blur', close)
      window.removeEventListener('keydown', closeOnEscape)
    }
  }, [])

  function activeSheet(next: WorkbookData) {
    return next.sheets.find((item) => item.id === next.activeSheetId) || next.sheets[0]
  }

  function mutate(change: (next: WorkbookData) => void, allowProtected = false) {
    if (sheet.protected && !allowProtected) {
      window.setTimeout(() => setNotice('This sheet is protected. Unprotect it from Review to edit.'), 0)
      return
    }

    setHistory((items) => [...items, book].slice(-75))
    setFuture([])
    const next = structuredClone(book)
    change(next)
    next.updatedAt = Date.now()
    setBook(next)
  }

  function setValue(target: Point, value: string) {
    mutate((next) => {
      const s = activeSheet(next)
      const key = cellKey(target.row, target.col)
      const old = s.cells[key] || { value: '' }

      if (!value && !old.format) delete s.cells[key]
      else s.cells[key] = { ...old, value }
    })
  }

  function beginEdit(seed?: string) {
    if (sheet.protected) {
      setNotice('This sheet is protected. Unprotect it from Review to edit.')
      return
    }
    setDraft(seed === undefined ? cell.value : seed)
    setEditing(true)
  }

  function endEdit(save = true) {
    if (save) {
      const validation = dataValidationForCell(sheet, point.row, point.col)
      const valid = validationAllowsValue(validation, draft)

      if (!valid && validation) {
        const description = validationDescription(validation)
        const style = validation.errorStyle || 'stop'

        if (style === 'stop') {
          setNotice(description)
          setEditing(false)
          setDraft('')
          window.requestAnimationFrame(() => gridRef.current?.focus())
          return
        }

        if (style === 'warning' && !window.confirm(description + '. Keep this value anyway?')) {
          setEditing(false)
          setDraft('')
          window.requestAnimationFrame(() => gridRef.current?.focus())
          return
        }

        if (style === 'information') setNotice(description)
      }

      setValue(point, draft)
    }

    setEditing(false)
    setDraft('')
    window.requestAnimationFrame(() => gridRef.current?.focus())
  }

  function move(dr: number, dc: number, extend = false) {
    const next = {
      row: Math.max(0, Math.min(ROWS - 1, point.row + dr)),
      col: Math.max(0, Math.min(COLS - 1, point.col + dc)),
    }

    setSelection((current) => ({
      anchor: extend ? current.anchor : next,
      focus: next,
    }))
  }

  function selectPoint(target: Point) {
    setSelection({ anchor: target, focus: target })
  }

  function selectedPointsForMutation(limit = 100_000) {
    const selected = normalizeSelection(selection)
    const area = (selected.bottom - selected.top + 1) * (selected.right - selected.left + 1)
    if (area > limit) {
      setNotice(`That selection contains ${area.toLocaleString()} cells. Narrow the range for this operation.`)
      return null
    }
    return getSelectedPoints(selection)
  }

  function openContextMenu(
    event: React.MouseEvent,
    kind: ContextMenuState['kind'],
    index: number,
    target?: Point,
  ) {
    event.preventDefault()
    event.stopPropagation()

    if (kind === 'row') {
      setSelection({
        anchor: { row: index, col: 0 },
        focus: { row: index, col: COLS - 1 },
      })
    } else if (kind === 'col') {
      setSelection({
        anchor: { row: 0, col: index },
        focus: { row: ROWS - 1, col: index },
      })
    } else if (target) {
      const current = normalizeSelection(selection)
      const inside =
        target.row >= current.top &&
        target.row <= current.bottom &&
        target.col >= current.left &&
        target.col <= current.right

      if (!inside) selectPoint(target)
    }

    const menuWidth = 210
    const menuHeight = kind === 'cell' ? 250 : 220
    setContextMenu({
      x: Math.max(8, Math.min(window.innerWidth - menuWidth - 8, event.clientX)),
      y: Math.max(8, Math.min(window.innerHeight - menuHeight - 8, event.clientY)),
      kind,
      index,
    })
  }

  function runContextAction(action: () => void) {
    setContextMenu(null)
    window.requestAnimationFrame(action)
  }

  function applyBorders(mode: 'all' | 'outside' | 'bottom' | 'none') {
    const selected = normalizeSelection(selection)
    const color = cell.format?.borderColor || '#808080'

    mutate((next) => {
      const s = activeSheet(next)

      for (let row = selected.top; row <= selected.bottom; row += 1) {
        for (let col = selected.left; col <= selected.right; col += 1) {
          const key = cellKey(row, col)
          const old = s.cells[key] || { value: '' }
          const format = { ...(old.format || {}) }

          if (mode === 'none') {
            format.borderTop = false
            format.borderRight = false
            format.borderBottom = false
            format.borderLeft = false
          } else if (mode === 'all') {
            format.borderTop = true
            format.borderRight = true
            format.borderBottom = true
            format.borderLeft = true
          } else if (mode === 'outside') {
            format.borderTop = row === selected.top
            format.borderBottom = row === selected.bottom
            format.borderLeft = col === selected.left
            format.borderRight = col === selected.right
          } else {
            format.borderBottom = true
          }

          format.borderColor = color
          s.cells[key] = { ...old, format }
        }
      }
    })
  }

  function mergeSelection() {
    const selected = normalizeSelection(selection)
    if (selected.top === selected.bottom && selected.left === selected.right) {
      setNotice('Select two or more cells to merge')
      return
    }

    mutate((next) => {
      const s = activeSheet(next)
      s.merges ||= []
      s.merges = s.merges.filter((merge) => (
        merge.bottom < selected.top ||
        merge.top > selected.bottom ||
        merge.right < selected.left ||
        merge.left > selected.right
      ))

      const originKey = cellKey(selected.top, selected.left)
      const origin = s.cells[originKey] || { value: '' }

      for (let row = selected.top; row <= selected.bottom; row += 1) {
        for (let col = selected.left; col <= selected.right; col += 1) {
          if (row === selected.top && col === selected.left) continue
          const key = cellKey(row, col)
          const existing = s.cells[key]
          if (!existing) continue
          if (existing.format || existing.hyperlink) s.cells[key] = { ...existing, value: '' }
          else delete s.cells[key]
        }
      }

      s.cells[originKey] = origin
      s.merges.push({
        id: crypto.randomUUID(),
        top: selected.top,
        bottom: selected.bottom,
        left: selected.left,
        right: selected.right,
      })
    })

    selectPoint({ row: selected.top, col: selected.left })
    setNotice('Cells merged')
  }

  function unmergeSelection() {
    const selected = normalizeSelection(selection)
    const matches = (sheet.merges || []).filter((merge) => !(
      merge.bottom < selected.top ||
      merge.top > selected.bottom ||
      merge.right < selected.left ||
      merge.left > selected.right
    ))

    if (!matches.length) {
      setNotice('No merged cells in the selection')
      return
    }

    mutate((next) => {
      const s = activeSheet(next)
      s.merges = (s.merges || []).filter((merge) => !matches.some((match) => match.id === merge.id))
    })
    setNotice('Cells unmerged')
  }

  function openHyperlink(target: Point = point) {
    setHyperlinkTarget(target)
    setHyperlinkDraft(getCell(sheet, target.row, target.col).hyperlink || 'https://')
    setHyperlinkOpen(true)
  }

  function saveHyperlink() {
    const url = hyperlinkDraft.trim()
    mutate((next) => {
      const s = activeSheet(next)
      const key = cellKey(hyperlinkTarget.row, hyperlinkTarget.col)
      const old = s.cells[key] || { value: '' }
      s.cells[key] = { ...old, hyperlink: url || undefined }
    })
    setHyperlinkOpen(false)
    setNotice(url ? 'Hyperlink added' : 'Hyperlink removed')
  }

  function applyFormat(patch: Partial<CellFormat>) {
    const points = selectedPointsForMutation()
    if (!points) return

    mutate((next) => {
      const s = activeSheet(next)

      points.forEach((p) => {
        const key = cellKey(p.row, p.col)
        const old = s.cells[key] || { value: '' }
        s.cells[key] = {
          ...old,
          format: { ...(old.format || {}), ...patch },
        }
      })
    })
  }

  function clearSelected() {
    const points = selectedPointsForMutation()
    if (!points) return

    mutate((next) => {
      const s = activeSheet(next)

      points.forEach((p) => {
        const key = cellKey(p.row, p.col)
        const old = s.cells[key]
        if (!old) return
        if (old.format || old.hyperlink || old.checkbox) {
          s.cells[key] = { ...old, value: old.checkbox ? 'FALSE' : '' }
        } else {
          delete s.cells[key]
        }
      })
    })
  }

  function clearFormatting() {
    const points = selectedPointsForMutation()
    if (!points) return

    mutate((next) => {
      const s = activeSheet(next)

      points.forEach((p) => {
        const key = cellKey(p.row, p.col)
        const old = s.cells[key]
        if (!old) return
        if (old.value || old.hyperlink || old.checkbox) {
          s.cells[key] = {
            value: old.value,
            hyperlink: old.hyperlink,
            checkbox: old.checkbox,
          }
        } else {
          delete s.cells[key]
        }
      })
    })
  }

  function undo() {
    if (sheet.protected) {
      setNotice('Unprotect the sheet before undoing edits.')
      return
    }
    const previous = history.at(-1)
    if (!previous) return

    setFuture((items) => [book, ...items].slice(0, 75))
    setHistory((items) => items.slice(0, -1))
    setBook(previous)
  }

  function redo() {
    if (sheet.protected) {
      setNotice('Unprotect the sheet before redoing edits.')
      return
    }
    const next = future[0]
    if (!next) return

    setHistory((items) => [...items, book].slice(-75))
    setFuture((items) => items.slice(1))
    setBook(next)
  }

  async function copySelected() {
    const r = normalizeSelection(selection)
    const area = (r.bottom - r.top + 1) * (r.right - r.left + 1)
    if (area > 200_000) {
      setNotice(`Copy is limited to 200,000 cells at a time. Current selection: ${area.toLocaleString()}.`)
      return
    }
    const rows: string[] = []

    for (let row = r.top; row <= r.bottom; row += 1) {
      const values: string[] = []
      for (let col = r.left; col <= r.right; col += 1) {
        values.push(getCell(sheet, row, col).value)
      }
      rows.push(values.join('\t'))
    }

    try {
      await navigator.clipboard.writeText(rows.join('\n'))
      setNotice('Copied selection')
    } catch {
      setNotice('Clipboard permission was blocked')
    }
  }

  function pasteText(text: string) {
    const rows = text
      .replace(/\r/g, '')
      .split('\n')
      .map((row) => row.split('\t'))

    let rejected = 0
    let warned = 0

    mutate((next) => {
      const s = activeSheet(next)

      rows.forEach((values, ro) => {
        values.forEach((value, co) => {
          const row = point.row + ro
          const col = point.col + co
          if (row >= ROWS || col >= COLS) return

          const validation = dataValidationForCell(s, row, col)
          const valid = validationAllowsValue(validation, value)

          if (!valid && validation) {
            const style = validation.errorStyle || 'stop'
            if (style === 'stop') {
              rejected += 1
              return
            }
            warned += 1
          }

          const key = cellKey(row, col)
          s.cells[key] = {
            ...(s.cells[key] || { value: '' }),
            value,
          }
        })
      })
    })

    if (rejected > 0) {
      setNotice(`${rejected} pasted value${rejected === 1 ? '' : 's'} rejected by validation`)
    } else if (warned > 0) {
      setNotice(`${warned} pasted value${warned === 1 ? '' : 's'} did not meet validation rules`)
    }
  }

  function fillDown() {
    if (range.bottom <= range.top) {
      setNotice('Select two or more rows to fill down')
      return
    }

    mutate((next) => {
      const s = activeSheet(next)

      for (let col = range.left; col <= range.right; col += 1) {
        const source = structuredClone(getCell(s, range.top, col))

        for (let row = range.top + 1; row <= range.bottom; row += 1) {
          const shifted = {
            ...structuredClone(source),
            value: shiftFormulaReferences(source.value, row - range.top, 0),
          }
          s.cells[cellKey(row, col)] = shifted
        }
      }
    })

    setNotice('Filled down')
  }

  function fillRight() {
    if (range.right <= range.left) {
      setNotice('Select two or more columns to fill right')
      return
    }

    mutate((next) => {
      const s = activeSheet(next)

      for (let row = range.top; row <= range.bottom; row += 1) {
        const source = structuredClone(getCell(s, row, range.left))

        for (let col = range.left + 1; col <= range.right; col += 1) {
          const shifted = {
            ...structuredClone(source),
            value: shiftFormulaReferences(source.value, 0, col - range.left),
          }
          s.cells[cellKey(row, col)] = shifted
        }
      }
    })

    setNotice('Filled right')
  }

  useEffect(() => {
    const onPaste = (event: ClipboardEvent) => {
      if (editing) return
      const text = event.clipboardData?.getData('text/plain')
      if (!text) return
      event.preventDefault()
      pasteText(text)
    }

    window.addEventListener('paste', onPaste)
    return () => window.removeEventListener('paste', onPaste)
  })

  function onGridKey(event: React.KeyboardEvent<HTMLDivElement>) {
    if (editing) return

    const cmd = event.ctrlKey || event.metaKey
    const key = event.key.toLowerCase()

    if (cmd && key === 'z') {
      event.preventDefault()
      event.shiftKey ? redo() : undo()
      return
    }

    if (cmd && key === 'y') {
      event.preventDefault()
      redo()
      return
    }

    if (cmd && key === 'c') {
      event.preventDefault()
      void copySelected()
      return
    }

    if (cmd && key === 'f') {
      event.preventDefault()
      setFindOpen(true)
      return
    }

    if (cmd && key === 's') {
      event.preventDefault()
      exportXlsx()
      setNotice('Workbook downloaded')
      return
    }

    if (cmd && key === 'a') {
      event.preventDefault()
      setSelection({
        anchor: { row: 0, col: 0 },
        focus: { row: ROWS - 1, col: COLS - 1 },
      })
      return
    }

    if (cmd && key === 'b') {
      event.preventDefault()
      applyFormat({ bold: !cell.format?.bold })
      return
    }

    if (cmd && key === 'i') {
      event.preventDefault()
      applyFormat({ italic: !cell.format?.italic })
      return
    }

    if (cmd && key === 'u') {
      event.preventDefault()
      applyFormat({ underline: !cell.format?.underline })
      return
    }

    if (cmd && key === 'd') {
      event.preventDefault()
      fillDown()
      return
    }

    if (cmd && key === 'r') {
      event.preventDefault()
      fillRight()
      return
    }

    if (event.key === 'Delete' || event.key === 'Backspace') {
      event.preventDefault()
      clearSelected()
      return
    }

    if (event.key === 'Enter' || event.key === 'F2') {
      event.preventDefault()
      beginEdit()
      return
    }

    if (event.key === 'Tab') {
      event.preventDefault()
      move(0, event.shiftKey ? -1 : 1)
      return
    }

    const dirs: Record<string, [number, number]> = {
      ArrowUp: [-1, 0],
      ArrowDown: [1, 0],
      ArrowLeft: [0, -1],
      ArrowRight: [0, 1],
    }

    if (dirs[event.key]) {
      event.preventDefault()
      const d = dirs[event.key]
      move(d[0], d[1], event.shiftKey)
      return
    }

    if (!cmd && !event.altKey && event.key.length === 1) {
      event.preventDefault()
      beginEdit(event.key)
    }
  }

  function copyFormat() {
    setFormatPainter(structuredClone(cell.format || {}))
    setNotice('Format Painter active — click a cell to apply')
  }

  function hideCurrentSheet() {
    const visible = book.sheets.filter((item) => !item.hidden)
    if (visible.length <= 1) {
      setNotice('At least one worksheet must remain visible')
      return
    }

    const currentIndex = book.sheets.findIndex((item) => item.id === sheet.id)
    const nextVisible = book.sheets.find((item, index) => index > currentIndex && !item.hidden)
      || [...book.sheets].reverse().find((item, index) => book.sheets.length - 1 - index < currentIndex && !item.hidden)

    mutate((next) => {
      const target = next.sheets.find((item) => item.id === sheet.id)
      if (target) target.hidden = true
      if (nextVisible) next.activeSheetId = nextVisible.id
    })
    setNotice('Worksheet hidden')
  }

  function unhideAllSheets() {
    const hidden = book.sheets.filter((item) => item.hidden)
    if (!hidden.length) {
      setNotice('No hidden worksheets')
      return
    }

    mutate((next) => {
      next.sheets.forEach((item) => { item.hidden = false })
    })
    setNotice('All worksheets unhidden')
  }

  function removeDuplicates() {
    const selected = normalizeSelection(selection)
    if (selected.bottom <= selected.top) {
      setNotice('Select two or more rows to remove duplicates')
      return
    }

    let removed = 0

    mutate((next) => {
      const s = activeSheet(next)
      const uniqueRows: CellData[][] = []
      const seen = new Set<string>()

      for (let row = selected.top; row <= selected.bottom; row += 1) {
        const rowCells: CellData[] = []
        const signature: string[] = []

        for (let col = selected.left; col <= selected.right; col += 1) {
          const data = structuredClone(getCell(s, row, col))
          rowCells.push(data)
          signature.push(data.value)
        }

        const key = JSON.stringify(signature)
        if (seen.has(key)) {
          removed += 1
          continue
        }

        seen.add(key)
        uniqueRows.push(rowCells)
      }

      for (let row = selected.top; row <= selected.bottom; row += 1) {
        for (let col = selected.left; col <= selected.right; col += 1) {
          delete s.cells[cellKey(row, col)]
        }
      }

      uniqueRows.forEach((rowCells, rowOffset) => {
        rowCells.forEach((data, colOffset) => {
          if (!data.value && !data.format && !data.hyperlink) return
          s.cells[cellKey(selected.top + rowOffset, selected.left + colOffset)] = data
        })
      })
    })

    setNotice(removed ? 'Removed ' + removed + ' duplicate row' + (removed === 1 ? '' : 's') : 'No duplicates found')
  }

  function textToColumns(delimiter: ',' | '\t' | ';' = ',') {
    const selected = normalizeSelection(selection)
    if (selected.left !== selected.right) {
      setNotice('Text to Columns requires a single selected column')
      return
    }

    let widest = 1

    mutate((next) => {
      const s = activeSheet(next)

      for (let row = selected.top; row <= selected.bottom; row += 1) {
        const source = getCell(s, row, selected.left)
        const parts = source.value.split(delimiter)
        widest = Math.max(widest, parts.length)

        parts.forEach((part, offset) => {
          const col = selected.left + offset
          if (col >= COLS) return
          const key = cellKey(row, col)
          const old = s.cells[key] || { value: '' }
          s.cells[key] = {
            ...old,
            value: part.trim(),
          }
        })
      }
    })

    setSelection({
      anchor: { row: selected.top, col: selected.left },
      focus: { row: selected.bottom, col: Math.min(COLS - 1, selected.left + widest - 1) },
    })
    setNotice('Text split into ' + widest + ' column' + (widest === 1 ? '' : 's'))
  }

  function insertCheckboxes() {
    const points = selectedPointsForMutation(50_000)
    if (!points) return
    mutate((next) => {
      const s = activeSheet(next)
      points.forEach((target) => {
        const key = cellKey(target.row, target.col)
        const old = s.cells[key] || { value: '' }
        const checked = old.value.toUpperCase() === 'TRUE'
        s.cells[key] = {
          ...old,
          value: old.value === '' ? 'FALSE' : checked ? 'TRUE' : 'FALSE',
          checkbox: true,
        }
      })
    })
    setNotice('Checkboxes inserted')
  }

  function removeCheckboxes() {
    const points = selectedPointsForMutation(50_000)
    if (!points) return
    mutate((next) => {
      const s = activeSheet(next)
      points.forEach((target) => {
        const key = cellKey(target.row, target.col)
        const old = s.cells[key]
        if (!old) return
        s.cells[key] = { ...old, checkbox: undefined }
      })
    })
    setNotice('Checkbox controls removed')
  }

  function handleImageFile(file: File) {
    if (file.size > 1_500_000) {
      setNotice('Choose an image under 1.5 MB for local workbook storage')
      return
    }

    const reader = new FileReader()
    reader.onload = () => {
      if (typeof reader.result !== 'string') return
      mutate((next) => {
        const s = activeSheet(next)
        s.objects ||= []
        s.objects.push({
          id: crypto.randomUUID(),
          type: 'image',
          row: point.row,
          col: point.col,
          width: 240,
          height: 160,
          src: reader.result as string,
        })
      })
      setNotice('Image inserted')
    }
    reader.readAsDataURL(file)
  }

  function insertTextBox() {
    const text = window.prompt('Text box content', 'Text box')
    if (text === null) return
    mutate((next) => {
      const s = activeSheet(next)
      s.objects ||= []
      s.objects.push({
        id: crypto.randomUUID(),
        type: 'textBox',
        row: point.row,
        col: point.col,
        width: 220,
        height: 90,
        text,
        fill: '#ffffff',
        border: '#777777',
      })
    })
    setNotice('Text box inserted')
  }

  function insertShape() {
    const text = window.prompt('Shape text', 'Shape')
    if (text === null) return
    mutate((next) => {
      const s = activeSheet(next)
      s.objects ||= []
      s.objects.push({
        id: crypto.randomUUID(),
        type: 'shape',
        row: point.row,
        col: point.col,
        width: 180,
        height: 80,
        text,
        fill: '#e2f0d9',
        border: '#70ad47',
      })
    })
    setNotice('Shape inserted')
  }

  function insertCurrentChart() {
    const selected = normalizeSelection(selection)
    const data = chartDataForRange(sheet, selected, book.sheets)
    if (!data.length) {
      setNotice('Select numeric data before inserting a chart')
      return
    }

    const anchorCol = Math.min(COLS - 1, selected.right + 1)
    mutate((next) => {
      const s = activeSheet(next)
      s.objects ||= []
      s.objects.push({
        id: crypto.randomUUID(),
        type: 'chart',
        row: selected.top,
        col: anchorCol,
        width: 420,
        height: 260,
        chartType,
        chartRange: {
          top: selected.top,
          bottom: selected.bottom,
          left: selected.left,
          right: selected.right,
        },
        fill: '#ffffff',
        border: '#cfcfcf',
      })
    })
    setChartOpen(false)
    setNotice('Chart added to worksheet')
  }

  function removeSheetObject(id: string) {
    mutate((next) => {
      const s = activeSheet(next)
      s.objects = (s.objects || []).filter((object) => object.id !== id)
    })
    setNotice('Object removed')
  }

  function editSheetObject(id: string) {
    const object = (sheet.objects || []).find((item) => item.id === id)
    if (!object || object.type === 'image' || object.type === 'chart') return
    const text = window.prompt('Edit object text', object.text || '')
    if (text === null) return
    mutate((next) => {
      const s = activeSheet(next)
      const target = (s.objects || []).find((item) => item.id === id)
      if (target) target.text = text
    })
  }

  function addSheet() {
    mutate((next) => {
      const s = createBlankSheet(next.sheets.length + 1)
      s.name = uniqueSheetName(next.sheets, s.name)
      next.sheets.push(s)
      next.activeSheetId = s.id
    })

    selectPoint({ row: 0, col: 0 })
  }

  function duplicateSheet() {
    mutate((next) => {
      const source = activeSheet(next)
      const duplicate = structuredClone(source)
      duplicate.id = crypto.randomUUID()
      duplicate.name = uniqueSheetName(next.sheets, `${source.name} copy`)
      next.sheets.splice(next.sheets.findIndex((s) => s.id === source.id) + 1, 0, duplicate)
      next.activeSheetId = duplicate.id
    })

    setNotice('Sheet duplicated')
  }

  function renameSheet(item: SheetData) {
    const name = window.prompt('Rename sheet', item.name)?.trim()
    if (!name) return

    mutate((next) => {
      const target = next.sheets.find((s) => s.id === item.id)
      if (!target) return
      target.name = uniqueSheetName(
        next.sheets.filter((s) => s.id !== item.id),
        name,
      )
    })
  }

  function deleteSheet() {
    if (book.sheets.length === 1) return
    if (!window.confirm(`Delete "${sheet.name}"?`)) return

    mutate((next) => {
      const index = next.sheets.findIndex((s) => s.id === next.activeSheetId)
      next.sheets.splice(index, 1)
      next.activeSheetId = next.sheets[Math.max(0, index - 1)].id
    })
  }

  function createFreshWorkbook() {
    const used = book.sheets.some((s) => Object.keys(s.cells).length > 0)
    if (used && !window.confirm('Create a new workbook? Your current workbook is saved locally, but this will replace the active local workbook.')) return

    setHistory((items) => [...items, book].slice(-75))
    setFuture([])
    setBook(newWorkbook())
    selectPoint({ row: 0, col: 0 })
    setNotice('New workbook created')
  }

  function sortSelected(direction: 1 | -1) {
    if (range.top === range.bottom) {
      setNotice('Select two or more rows to sort')
      return
    }

    mutate((next) => {
      const s = activeSheet(next)
      const rows: CellData[][] = []

      for (let row = range.top; row <= range.bottom; row += 1) {
        const values: CellData[] = []
        for (let col = range.left; col <= range.right; col += 1) {
          values.push(structuredClone(getCell(s, row, col)))
        }
        rows.push(values)
      }

      rows.sort((a, b) => {
        const av = displayValue(a[0].value, s, next.sheets)
        const bv = displayValue(b[0].value, s, next.sheets)
        const an = Number(av)
        const bn = Number(bv)
        const result = Number.isFinite(an) && Number.isFinite(bn)
          ? an - bn
          : av.localeCompare(bv, undefined, { numeric: true })
        return result * direction
      })

      rows.forEach((values, ro) => {
        values.forEach((value, co) => {
          s.cells[cellKey(range.top + ro, range.left + co)] = value
        })
      })
    })
  }

  function replaceMatch(replaceAll = false) {
    const term = find.trim()
    if (!term) return

    const escaped = term.replace(/[.*+?^${}()|[\]\\]/g, '\\  function findNext() {')
    const pattern = new RegExp(escaped, 'gi')
    let replaced = 0

    mutate((next) => {
      const s = activeSheet(next)

      Object.entries(s.cells).forEach(([key, data]) => {
        if (!replaceAll) {
          const [row, col] = key.split(':').map(Number)
          if (row !== point.row || col !== point.col) return
        }

        if (!data.value || !pattern.test(data.value)) {
          pattern.lastIndex = 0
          return
        }

        pattern.lastIndex = 0
        s.cells[key] = { ...data, value: data.value.replace(pattern, replaceText) }
        replaced += 1
      })
    })

    if (!replaceAll && replaced === 0) {
      findNext()
      setNotice('Moved to the next match')
    } else {
      setNotice(replaced ? 'Replaced ' + replaced + ' cell' + (replaced === 1 ? '' : 's') : 'No matches found')
    }
  }

  function findNext() {
    const term = find.trim().toLowerCase()
    if (!term) return

    const start = point.row * COLS + point.col + 1

    for (let i = 0; i < ROWS * COLS; i += 1) {
      const n = (start + i) % (ROWS * COLS)
      const p = { row: Math.floor(n / COLS), col: n % COLS }

      if (getCell(sheet, p.row, p.col).value.toLowerCase().includes(term)) {
        selectPoint(p)
        document
          .querySelector('[data-cell="' + p.row + ':' + p.col + '"]')
          ?.scrollIntoView({ block: 'center', inline: 'center' })
        return
      }
    }

    setNotice('No match found')
  }

  function goToNameBox() {
    const normalizedName = nameBox.trim().toUpperCase()
    const named = sheet.namedRanges?.[normalizedName]
    const parsed = named
      ? {
          anchor: { row: named.top, col: named.left },
          focus: { row: named.bottom, col: named.right },
        }
      : parseSelectionAddress(nameBox)

    if (!parsed) {
      setNameBox(selectionToAddress(selection))
      setNotice('Enter a cell, range, or named range')
      return
    }

    setSelection(parsed)

    window.requestAnimationFrame(() => {
      document
        .querySelector('[data-cell="' + parsed.focus.row + ':' + parsed.focus.col + '"]')
        ?.scrollIntoView({ block: 'center', inline: 'center' })
      gridRef.current?.focus()
    })
  }

  function createNamedRange() {
    const name = namedRangeName.trim().toUpperCase()

    if (!/^[A-Z_][A-Z0-9_.]*$/.test(name) || parseSelectionAddress(name)) {
      setNotice('Use a name like SALES_TOTAL or Q1_DATA')
      return
    }

    mutate((next) => {
      const s = activeSheet(next)
      s.namedRanges ||= {}
      s.namedRanges[name] = {
        top: range.top,
        bottom: range.bottom,
        left: range.left,
        right: range.right,
      }
    })

    setNamedRangeName('')
    setNotice(`Named range ${name} created`)
  }

  function removeNamedRange(name: string) {
    mutate((next) => {
      const s = activeSheet(next)
      if (!s.namedRanges) return
      delete s.namedRanges[name]
    })
    setNotice(`Named range ${name} removed`)
  }

  function smartFunction(name: 'SUM' | 'AVERAGE' | 'MIN' | 'MAX' | 'COUNT' | 'PRODUCT' | 'MEDIAN') {
    const singleCell = range.top === range.bottom && range.left === range.right

    if (singleCell) {
      beginEdit(`=${name}(`)
      return
    }

    const formulaRange = `${pointToAddress({ row: range.top, col: range.left })}:${pointToAddress({ row: range.bottom, col: range.right })}`
    let target: Point | null = null

    if (range.left === range.right && range.bottom + 1 < ROWS) {
      target = { row: range.bottom + 1, col: range.left }
    } else if (range.top === range.bottom && range.right + 1 < COLS) {
      target = { row: range.top, col: range.right + 1 }
    } else if (range.bottom + 1 < ROWS) {
      target = { row: range.bottom + 1, col: range.left }
    }

    if (!target) {
      beginEdit(`=${name}(`)
      return
    }

    setValue(target, `=${name}(${formulaRange})`)
    selectPoint(target)
    setNotice(`${name} formula inserted`)
  }

  function adjustDecimals(delta: number) {
    const current = cell.format?.decimals ?? (cell.format?.numberFormat === 'currency' ? 2 : 0)
    applyFormat({ decimals: Math.max(0, Math.min(8, current + delta)) })
  }

  function beginColumnResize(event: React.MouseEvent, col: number) {
    event.preventDefault()
    event.stopPropagation()
    setHistory((items) => [...items, book].slice(-75))
    setFuture([])
    columnResizeRef.current = {
      col,
      startX: event.clientX,
      startWidth: getColumnWidth(sheet, col),
    }
    document.body.classList.add('is-column-resizing')
  }

  function autoFitColumn(col: number) {
    let longest = colToName(col).length

    for (let row = 0; row < ROWS; row += 1) {
      const value = formatted(getCell(sheet, row, col), sheet, book.sheets)
      longest = Math.max(longest, value.length)
    }

    const width = Math.max(58, Math.min(360, 20 + longest * 7.2))

    mutate((next) => {
      const s = activeSheet(next)
      s.columnWidths ||= {}
      s.columnWidths[String(col)] = Math.round(width)
    })

    setNotice(`AutoFit ${colToName(col)}`)
  }

  function adjustColumnWidth(delta: number) {
    mutate((next) => {
      const s = activeSheet(next)
      s.columnWidths ||= {}

      for (let col = range.left; col <= range.right; col += 1) {
        const width = s.columnWidths[String(col)] ?? DEFAULT_COLUMN_WIDTH
        s.columnWidths[String(col)] = Math.max(48, Math.min(360, width + delta))
      }
    })
  }

  function resetColumnWidth() {
    mutate((next) => {
      const s = activeSheet(next)
      s.columnWidths ||= {}

      for (let col = range.left; col <= range.right; col += 1) {
        delete s.columnWidths[String(col)]
      }
    })
  }

  function beginRowResize(event: React.MouseEvent, row: number) {
    event.preventDefault()
    event.stopPropagation()
    setHistory((items) => [...items, book].slice(-75))
    setFuture([])
    rowResizeRef.current = {
      row,
      startY: event.clientY,
      startHeight: getRowHeight(sheet, row),
    }
    document.body.classList.add('is-row-resizing')
  }

  function autoFitRow(row: number) {
    let height = DEFAULT_ROW_HEIGHT

    for (let col = 0; col < COLS; col += 1) {
      const data = getCell(sheet, row, col)
      const value = formatted(data, sheet, book.sheets)
      const fontSize = data.format?.fontSize || 11
      let lines = 1

      if (data.format?.wrap && value) {
        const charsPerLine = Math.max(4, Math.floor((getColumnWidth(sheet, col) - 10) / Math.max(5.5, fontSize * 0.58)))
        lines = Math.max(1, Math.ceil(value.length / charsPerLine))
      }

      height = Math.max(height, 8 + lines * Math.max(15, fontSize * 1.45))
    }

    mutate((next) => {
      const s = activeSheet(next)
      s.rowHeights ||= {}
      s.rowHeights[String(row)] = Math.round(Math.min(180, height))
    })

    setNotice(`AutoFit row ${row + 1}`)
  }

  function adjustRowHeight(delta: number) {
    mutate((next) => {
      const s = activeSheet(next)
      s.rowHeights ||= {}

      for (let row = range.top; row <= range.bottom; row += 1) {
        const height = s.rowHeights[String(row)] ?? DEFAULT_ROW_HEIGHT
        s.rowHeights[String(row)] = Math.max(20, Math.min(180, height + delta))
      }
    })
  }

  function resetRowHeight() {
    mutate((next) => {
      const s = activeSheet(next)
      s.rowHeights ||= {}

      for (let row = range.top; row <= range.bottom; row += 1) {
        delete s.rowHeights[String(row)]
      }
    })
  }

  function restructureSheet(s: SheetData, axis: 'row' | 'col', index: number, delta: 1 | -1) {
    const moved: Record<string, CellData> = {}

    Object.entries(s.cells).forEach(([key, data]) => {
      let [row, col] = key.split(':').map(Number)

      if (axis === 'row') {
        if (delta === -1 && row === index) return
        if (delta === 1 && row >= index) row += 1
        if (delta === -1 && row > index) row -= 1
      } else {
        if (delta === -1 && col === index) return
        if (delta === 1 && col >= index) col += 1
        if (delta === -1 && col > index) col -= 1
      }

      if (row < 0 || row >= ROWS || col < 0 || col >= COLS) return

      moved[cellKey(row, col)] = {
        ...structuredClone(data),
        value: shiftFormulaForStructure(data.value, axis, index, delta),
      }
    })

    s.cells = moved

    const movedNotes: NonNullable<SheetData['notes']> = {}
    Object.entries(s.notes || {}).forEach(([key, note]) => {
      let [row, col] = key.split(':').map(Number)

      if (axis === 'row') {
        if (delta === -1 && row === index) return
        if (delta === 1 && row >= index) row += 1
        if (delta === -1 && row > index) row -= 1
      } else {
        if (delta === -1 && col === index) return
        if (delta === 1 && col >= index) col += 1
        if (delta === -1 && col > index) col -= 1
      }

      if (row >= 0 && row < ROWS && col >= 0 && col < COLS) {
        movedNotes[cellKey(row, col)] = note
      }
    })
    s.notes = movedNotes

    s.objects = (s.objects || []).flatMap((object) => {
      const nextObject = { ...object }

      if (axis === 'row') {
        if (delta === -1 && nextObject.row === index) return []
        if (delta === 1 && nextObject.row >= index) nextObject.row += 1
        if (delta === -1 && nextObject.row > index) nextObject.row -= 1
      } else {
        if (delta === -1 && nextObject.col === index) return []
        if (delta === 1 && nextObject.col >= index) nextObject.col += 1
        if (delta === -1 && nextObject.col > index) nextObject.col -= 1
      }

      if (nextObject.chartRange) {
        const chartRange = { ...nextObject.chartRange }

        if (axis === 'row') {
          if (delta === 1 && index <= chartRange.top) {
            chartRange.top += 1
            chartRange.bottom += 1
          } else if (delta === 1 && index <= chartRange.bottom) {
            chartRange.bottom = Math.min(ROWS - 1, chartRange.bottom + 1)
          } else if (delta === -1 && index < chartRange.top) {
            chartRange.top = Math.max(0, chartRange.top - 1)
            chartRange.bottom = Math.max(chartRange.top, chartRange.bottom - 1)
          } else if (delta === -1 && index <= chartRange.bottom) {
            chartRange.bottom = Math.max(chartRange.top, chartRange.bottom - 1)
          }
        } else {
          if (delta === 1 && index <= chartRange.left) {
            chartRange.left += 1
            chartRange.right += 1
          } else if (delta === 1 && index <= chartRange.right) {
            chartRange.right = Math.min(COLS - 1, chartRange.right + 1)
          } else if (delta === -1 && index < chartRange.left) {
            chartRange.left = Math.max(0, chartRange.left - 1)
            chartRange.right = Math.max(chartRange.left, chartRange.right - 1)
          } else if (delta === -1 && index <= chartRange.right) {
            chartRange.right = Math.max(chartRange.left, chartRange.right - 1)
          }
        }

        nextObject.chartRange = chartRange
      }

      if (nextObject.row < 0 || nextObject.row >= ROWS || nextObject.col < 0 || nextObject.col >= COLS) return []
      return [nextObject]
    })

    const shiftDimensionMap = (
      source: Record<string, number> | undefined,
      max: number,
    ) => {
      const result: Record<string, number> = {}

      Object.entries(source || {}).forEach(([key, value]) => {
        let position = Number(key)

        if (delta === -1 && position === index) return
        if (delta === 1 && position >= index) position += 1
        if (delta === -1 && position > index) position -= 1
        if (position >= 0 && position < max) result[String(position)] = value
      })

      return result
    }

    const shiftHiddenMap = (
      source: Record<string, boolean> | undefined,
      max: number,
    ) => {
      const result: Record<string, boolean> = {}
      Object.entries(source || {}).forEach(([key, value]) => {
        let position = Number(key)
        if (delta === -1 && position === index) return
        if (delta === 1 && position >= index) position += 1
        if (delta === -1 && position > index) position -= 1
        if (position >= 0 && position < max && value) result[String(position)] = true
      })
      return result
    }

    if (axis === 'row') {
      s.rowHeights = shiftDimensionMap(s.rowHeights, ROWS)
      s.hiddenRows = shiftHiddenMap(s.hiddenRows, ROWS)
    } else {
      s.columnWidths = shiftDimensionMap(s.columnWidths, COLS)
      s.hiddenColumns = shiftHiddenMap(s.hiddenColumns, COLS)
    }

    s.merges = (s.merges || []).flatMap((merge) => {
      const nextMerge = { ...merge }

      if (axis === 'row') {
        if (delta === 1 && index <= nextMerge.top) {
          nextMerge.top += 1
          nextMerge.bottom += 1
        } else if (delta === 1 && index <= nextMerge.bottom) {
          nextMerge.bottom = Math.min(ROWS - 1, nextMerge.bottom + 1)
        } else if (delta === -1 && index < nextMerge.top) {
          nextMerge.top = Math.max(0, nextMerge.top - 1)
          nextMerge.bottom = Math.max(nextMerge.top, nextMerge.bottom - 1)
        } else if (delta === -1 && index <= nextMerge.bottom) {
          if (nextMerge.top === nextMerge.bottom) return []
          nextMerge.bottom -= 1
        }
      } else {
        if (delta === 1 && index <= nextMerge.left) {
          nextMerge.left += 1
          nextMerge.right += 1
        } else if (delta === 1 && index <= nextMerge.right) {
          nextMerge.right = Math.min(COLS - 1, nextMerge.right + 1)
        } else if (delta === -1 && index < nextMerge.left) {
          nextMerge.left = Math.max(0, nextMerge.left - 1)
          nextMerge.right = Math.max(nextMerge.left, nextMerge.right - 1)
        } else if (delta === -1 && index <= nextMerge.right) {
          if (nextMerge.left === nextMerge.right) return []
          nextMerge.right -= 1
        }
      }

      return [nextMerge]
    })

    if (s.filterRange) {
      const filterRange = s.filterRange

      if (axis === 'row') {
        if (delta === 1 && index <= filterRange.top) {
          filterRange.top += 1
          filterRange.bottom += 1
        } else if (delta === 1 && index <= filterRange.bottom) {
          filterRange.bottom = Math.min(ROWS - 1, filterRange.bottom + 1)
        } else if (delta === -1 && index < filterRange.top) {
          filterRange.top = Math.max(0, filterRange.top - 1)
          filterRange.bottom = Math.max(filterRange.top, filterRange.bottom - 1)
        } else if (delta === -1 && index <= filterRange.bottom) {
          filterRange.bottom = Math.max(filterRange.top, filterRange.bottom - 1)
        }
      } else {
        if (delta === 1 && index <= filterRange.left) {
          filterRange.left += 1
          filterRange.right += 1
        } else if (delta === 1 && index <= filterRange.right) {
          filterRange.right = Math.min(COLS - 1, filterRange.right + 1)
        } else if (delta === -1 && index < filterRange.left) {
          filterRange.left = Math.max(0, filterRange.left - 1)
          filterRange.right = Math.max(filterRange.left, filterRange.right - 1)
        } else if (delta === -1 && index <= filterRange.right) {
          filterRange.right = Math.max(filterRange.left, filterRange.right - 1)
        }

        const shiftedFilters: NonNullable<SheetData['filters']> = {}
        Object.entries(s.filters || {}).forEach(([key, rule]) => {
          let col = Number(key)
          if (delta === -1 && col === index) return
          if (delta === 1 && col >= index) col += 1
          if (delta === -1 && col > index) col -= 1
          if (col >= 0 && col < COLS) shiftedFilters[String(col)] = rule
        })
        s.filters = shiftedFilters
      }
    }

    s.namedRanges = Object.fromEntries(
      Object.entries(s.namedRanges || {}).flatMap(([name, named]) => {
        const nextNamed = { ...named }

        if (axis === 'row') {
          if (delta === 1 && index <= nextNamed.top) {
            nextNamed.top += 1
            nextNamed.bottom += 1
          } else if (delta === 1 && index <= nextNamed.bottom) {
            nextNamed.bottom = Math.min(ROWS - 1, nextNamed.bottom + 1)
          } else if (delta === -1 && index < nextNamed.top) {
            nextNamed.top = Math.max(0, nextNamed.top - 1)
            nextNamed.bottom = Math.max(nextNamed.top, nextNamed.bottom - 1)
          } else if (delta === -1 && index <= nextNamed.bottom) {
            if (nextNamed.top === nextNamed.bottom) return []
            nextNamed.bottom -= 1
          }
        } else {
          if (delta === 1 && index <= nextNamed.left) {
            nextNamed.left += 1
            nextNamed.right += 1
          } else if (delta === 1 && index <= nextNamed.right) {
            nextNamed.right = Math.min(COLS - 1, nextNamed.right + 1)
          } else if (delta === -1 && index < nextNamed.left) {
            nextNamed.left = Math.max(0, nextNamed.left - 1)
            nextNamed.right = Math.max(nextNamed.left, nextNamed.right - 1)
          } else if (delta === -1 && index <= nextNamed.right) {
            if (nextNamed.left === nextNamed.right) return []
            nextNamed.right -= 1
          }
        }

        return [[name, nextNamed] as const]
      }),
    )

    s.tables = (s.tables || []).flatMap((table) => {
      const nextTable = { ...table }

      if (axis === 'row') {
        if (delta === 1 && index <= nextTable.top) {
          nextTable.top += 1
          nextTable.bottom += 1
        } else if (delta === 1 && index <= nextTable.bottom) {
          nextTable.bottom = Math.min(ROWS - 1, nextTable.bottom + 1)
        } else if (delta === -1 && index < nextTable.top) {
          nextTable.top = Math.max(0, nextTable.top - 1)
          nextTable.bottom = Math.max(nextTable.top, nextTable.bottom - 1)
        } else if (delta === -1 && index <= nextTable.bottom) {
          if (nextTable.top === nextTable.bottom) return []
          nextTable.bottom -= 1
        }
      } else {
        if (delta === 1 && index <= nextTable.left) {
          nextTable.left += 1
          nextTable.right += 1
        } else if (delta === 1 && index <= nextTable.right) {
          nextTable.right = Math.min(COLS - 1, nextTable.right + 1)
        } else if (delta === -1 && index < nextTable.left) {
          nextTable.left = Math.max(0, nextTable.left - 1)
          nextTable.right = Math.max(nextTable.left, nextTable.right - 1)
        } else if (delta === -1 && index <= nextTable.right) {
          if (nextTable.left === nextTable.right) return []
          nextTable.right -= 1
        }
      }

      return [nextTable]
    })

    s.dataValidations = (s.dataValidations || []).flatMap((rule) => {
      const nextRule = { ...rule }

      if (axis === 'row') {
        if (delta === 1 && index <= nextRule.top) {
          nextRule.top += 1
          nextRule.bottom += 1
        } else if (delta === 1 && index <= nextRule.bottom) {
          nextRule.bottom = Math.min(ROWS - 1, nextRule.bottom + 1)
        } else if (delta === -1 && index < nextRule.top) {
          nextRule.top = Math.max(0, nextRule.top - 1)
          nextRule.bottom = Math.max(nextRule.top, nextRule.bottom - 1)
        } else if (delta === -1 && index <= nextRule.bottom) {
          if (nextRule.top === nextRule.bottom) return []
          nextRule.bottom -= 1
        }
      } else {
        if (delta === 1 && index <= nextRule.left) {
          nextRule.left += 1
          nextRule.right += 1
        } else if (delta === 1 && index <= nextRule.right) {
          nextRule.right = Math.min(COLS - 1, nextRule.right + 1)
        } else if (delta === -1 && index < nextRule.left) {
          nextRule.left = Math.max(0, nextRule.left - 1)
          nextRule.right = Math.max(nextRule.left, nextRule.right - 1)
        } else if (delta === -1 && index <= nextRule.right) {
          if (nextRule.left === nextRule.right) return []
          nextRule.right -= 1
        }
      }

      return [nextRule]
    })

    s.conditionalFormats = (s.conditionalFormats || []).flatMap((rule) => {
      const nextRule = { ...rule }

      if (axis === 'row') {
        if (delta === 1 && index <= nextRule.top) {
          nextRule.top += 1
          nextRule.bottom += 1
        } else if (delta === 1 && index <= nextRule.bottom) {
          nextRule.bottom = Math.min(ROWS - 1, nextRule.bottom + 1)
        } else if (delta === -1 && index < nextRule.top) {
          nextRule.top = Math.max(0, nextRule.top - 1)
          nextRule.bottom = Math.max(nextRule.top, nextRule.bottom - 1)
        } else if (delta === -1 && index <= nextRule.bottom) {
          if (nextRule.top === nextRule.bottom) return []
          nextRule.bottom -= 1
        }
      } else {
        if (delta === 1 && index <= nextRule.left) {
          nextRule.left += 1
          nextRule.right += 1
        } else if (delta === 1 && index <= nextRule.right) {
          nextRule.right = Math.min(COLS - 1, nextRule.right + 1)
        } else if (delta === -1 && index < nextRule.left) {
          nextRule.left = Math.max(0, nextRule.left - 1)
          nextRule.right = Math.max(nextRule.left, nextRule.right - 1)
        } else if (delta === -1 && index <= nextRule.right) {
          if (nextRule.left === nextRule.right) return []
          nextRule.right -= 1
        }
      }

      return [nextRule]
    })
  }

  function insertRow() {
    const index = range.top
    mutate((next) => restructureSheet(activeSheet(next), 'row', index, 1))
    selectPoint({ row: index, col: point.col })
    setNotice(`Inserted row ${index + 1}`)
  }

  function deleteRow() {
    const index = range.top
    mutate((next) => restructureSheet(activeSheet(next), 'row', index, -1))
    selectPoint({ row: Math.min(index, ROWS - 1), col: point.col })
    setNotice(`Deleted row ${index + 1}`)
  }

  function insertColumn() {
    const index = range.left
    mutate((next) => restructureSheet(activeSheet(next), 'col', index, 1))
    selectPoint({ row: point.row, col: index })
    setNotice(`Inserted column ${colToName(index)}`)
  }

  function deleteColumn() {
    const index = range.left
    mutate((next) => restructureSheet(activeSheet(next), 'col', index, -1))
    selectPoint({ row: point.row, col: Math.min(index, COLS - 1) })
    setNotice(`Deleted column ${colToName(index)}`)
  }

  function openNote(target: Point = point) {
    const existing = sheet.notes?.[cellKey(target.row, target.col)]?.text || ''
    setNoteTarget(target)
    setNoteDraft(existing)
    setNoteOpen(true)
  }

  function saveNote() {
    mutate((next) => {
      const s = activeSheet(next)
      s.notes ||= {}
      const key = cellKey(noteTarget.row, noteTarget.col)
      const text = noteDraft.trim()

      if (!text) delete s.notes[key]
      else {
        s.notes[key] = {
          text,
          updatedAt: Date.now(),
        }
      }
    })

    setNoteOpen(false)
    setNotice(noteDraft.trim() ? 'Note saved' : 'Note removed')
  }

  function removeNote() {
    mutate((next) => {
      const s = activeSheet(next)
      if (s.notes) delete s.notes[cellKey(noteTarget.row, noteTarget.col)]
    })
    setNoteDraft('')
    setNoteOpen(false)
    setNotice('Note removed')
  }

  function toggleProtection() {
    mutate((next) => {
      const s = activeSheet(next)
      s.protected = !s.protected
    }, true)

    setNotice(sheet.protected ? 'Worksheet unprotected' : 'Worksheet protected')
  }

  function saveVersion(label = 'Manual version') {
    const snapshot: VersionSnapshot = {
      id: crypto.randomUUID(),
      savedAt: Date.now(),
      label,
      workbook: structuredClone(book),
    }
    setVersions((current) => [snapshot, ...current].slice(0, 12))
    setNotice('Workbook version saved')
  }

  function restoreVersion(snapshot: VersionSnapshot) {
    setHistory((items) => [...items, book].slice(-75))
    setFuture([])
    setBook(structuredClone(snapshot.workbook))
    setVersionOpen(false)
    setNotice(`Restored version from ${new Date(snapshot.savedAt).toLocaleString()}`)
  }

  function deleteVersion(id: string) {
    setVersions((current) => current.filter((version) => version.id !== id))
  }

  function updatePageLayout(patch: Partial<NonNullable<SheetData['pageLayout']>>) {
    mutate((next) => {
      const s = activeSheet(next)
      s.pageLayout = {
        orientation: 'portrait',
        paperSize: 'letter',
        margins: 'normal',
        printGridlines: true,
        ...(s.pageLayout || {}),
        ...patch,
      }
    }, true)
  }

  function setPrintAreaFromSelection() {
    const selected = normalizeSelection(selection)
    updatePageLayout({
      printArea: {
        top: selected.top,
        bottom: selected.bottom,
        left: selected.left,
        right: selected.right,
      },
    })
    setNotice('Print area set to ' + selectionToAddress(selection))
  }

  function clearPrintArea() {
    mutate((next) => {
      const s = activeSheet(next)
      s.pageLayout = {
        orientation: 'portrait',
        paperSize: 'letter',
        margins: 'normal',
        printGridlines: true,
        scalePercent: 100,
        headerText: '',
        footerText: '',
        ...(s.pageLayout || {}),
        printArea: undefined,
      }
    }, true)
    setNotice('Print area cleared')
  }

  function editPrintHeader() {
    const value = window.prompt('Print header text', sheet.pageLayout?.headerText || book.title)
    if (value === null) return
    updatePageLayout({ headerText: value })
  }

  function editPrintFooter() {
    const value = window.prompt('Print footer text', sheet.pageLayout?.footerText || sheet.name)
    if (value === null) return
    updatePageLayout({ footerText: value })
  }

  function printSheet() {
    window.requestAnimationFrame(() => window.print())
  }

  function openPivotBuilder() {
    const selected = normalizeSelection(selection)
    const target = selected.bottom > selected.top && selected.right > selected.left
      ? selected
      : usedRange()

    if (target.bottom <= target.top || target.right < target.left) {
      setNotice('Select a table-like range with headers and data first')
      return
    }

    setPivotSource(target)
    setPivotRowField(target.left)
    setPivotValueField(Math.min(target.right, target.left + 1))
    setPivotAggregator('sum')
    setPivotOpen(true)
  }

  function createPivot() {
    const rowHeader = formatted(getCell(sheet, pivotSource.top, pivotRowField), sheet, book.sheets) || colToName(pivotRowField)
    const valueHeader = formatted(getCell(sheet, pivotSource.top, pivotValueField), sheet, book.sheets) || colToName(pivotValueField)
    const grouped = new Map<string, { sum: number; count: number }>()

    for (let row = pivotSource.top + 1; row <= pivotSource.bottom; row += 1) {
      const label = formatted(getCell(sheet, row, pivotRowField), sheet, book.sheets) || '(blank)'
      const rawValue = displayValue(getCell(sheet, row, pivotValueField).value, sheet, book.sheets)
      const numeric = Number(rawValue.replace(/[$,%]/g, ''))
      const current = grouped.get(label) || { sum: 0, count: 0 }

      if (pivotAggregator === 'count') {
        if (rawValue !== '') current.count += 1
      } else if (Number.isFinite(numeric)) {
        current.sum += numeric
        current.count += 1
      }

      grouped.set(label, current)
    }

    const output = createBlankSheet(book.sheets.length + 1)
    output.name = uniqueSheetName(book.sheets, 'Pivot')
    output.cells[cellKey(0, 0)] = { value: rowHeader, format: { bold: true, background: '#107c41', color: '#ffffff' } }
    output.cells[cellKey(0, 1)] = {
      value: `${pivotAggregator.toUpperCase()} of ${valueHeader}`,
      format: { bold: true, background: '#107c41', color: '#ffffff' },
    }

    Array.from(grouped.entries())
      .sort(([a], [b]) => a.localeCompare(b, undefined, { numeric: true }))
      .forEach(([label, summary], index) => {
        const row = index + 1
        const value = pivotAggregator === 'count'
          ? summary.count
          : pivotAggregator === 'average'
            ? (summary.count ? summary.sum / summary.count : 0)
            : summary.sum
        output.cells[cellKey(row, 0)] = { value: label }
        output.cells[cellKey(row, 1)] = { value: String(value), format: { numberFormat: 'number', decimals: 2 } }
      })

    output.columnWidths = { '0': 180, '1': 140 }

    mutate((next) => {
      next.sheets.push(output)
      next.activeSheetId = output.id
    }, true)

    setPivotOpen(false)
    selectPoint({ row: 0, col: 0 })
    setNotice('Pivot summary created on a new sheet')
  }

  function usedRange() {
    let top = ROWS - 1
    let bottom = 0
    let left = COLS - 1
    let right = 0
    let found = false

    Object.keys(sheet.cells).forEach((key) => {
      const [row, col] = key.split(':').map(Number)
      found = true
      top = Math.min(top, row)
      bottom = Math.max(bottom, row)
      left = Math.min(left, col)
      right = Math.max(right, col)
    })

    if (!found) return { top: 0, bottom: 1, left: 0, right: 0 }
    return {
      top,
      bottom: Math.max(top + 1, bottom),
      left,
      right,
    }
  }

  function toggleFilterRange() {
    if (sheet.filterRange) {
      mutate((next) => {
        const s = activeSheet(next)
        delete s.filterRange
        s.filters = {}
      })
      setFilterEditor(null)
      setNotice('Filters removed')
      return
    }

    const selected = normalizeSelection(selection)
    const target = selected.bottom > selected.top ? selected : usedRange()

    mutate((next) => {
      const s = activeSheet(next)
      s.filterRange = {
        top: target.top,
        bottom: Math.min(ROWS - 1, target.bottom),
        left: target.left,
        right: target.right,
      }
      s.filters = {}
    })

    setNotice('Filter headers enabled')
  }

  function openFilterEditor(col: number) {
    const rule = sheet.filters?.[String(col)]
    setFilterEditor({
      col,
      operator: rule?.operator || 'contains',
      value: rule?.value || '',
    })
  }

  function applyFilterEditor() {
    if (!filterEditor) return

    mutate((next) => {
      const s = activeSheet(next)
      s.filters ||= {}
      s.filters[String(filterEditor.col)] = {
        operator: filterEditor.operator,
        value: filterEditor.operator === 'notBlank' ? '' : filterEditor.value,
      }
    })

    setFilterEditor(null)
    setNotice(`Filter applied to ${colToName(filterEditor.col)}`)
  }

  function clearFilterColumn(col: number) {
    mutate((next) => {
      const s = activeSheet(next)
      if (!s.filters) return
      delete s.filters[String(col)]
    })
    setFilterEditor(null)
    setNotice(`Filter cleared from ${colToName(col)}`)
  }

  function clearAllFilters() {
    mutate((next) => {
      activeSheet(next).filters = {}
    })
    setNotice('Filter conditions cleared')
  }

  function nextTableName() {
    const existing = new Set((sheet.tables || []).map((table) => table.name.toLowerCase()))
    let index = 1
    while (existing.has(`table${index}`)) index += 1
    return `Table${index}`
  }

  function openTableDialog() {
    setTableName(nextTableName())
    setTableOpen(true)
  }

  function addTable() {
    const selected = normalizeSelection(selection)
    const target = selected.bottom > selected.top ? selected : usedRange()

    if (target.bottom <= target.top) {
      setNotice('Select a header row plus at least one data row')
      return
    }

    let name = tableName.trim().replace(/[^A-Za-z0-9_]/g, '_') || nextTableName()
    const existing = new Set((sheet.tables || []).map((table) => table.name.toLowerCase()))
    if (existing.has(name.toLowerCase())) {
      let suffix = 2
      const base = name
      while (existing.has(`${base}_${suffix}`.toLowerCase())) suffix += 1
      name = `${base}_${suffix}`
    }

    mutate((next) => {
      const s = activeSheet(next)
      s.tables ||= []
      s.tables.push({
        id: crypto.randomUUID(),
        name,
        top: target.top,
        bottom: target.bottom,
        left: target.left,
        right: target.right,
        style: tableStyle,
        bandedRows: tableBandedRows,
      })
      s.filterRange = { ...target }
      s.filters = {}
    })

    setSelection({
      anchor: { row: target.top, col: target.left },
      focus: { row: target.bottom, col: target.right },
    })
    setTableOpen(false)
    setNotice(`${name} created`)
  }

  function removeTable(id: string) {
    mutate((next) => {
      const s = activeSheet(next)
      const table = (s.tables || []).find((item) => item.id === id)
      s.tables = (s.tables || []).filter((item) => item.id !== id)

      if (
        table &&
        s.filterRange &&
        s.filterRange.top === table.top &&
        s.filterRange.bottom === table.bottom &&
        s.filterRange.left === table.left &&
        s.filterRange.right === table.right
      ) {
        delete s.filterRange
        s.filters = {}
      }
    })
    setNotice('Table removed')
  }

  function addDataValidation() {
    const options = Array.from(new Set(
      validationOptions
        .split(/[\n,]/)
        .map((item) => item.trim())
        .filter(Boolean),
    ))

    if (validationType === 'list' && !options.length) {
      setNotice('Add at least one dropdown option')
      return
    }

    if (validationType !== 'list' && !validationMinimum.trim()) {
      setNotice('Enter a validation value')
      return
    }

    if (validationType !== 'list' && validationOperator === 'between' && !validationMaximum.trim()) {
      setNotice('Enter both ends of the validation range')
      return
    }

    mutate((next) => {
      const s = activeSheet(next)
      s.dataValidations ||= []
      s.dataValidations.push({
        id: crypto.randomUUID(),
        top: range.top,
        bottom: range.bottom,
        left: range.left,
        right: range.right,
        type: validationType,
        options: validationType === 'list' ? options : undefined,
        operator: validationType === 'list' ? undefined : validationOperator,
        minimum: validationType === 'list' ? undefined : validationMinimum,
        maximum: validationType === 'list' ? undefined : validationMaximum,
        allowBlank: validationAllowBlank,
        inputTitle: validationInputTitle.trim() || undefined,
        inputMessage: validationInputMessage.trim() || undefined,
        errorStyle: validationErrorStyle,
      })
    })

    setValidationOpen(false)
    setNotice(validationType === 'list' ? 'Dropdown validation added' : 'Validation rule added')
  }

  function clearDataValidation() {
    mutate((next) => {
      const s = activeSheet(next)
      s.dataValidations = (s.dataValidations || []).filter((rule) => (
        rule.bottom < range.top ||
        rule.top > range.bottom ||
        rule.right < range.left ||
        rule.left > range.right
      ))
    })
    setValidationPicker(null)
    setNotice('Validation cleared from selection')
  }

  function chooseValidationValue(value: string) {
    if (!validationPicker) return
    setValue({ row: validationPicker.row, col: validationPicker.col }, value)
    selectPoint({ row: validationPicker.row, col: validationPicker.col })
    setValidationPicker(null)
  }

  function addConditionalFormat() {
    mutate((next) => {
      const s = activeSheet(next)
      s.conditionalFormats ||= []
      s.conditionalFormats.push({
        id: crypto.randomUUID(),
        top: range.top,
        bottom: range.bottom,
        left: range.left,
        right: range.right,
        operator: conditionalOperator,
        value: conditionalOperator === 'notBlank' ? '' : conditionalValue,
        background: conditionalBackground,
        color: conditionalColor,
      })
    })

    setConditionalOpen(false)
    setNotice('Conditional formatting rule added')
  }

  function addConditionalVisual(kind: Exclude<ConditionalFormatKind, 'cell'>) {
    mutate((next) => {
      const s = activeSheet(next)
      s.conditionalFormats ||= []
      s.conditionalFormats.push({
        id: crypto.randomUUID(),
        top: range.top,
        bottom: range.bottom,
        left: range.left,
        right: range.right,
        operator: 'notBlank',
        background: kind === 'duplicate' ? '#ffc7ce' : kind === 'unique' ? '#c6efce' : '#ffffff',
        color: kind === 'duplicate' ? '#9c0006' : kind === 'unique' ? '#006100' : '#202020',
        kind,
        dataBarColor: '#63be7b',
        minColor: '#f8696b',
        maxColor: '#63be7b',
      })
    })
    setNotice(
      kind === 'dataBar'
        ? 'Data bars applied'
        : kind === 'colorScale'
          ? 'Color scale applied'
          : kind === 'duplicate'
            ? 'Duplicate values highlighted'
            : 'Unique values highlighted',
    )
  }

  function removeConditionalFormat(id: string) {
    mutate((next) => {
      const s = activeSheet(next)
      s.conditionalFormats = (s.conditionalFormats || []).filter((rule) => rule.id !== id)
    })
  }

  function clearConditionalFormats() {
    mutate((next) => {
      activeSheet(next).conditionalFormats = []
    })
    setNotice('Conditional formatting cleared')
  }

  function hideSelectedRows() {
    mutate((next) => {
      const s = activeSheet(next)
      s.hiddenRows ||= {}
      for (let row = range.top; row <= range.bottom; row += 1) s.hiddenRows[String(row)] = true
    })
    setNotice('Selected rows hidden')
  }

  function unhideAllRows() {
    mutate((next) => { activeSheet(next).hiddenRows = {} })
    setNotice('All rows unhidden')
  }

  function hideSelectedColumns() {
    mutate((next) => {
      const s = activeSheet(next)
      s.hiddenColumns ||= {}
      for (let col = range.left; col <= range.right; col += 1) s.hiddenColumns[String(col)] = true
    })
    setNotice('Selected columns hidden')
  }

  function unhideAllColumns() {
    mutate((next) => { activeSheet(next).hiddenColumns = {} })
    setNotice('All columns unhidden')
  }

  function toggleSheetView(key: 'showGridlines' | 'freezeTopRow' | 'freezeFirstColumn') {
    mutate((next) => {
      const s = activeSheet(next)
      if (key === 'showGridlines') s.showGridlines = s.showGridlines === false
      else s[key] = !s[key]
    })
  }

  function exportXlsx() {
    const out = XLSX.utils.book_new()

    book.sheets.forEach((s) => {
      let maxRow = 0
      let maxCol = 0

      Object.keys(s.cells).forEach((key) => {
        const bits = key.split(':').map(Number)
        maxRow = Math.max(maxRow, bits[0])
        maxCol = Math.max(maxCol, bits[1])
      })

      const rows: (string | number)[][] = []

      for (let row = 0; row <= maxRow; row += 1) {
        const values: (string | number)[] = []

        for (let col = 0; col <= maxCol; col += 1) {
          const raw = getCell(s, row, col).value
          const n = Number(raw)
          values.push(raw && !raw.startsWith('=') && Number.isFinite(n) ? n : raw)
        }

        rows.push(values)
      }

      const ws = XLSX.utils.aoa_to_sheet(rows)

      for (let row = 0; row <= maxRow; row += 1) {
        for (let col = 0; col <= maxCol; col += 1) {
          const raw = getCell(s, row, col).value
          if (!raw.startsWith('=')) continue

          const address = XLSX.utils.encode_cell({ r: row, c: col })
          ws[address] = { t: 'n', f: raw.slice(1) }
        }
      }

      ws['!cols'] = Array.from({ length: maxCol + 1 }, (_, col) => ({
        wch: Math.max(6, Math.round(getColumnWidth(s, col) / 7)),
      }))

      XLSX.utils.book_append_sheet(out, ws, s.name.slice(0, 31))
    })

    XLSX.writeFile(out, safeName(book.title) + '.xlsx')
  }

  function fromRows(rows: unknown[][], name: string): SheetData {
    const s = createBlankSheet(1)
    s.name = name.slice(0, 31) || 'Sheet1'

    rows.slice(0, ROWS).forEach((row, r) => {
      row.slice(0, COLS).forEach((value, c) => {
        if (value !== '' && value != null) {
          s.cells[cellKey(r, c)] = { value: String(value) }
        }
      })
    })

    return s
  }

  async function importFile(file: File) {
    const ext = file.name.split('.').pop()?.toLowerCase()

    if (ext === 'json') {
      const parsed = JSON.parse(await file.text()) as WorkbookData
      if (!parsed.sheets?.length) throw new Error('Invalid Excel Micro backup')
      setHistory((items) => [...items, book].slice(-75))
      setBook(parsed)
      setNotice('Workbook restored')
      return
    }

    if (ext === 'csv') {
      const s = fromRows(parseCsv(await file.text()), file.name.replace(/\.csv$/i, ''))

      mutate((next) => {
        s.name = uniqueSheetName(next.sheets, s.name)
        next.sheets.push(s)
        next.activeSheetId = s.id
      })

      setNotice('CSV imported')
      return
    }

    const source = XLSX.read(await file.arrayBuffer(), {
      type: 'array',
      cellFormula: true,
      cellStyles: true,
    })

    const sheets = source.SheetNames.map((name) => {
      const ws = source.Sheets[name]
      const ref = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1')
      const rows: string[][] = []

      for (let r = ref.s.r; r <= Math.min(ref.e.r, ROWS - 1); r += 1) {
        const values: string[] = []

        for (let c = ref.s.c; c <= Math.min(ref.e.c, COLS - 1); c += 1) {
          const sourceCell = ws[XLSX.utils.encode_cell({ r, c })]
          values.push(
            sourceCell?.f
              ? '=' + sourceCell.f
              : sourceCell?.v == null
                ? ''
                : String(sourceCell.v),
          )
        }

        rows.push(values)
      }

      const imported = fromRows(rows, name)
      const columnInfo = ws['!cols'] as Array<{ wpx?: number; wch?: number }> | undefined

      if (columnInfo?.length) {
        imported.columnWidths = {}

        columnInfo.slice(0, COLS).forEach((info, index) => {
          const width = info?.wpx ?? (info?.wch ? info.wch * 7 : 0)
          if (width > 0) imported.columnWidths![String(index)] = Math.max(48, Math.min(360, Math.round(width)))
        })
      }

      return imported
    })

    if (!sheets.length) return

    setHistory((items) => [...items, book].slice(-75))
    setBook({
      id: crypto.randomUUID(),
      title: file.name.replace(/\.[^.]+$/, ''),
      sheets,
      activeSheetId: sheets[0].id,
      updatedAt: Date.now(),
    })
    selectPoint({ row: 0, col: 0 })
    setNotice('Excel workbook imported')
  }

  const stats = useMemo(() => {
    const selected = normalizeSelection(selection)
    const area = (selected.bottom - selected.top + 1) * (selected.right - selected.left + 1)
    const nums: number[] = []

    if (area <= 100_000) {
      for (let row = selected.top; row <= selected.bottom; row += 1) {
        for (let col = selected.left; col <= selected.right; col += 1) {
          const value = Number(displayValue(getCell(sheet, row, col).value, sheet, book.sheets))
          if (Number.isFinite(value)) nums.push(value)
        }
      }
    } else {
      Object.entries(sheet.cells).forEach(([key, data]) => {
        const [row, col] = key.split(':').map(Number)
        if (
          row < selected.top ||
          row > selected.bottom ||
          col < selected.left ||
          col > selected.right
        ) return
        const value = Number(displayValue(data.value, sheet, book.sheets))
        if (Number.isFinite(value)) nums.push(value)
      })
    }

    return {
      count: nums.length,
      sum: nums.reduce((a, b) => a + b, 0),
      avg: nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0,
    }
  }, [book.sheets, selection, sheet])

  const chart = useMemo(
    () => chartDataForRange(sheet, range, book.sheets).slice(0, 24),
    [book.sheets, range.bottom, range.left, range.right, range.top, sheet],
  )

  const headerHeight = Math.max(20, Math.round(DEFAULT_ROW_HEIGHT * zoom / 100))

  const columnMetrics = useMemo(() => {
    const sizes = Array.from({ length: COLS }, (_, col) => {
      if (sheet.hiddenColumns?.[String(col)]) return 0
      return Math.round(getColumnWidth(sheet, col) * zoom / 100)
    })
    return buildTrackMetrics(sizes, 46)
  }, [sheet, zoom])

  const rowMetrics = useMemo(() => {
    const sizes = Array.from({ length: ROWS }, (_, row) => {
      if (hiddenRows.has(row) || sheet.hiddenRows?.[String(row)]) return 0
      return Math.max(20, Math.round(getRowHeight(sheet, row) * zoom / 100))
    })
    return buildTrackMetrics(sizes, headerHeight)
  }, [headerHeight, hiddenRows, sheet, zoom])

  const columnTemplate = useMemo(
    () => `46px ${columnMetrics.sizes.map((size) => `${size}px`).join(' ')}`,
    [columnMetrics],
  )

  const rowTemplate = useMemo(
    () => `${headerHeight}px ${rowMetrics.sizes.map((size) => `${size}px`).join(' ')}`,
    [headerHeight, rowMetrics],
  )

  const visibleColumns = useMemo(
    () => visibleTrackIndexes(columnMetrics, viewport.scrollLeft, viewport.width, 3),
    [columnMetrics, viewport.scrollLeft, viewport.width],
  )

  const visibleRows = useMemo(
    () => visibleTrackIndexes(rowMetrics, viewport.scrollTop, viewport.height, 8),
    [rowMetrics, viewport.height, viewport.scrollTop],
  )

  const gridStyle = {
    '--row-height': `${headerHeight}px`,
    gridTemplateColumns: columnTemplate,
    gridTemplateRows: rowTemplate,
    fontSize: `${Math.max(10, 12 * zoom / 100)}px`,
  } as CSSProperties

  useEffect(() => {
    const node = gridRef.current
    if (!node) return

    const rowTop = rowMetrics.offsets[point.row] ?? headerHeight
    const rowBottom = rowTop + (rowMetrics.sizes[point.row] || headerHeight)
    const colLeft = columnMetrics.offsets[point.col] ?? 46
    const colRight = colLeft + (columnMetrics.sizes[point.col] || DEFAULT_COLUMN_WIDTH)

    const visibleTop = node.scrollTop + headerHeight
    const visibleBottom = node.scrollTop + node.clientHeight - 8
    const visibleLeft = node.scrollLeft + 46
    const visibleRight = node.scrollLeft + node.clientWidth - 8

    if (rowTop < visibleTop) {
      node.scrollTop = Math.max(0, rowTop - headerHeight)
    } else if (rowBottom > visibleBottom) {
      node.scrollTop = Math.max(0, rowBottom - node.clientHeight + 12)
    }

    if (colLeft < visibleLeft) {
      node.scrollLeft = Math.max(0, colLeft - 46)
    } else if (colRight > visibleRight) {
      node.scrollLeft = Math.max(0, colRight - node.clientWidth + 12)
    }
  }, [
    columnMetrics,
    headerHeight,
    point.col,
    point.row,
    rowMetrics,
  ])

  const pageLayout = {
    orientation: 'portrait' as const,
    paperSize: 'letter' as const,
    margins: 'normal' as const,
    printGridlines: true,
    scalePercent: 100,
    headerText: '',
    footerText: '',
    ...(sheet.pageLayout || {}),
  }
  const printBounds = pageLayout.printArea || usedRange()
  const printMargin = pageLayout.margins === 'narrow'
    ? '0.25in'
    : pageLayout.margins === 'wide'
      ? '1in'
      : '0.5in'
  const printPageSize = `${pageLayout.paperSize === 'a4' ? 'A4' : 'letter'} ${pageLayout.orientation}`

  return (
    <div className="app-shell">
      <header className="titlebar">
        <div className="brand">
          <div className="logo-box">X</div>
          <div className="brand-copy">
            <strong>Excel Micro</strong>
            <span>Spreadsheet</span>
          </div>
        </div>

        <input
          className="book-title"
          value={book.title}
          aria-label="Workbook title"
          onChange={(e) => setBook((current) => ({ ...current, title: e.target.value }))}
        />

        <span className={'saved ' + (saveStatus === 'Saved' ? 'is-saved' : '')}>
          <span className="save-dot" />
          {saveStatus}
        </span>

        <div className="title-actions">
          <button onClick={createFreshWorkbook}>New</button>
          <button onClick={() => fileRef.current?.click()}>Open</button>
          <button className="download-btn" onClick={exportXlsx}>Export .xlsx</button>
        </div>
      </header>

      <div className="quickbar">
        <button className="icon-button" title="Undo (Ctrl+Z)" onClick={undo} disabled={!history.length}>↶</button>
        <button className="icon-button" title="Redo (Ctrl+Y)" onClick={redo} disabled={!future.length}>↷</button>
        <span className="divider" />
        <button title="Copy (Ctrl+C)" onClick={() => void copySelected()}>Copy</button>
        <button title="Fill down (Ctrl+D)" onClick={fillDown}>Fill ↓</button>
        <button title="Fill right (Ctrl+R)" onClick={fillRight}>Fill →</button>
        <span className="divider" />
        <button title="Find (Ctrl+F)" onClick={() => setFindOpen(true)}>Find</button>
        <span className="quick-hint">Ctrl+S exports • Double-click a sheet tab to rename</span>
      </div>

      <nav className="tabs">
        {(['Home', 'Insert', 'Formulas', 'Data', 'Page Layout', 'Review', 'View'] as Tab[]).map((item) => (
          <button
            key={item}
            className={tab === item ? 'active' : ''}
            onClick={() => setTab(item)}
          >
            {item}
          </button>
        ))}
      </nav>

      <section className="ribbon">
        {tab === 'Home' && (
          <>
            <Group name="Clipboard">
              <RibbonButton icon="⧉" label="Copy" onClick={() => void copySelected()} />
              <RibbonButton icon="🖌" label="Format Painter" primary={Boolean(formatPainter)} onClick={copyFormat} />
              <RibbonButton icon="↓" label="Fill down" onClick={fillDown} />
              <RibbonButton icon="→" label="Fill right" onClick={fillRight} />
            </Group>

            <Group name="Font">
              <select
                className="font-size-select"
                aria-label="Font size"
                value={cell.format?.fontSize || 11}
                onChange={(e) => applyFormat({ fontSize: Number(e.target.value) })}
              >
                {[9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32].map((size) => (
                  <option key={size} value={size}>{size}</option>
                ))}
              </select>

              <button
                className={cell.format?.bold ? 'on compact' : 'compact'}
                title="Bold (Ctrl+B)"
                onClick={() => applyFormat({ bold: !cell.format?.bold })}
              >
                <b>B</b>
              </button>
              <button
                className={cell.format?.italic ? 'on compact' : 'compact'}
                title="Italic (Ctrl+I)"
                onClick={() => applyFormat({ italic: !cell.format?.italic })}
              >
                <i>I</i>
              </button>
              <button
                className={cell.format?.underline ? 'on compact' : 'compact'}
                title="Underline (Ctrl+U)"
                onClick={() => applyFormat({ underline: !cell.format?.underline })}
              >
                <u>U</u>
              </button>
              <button
                className={cell.format?.strikethrough ? 'on compact' : 'compact'}
                title="Strikethrough"
                onClick={() => applyFormat({ strikethrough: !cell.format?.strikethrough })}
              >
                <s>S</s>
              </button>

              <label className="color-label" title="Text color">
                A
                <input
                  type="color"
                  value={cell.format?.color || '#202020'}
                  onChange={(e) => applyFormat({ color: e.target.value })}
                />
              </label>

              <label className="color-label fill-color" title="Fill color">
                ▰
                <input
                  type="color"
                  value={cell.format?.background || '#ffffff'}
                  onChange={(e) => applyFormat({ background: e.target.value })}
                />
              </label>
            </Group>

            <Group name="Alignment">
              <button className={cell.format?.align === 'left' || !cell.format?.align ? 'on compact wide' : 'compact wide'} onClick={() => applyFormat({ align: 'left' })}>Left</button>
              <button className={cell.format?.align === 'center' ? 'on compact wide' : 'compact wide'} onClick={() => applyFormat({ align: 'center' })}>Center</button>
              <button className={cell.format?.align === 'right' ? 'on compact wide' : 'compact wide'} onClick={() => applyFormat({ align: 'right' })}>Right</button>
              <button className={cell.format?.wrap ? 'on compact wide' : 'compact wide'} onClick={() => applyFormat({ wrap: !cell.format?.wrap })}>Wrap</button>
              <button className="compact wide" onClick={mergeSelection}>Merge</button>
              <button className="compact wide" onClick={unmergeSelection}>Unmerge</button>
            </Group>

            <Group name="Borders">
              <button className="compact" onClick={() => applyBorders('all')}>All</button>
              <button className="compact" onClick={() => applyBorders('outside')}>Outside</button>
              <button className="compact" onClick={() => applyBorders('bottom')}>Bottom</button>
              <button className="compact" onClick={() => applyBorders('none')}>None</button>
              <label className="color-label" title="Border color">
                ▦
                <input
                  type="color"
                  value={cell.format?.borderColor || '#808080'}
                  onChange={(e) => applyFormat({ borderColor: e.target.value })}
                />
              </label>
            </Group>

            <Group name="Number">
              <select
                value={cell.format?.numberFormat || 'general'}
                onChange={(e) => applyFormat({ numberFormat: e.target.value as NumberFormat })}
              >
                <option value="general">General</option>
                <option value="number">Number</option>
                <option value="currency">Currency</option>
                <option value="accounting">Accounting</option>
                <option value="percent">Percent</option>
                <option value="date">Date</option>
                <option value="time">Time</option>
                <option value="datetime">Date & Time</option>
                <option value="scientific">Scientific</option>
                <option value="fraction">Fraction</option>
              </select>
              <button className="compact" title="Decrease decimals" onClick={() => adjustDecimals(-1)}>.0←</button>
              <button className="compact" title="Increase decimals" onClick={() => adjustDecimals(1)}>→.00</button>
            </Group>

            <Group name="Cells">
              <RibbonButton icon="＋R" label="Insert row" onClick={insertRow} />
              <RibbonButton icon="＋C" label="Insert column" onClick={insertColumn} />
              <RibbonButton icon="−R" label="Delete row" onClick={deleteRow} />
              <RibbonButton icon="−C" label="Delete column" onClick={deleteColumn} />
            </Group>

            <Group name="Editing">
              <RibbonButton icon="⌕" label="Find" onClick={() => setFindOpen(true)} />
              <RibbonButton icon="⌫" label="Clear values" onClick={clearSelected} />
              <RibbonButton icon="Tx" label="Clear format" onClick={clearFormatting} />
            </Group>
          </>
        )}

        {tab === 'Insert' && (
          <>
            <Group name="Quick calculation">
              <RibbonButton icon="Σ" label="AutoSum" primary onClick={() => smartFunction('SUM')} />
              <RibbonButton icon="x̄" label="Average" onClick={() => smartFunction('AVERAGE')} />
              <RibbonButton icon="↑" label="Maximum" onClick={() => smartFunction('MAX')} />
              <RibbonButton icon="↓" label="Minimum" onClick={() => smartFunction('MIN')} />
            </Group>

            <Group name="Tables">
              <RibbonButton icon="▦" label="Format as table" primary onClick={openTableDialog} />
            </Group>

            <Group name="Visuals">
              <RibbonButton icon="▥" label="Bar chart" onClick={() => { setChartType('bar'); setChartOpen(true) }} />
              <RibbonButton icon="⌁" label="Line chart" onClick={() => { setChartType('line'); setChartOpen(true) }} />
              <RibbonButton icon="◔" label="Pie chart" onClick={() => { setChartType('pie'); setChartOpen(true) }} />
            </Group>

            <Group name="Links & controls">
              <RibbonButton icon="↗" label="Hyperlink" onClick={() => openHyperlink()} />
              <RibbonButton icon="☑" label="Checkbox" onClick={insertCheckboxes} />
              <RibbonButton icon="×" label="Remove checkbox" onClick={removeCheckboxes} />
            </Group>

            <Group name="Objects">
              <RibbonButton icon="▧" label="Image" onClick={() => imageRef.current?.click()} />
              <RibbonButton icon="T" label="Text box" onClick={insertTextBox} />
              <RibbonButton icon="▭" label="Shape" onClick={insertShape} />
            </Group>

            <Group name="Worksheets">
              <RibbonButton icon="＋" label="New sheet" onClick={addSheet} />
              <RibbonButton icon="⧉" label="Duplicate" onClick={duplicateSheet} />
            </Group>
          </>
        )}

        {tab === 'Formulas' && (
          <>
            <Group name="Functions">
              <RibbonButton icon="Σ" label="SUM" primary onClick={() => smartFunction('SUM')} />
              <RibbonButton icon="x̄" label="AVERAGE" onClick={() => smartFunction('AVERAGE')} />
              <RibbonButton icon="#" label="COUNT" onClick={() => smartFunction('COUNT')} />
              <RibbonButton icon="×" label="PRODUCT" onClick={() => smartFunction('PRODUCT')} />
              <RibbonButton icon="M" label="MEDIAN" onClick={() => smartFunction('MEDIAN')} />
              <RibbonButton icon="↑" label="MAX" onClick={() => smartFunction('MAX')} />
              <RibbonButton icon="↓" label="MIN" onClick={() => smartFunction('MIN')} />
              <RibbonButton icon="?" label="IF" onClick={() => beginEdit('=IF(')} />
              <RibbonButton icon=".0" label="ROUND" onClick={() => beginEdit('=ROUND(')} />
              <RibbonButton icon="#" label="COUNTIF" onClick={() => beginEdit('=COUNTIF(')} />
              <RibbonButton icon="Σ?" label="SUMIF" onClick={() => beginEdit('=SUMIF(')} />
              <RibbonButton icon="X" label="XLOOKUP" onClick={() => beginEdit('=XLOOKUP(')} />
              <RibbonButton icon="!" label="IFERROR" onClick={() => beginEdit('=IFERROR(')} />
            </Group>

            <Group name="Defined names">
              <RibbonButton icon="N" label="Name range" onClick={() => setNamedRangeOpen(true)} />
              <RibbonButton icon="⌖" label="Manage names" onClick={() => setNamedRangeOpen(true)} />
            </Group>

            <div className="formula-help">
              <strong>Formula examples</strong>
              <span>=SUM(A1:A10)</span>
              <span>=AVERAGE(B2:B20)</span>
              <span>=IF(SUM(A1:A5)&gt;100,"Over","OK")</span>
              <span>=ROUND(A1*B1,2)</span>
              <span>=SUM(SALES)</span>
              <span>=COUNTIF(A2:A100,"Paid")</span>
              <span>=XLOOKUP(A2,IDs,Names,"Not found")</span>
            </div>
          </>
        )}

        {tab === 'Data' && (
          <>
            <Group name="Analysis">
              <RibbonButton icon="▤" label="PivotTable" primary onClick={openPivotBuilder} />
            </Group>
            <Group name="Sort & filter">
              <RibbonButton icon="A↓" label="A → Z" onClick={() => sortSelected(1)} />
              <RibbonButton icon="Z↓" label="Z → A" onClick={() => sortSelected(-1)} />
              <RibbonButton icon="▼" label={sheet.filterRange ? 'Remove filter' : 'Filter'} primary={Boolean(sheet.filterRange)} onClick={toggleFilterRange} />
              <RibbonButton icon="×" label="Clear filters" onClick={clearAllFilters} />
            </Group>

            <Group name="Data validation">
              <RibbonButton icon="▼" label="Dropdown" onClick={() => setValidationOpen(true)} />
              <RibbonButton icon="×" label="Clear validation" onClick={clearDataValidation} />
            </Group>

            <Group name="Conditional formatting">
              <RibbonButton icon="▦" label="New rule" onClick={() => setConditionalOpen(true)} />
              <RibbonButton icon="▰" label="Data bars" onClick={() => addConditionalVisual('dataBar')} />
              <RibbonButton icon="◫" label="Color scale" onClick={() => addConditionalVisual('colorScale')} />
              <RibbonButton icon="2×" label="Duplicates" onClick={() => addConditionalVisual('duplicate')} />
              <RibbonButton icon="1×" label="Unique" onClick={() => addConditionalVisual('unique')} />
              <RibbonButton icon="×" label="Clear rules" onClick={clearConditionalFormats} />
            </Group>

            <Group name="Data tools">
              <RibbonButton icon="≠" label="Remove duplicates" onClick={removeDuplicates} />
              <RibbonButton icon="↦" label="Comma to columns" onClick={() => textToColumns(',')} />
              <RibbonButton icon="⇥" label="Tab to columns" onClick={() => textToColumns('\t')} />
              <RibbonButton icon=";" label="Semicolon split" onClick={() => textToColumns(';')} />
            </Group>

            <Group name="Import & export">
              <RibbonButton icon="↑" label="Import" onClick={() => fileRef.current?.click()} />
              <RibbonButton icon="CSV" label="Export CSV" onClick={() => download(safeName(sheet.name) + '.csv', toCsv(sheet), 'text/csv')} />
              <RibbonButton icon="◇" label="Backup" onClick={() => download(safeName(book.title) + '.excelmicro.json', JSON.stringify(book, null, 2), 'application/json')} />
            </Group>

            <div className="formula-help data-tip">
              <strong>File support</strong>
              <span>.xlsx / .xls</span>
              <span>.csv</span>
              <span>.excelmicro.json</span>
            </div>
          </>
        )}

        {tab === 'Page Layout' && (
          <>
            <Group name="Page setup">
              <select
                value={sheet.pageLayout?.orientation || 'portrait'}
                onChange={(e) => updatePageLayout({ orientation: e.target.value as 'portrait' | 'landscape' })}
              >
                <option value="portrait">Portrait</option>
                <option value="landscape">Landscape</option>
              </select>
              <select
                value={sheet.pageLayout?.paperSize || 'letter'}
                onChange={(e) => updatePageLayout({ paperSize: e.target.value as 'letter' | 'a4' })}
              >
                <option value="letter">Letter</option>
                <option value="a4">A4</option>
              </select>
              <select
                value={sheet.pageLayout?.margins || 'normal'}
                onChange={(e) => updatePageLayout({ margins: e.target.value as 'normal' | 'narrow' | 'wide' })}
              >
                <option value="normal">Normal margins</option>
                <option value="narrow">Narrow margins</option>
                <option value="wide">Wide margins</option>
              </select>
            </Group>

            <Group name="Print area">
              <RibbonButton icon="▣" label="Set print area" onClick={setPrintAreaFromSelection} />
              <RibbonButton icon="×" label="Clear print area" onClick={clearPrintArea} />
            </Group>

            <Group name="Scale">
              <select
                value={sheet.pageLayout?.scalePercent || 100}
                onChange={(e) => updatePageLayout({ scalePercent: Number(e.target.value) })}
              >
                {[50, 60, 70, 80, 90, 100, 110, 125, 150, 175, 200].map((value) => (
                  <option key={value} value={value}>{value}%</option>
                ))}
              </select>
            </Group>

            <Group name="Sheet options">
              <button
                className={sheet.pageLayout?.printGridlines === false ? '' : 'on'}
                onClick={() => updatePageLayout({ printGridlines: sheet.pageLayout?.printGridlines === false })}
              >
                Print gridlines
              </button>
              <button onClick={editPrintHeader}>Header</button>
              <button onClick={editPrintFooter}>Footer</button>
            </Group>

            <Group name="Print">
              <RibbonButton icon="🖨" label="Print sheet" primary onClick={printSheet} />
            </Group>
          </>
        )}

        {tab === 'Review' && (
          <>
            <Group name="Notes">
              <RibbonButton icon="▰" label={sheet.notes?.[cellKey(point.row, point.col)] ? 'Edit note' : 'New note'} onClick={() => openNote()} />
              <span className="review-count">{Object.keys(sheet.notes || {}).length} notes</span>
            </Group>

            <Group name="Protection">
              <RibbonButton
                icon={sheet.protected ? '🔒' : '🔓'}
                label={sheet.protected ? 'Unprotect sheet' : 'Protect sheet'}
                primary={Boolean(sheet.protected)}
                onClick={toggleProtection}
              />
            </Group>

            <Group name="Version history">
              <RibbonButton icon="◴" label="Save version" onClick={() => saveVersion()} />
              <RibbonButton icon="↺" label="History" onClick={() => setVersionOpen(true)} />
              <span className="review-count">{versions.length} saved</span>
            </Group>
          </>
        )}

        {tab === 'View' && (
          <>
            <Group name="Sheet view">
              <button
                className={sheet.showGridlines === false ? '' : 'on'}
                onClick={() => toggleSheetView('showGridlines')}
              >
                Gridlines
              </button>
              <button
                className={sheet.freezeTopRow ? 'on' : ''}
                onClick={() => toggleSheetView('freezeTopRow')}
              >
                Freeze top row
              </button>
              <button
                className={sheet.freezeFirstColumn ? 'on' : ''}
                onClick={() => toggleSheetView('freezeFirstColumn')}
              >
                Freeze first column
              </button>
            </Group>

            <Group name="Column width">
              <button className="compact" onClick={() => adjustColumnWidth(-12)}>Narrow</button>
              <button className="compact" onClick={() => adjustColumnWidth(12)}>Widen</button>
              <button className="compact" onClick={resetColumnWidth}>Reset</button>
            </Group>

            <Group name="Row height">
              <button className="compact" onClick={() => adjustRowHeight(-6)}>Shorter</button>
              <button className="compact" onClick={() => adjustRowHeight(6)}>Taller</button>
              <button className="compact" onClick={resetRowHeight}>Reset</button>
            </Group>

            <Group name="Visibility">
              <button className="compact" onClick={hideSelectedRows}>Hide rows</button>
              <button className="compact" onClick={unhideAllRows}>Unhide rows</button>
              <button className="compact" onClick={hideSelectedColumns}>Hide columns</button>
              <button className="compact" onClick={unhideAllColumns}>Unhide columns</button>
              <button className="compact" onClick={hideCurrentSheet}>Hide sheet</button>
              <button className="compact" onClick={unhideAllSheets}>Unhide sheets</button>
            </Group>

            <Group name="Zoom">
              <button className="compact zoom-control" onClick={() => setZoom((z) => Math.max(60, z - 10))}>−</button>
              <span className="zoom-readout">{zoom}%</span>
              <button className="compact zoom-control" onClick={() => setZoom((z) => Math.min(180, z + 10))}>＋</button>
            </Group>
          </>
        )}
      </section>

      <div className="formula-row">
        <input
          className="name-box"
          value={nameBox}
          aria-label="Name box"
          title="Type a cell or range, then press Enter"
          onChange={(e) => setNameBox(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') goToNameBox()
            if (e.key === 'Escape') setNameBox(selectionToAddress(selection))
          }}
        />
        <div className="fx">fx</div>
        <input
          className="formula-input"
          value={editing ? draft : cell.value}
          placeholder="Enter a value or formula"
          onFocus={() => {
            if (!editing) beginEdit()
          }}
          onChange={(e) => setDraft(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') endEdit(true)
            if (e.key === 'Escape') endEdit(false)
          }}
          onBlur={() => {
            if (editing) endEdit(true)
          }}
        />
        <div className="formula-meta">{selectionToAddress(selection)}</div>
      </div>

      <main className="workspace">
        <div
          className={'grid-viewport ' + (sheet.showGridlines === false ? 'hide-gridlines' : '')}
          ref={gridRef}
          tabIndex={0}
          onKeyDown={onGridKey}
          onScroll={(e) => {
            const node = e.currentTarget
            setViewport((current) => ({
              ...current,
              scrollTop: node.scrollTop,
              scrollLeft: node.scrollLeft,
              width: node.clientWidth,
              height: node.clientHeight,
            }))
          }}
        >
          <div className="sheet-grid" style={gridStyle}>
            <div
              className="corner"
              title="Select all"
              onMouseDown={(e) => {
                e.preventDefault()
                setSelection({
                  anchor: { row: 0, col: 0 },
                  focus: { row: ROWS - 1, col: COLS - 1 },
                })
                gridRef.current?.focus()
              }}
            />

            {visibleColumns.map((col) => (
              <div
                key={'h' + col}
                className={
                  'col-head ' +
                  (col >= range.left && col <= range.right && range.top === 0 && range.bottom === ROWS - 1 ? 'header-selected' : '')
                }
                onMouseDown={(e) => {
                  e.preventDefault()
                  setSelection({
                    anchor: { row: 0, col },
                    focus: { row: ROWS - 1, col },
                  })
                  gridRef.current?.focus()
                }}
                onContextMenu={(e) => openContextMenu(e, 'col', col)}
                style={{ gridColumnStart: col + 2 }}
              >
                {colToName(col)}
                <span
                  className="col-resizer"
                  title="Drag to resize • Double-click to AutoFit"
                  onMouseDown={(e) => beginColumnResize(e, col)}
                  onDoubleClick={(e) => {
                    e.preventDefault()
                    e.stopPropagation()
                    autoFitColumn(col)
                  }}
                />
              </div>
            ))}

            {visibleRows.map((row) => (
              <div className={'grid-row ' + (hiddenRows.has(row) || sheet.hiddenRows?.[String(row)] ? 'filtered-out' : '')} key={'r' + row}>
                <div
                  className={
                    'row-head ' +
                    (row >= range.top && row <= range.bottom && range.left === 0 && range.right === COLS - 1 ? 'header-selected' : '')
                  }
                  onMouseDown={(e) => {
                    e.preventDefault()
                    setSelection({
                      anchor: { row, col: 0 },
                      focus: { row, col: COLS - 1 },
                    })
                    gridRef.current?.focus()
                  }}
                  onContextMenu={(e) => openContextMenu(e, 'row', row)}
                  style={{ gridRowStart: row + 2, gridColumnStart: 1 }}
                >
                  {row + 1}
                  <span
                    className="row-resizer"
                    title="Drag to resize • Double-click to AutoFit"
                    onMouseDown={(e) => beginRowResize(e, row)}
                    onDoubleClick={(e) => {
                      e.preventDefault()
                      e.stopPropagation()
                      autoFitRow(row)
                    }}
                  />
                </div>

                {visibleColumns.map((col) => {
                  const data = getCell(sheet, row, col)
                  const active = point.row === row && point.col === col
                  const selected =
                    row >= range.top &&
                    row <= range.bottom &&
                    col >= range.left &&
                    col <= range.right
                  const frozenTop = Boolean(sheet.freezeTopRow && row === 0)
                  const frozenLeft = Boolean(sheet.freezeFirstColumn && col === 0)
                  const frozenClass =
                    frozenTop && frozenLeft
                      ? ' frozen-top frozen-left frozen-corner'
                      : frozenTop
                        ? ' frozen-top'
                        : frozenLeft
                          ? ' frozen-left'
                          : ''

                  const decoration = [
                    data.format?.underline ? 'underline' : '',
                    data.format?.strikethrough ? 'line-through' : '',
                  ].filter(Boolean).join(' ')
                  const display = formatted(data, sheet, book.sheets)
                  const conditionalStyle = conditionalStyleForCell(sheet, row, col, display, conditionalMetrics)
                  const tableStyleInfo = tableStyleForCell(sheet, row, col)
                  const filterHeader = Boolean(
                    sheet.filterRange &&
                    row === sheet.filterRange.top &&
                    col >= sheet.filterRange.left &&
                    col <= sheet.filterRange.right
                  )
                  const filterActive = Boolean(sheet.filters?.[String(col)])
                  const validationRule = dataValidationForCell(sheet, row, col)
                  const note = sheet.notes?.[cellKey(row, col)]
                  const merge = mergeForCell(sheet, row, col)
                  const mergeOrigin = Boolean(merge && merge.top === row && merge.left === col)
                  const mergeCovered = Boolean(merge && !mergeOrigin)
                  const mergeWidth = mergeOrigin && merge
                    ? Array.from({ length: merge.right - merge.left + 1 }, (_, offset) => {
                        const targetCol = merge.left + offset
                        return sheet.hiddenColumns?.[String(targetCol)]
                          ? 0
                          : getColumnWidth(sheet, targetCol) * zoom / 100
                      }).reduce((sum, width) => sum + width, 0)
                    : undefined
                  const mergeHeight = mergeOrigin && merge
                    ? Array.from({ length: merge.bottom - merge.top + 1 }, (_, offset) => {
                        const targetRow = merge.top + offset
                        return hiddenRows.has(targetRow) || sheet.hiddenRows?.[String(targetRow)]
                          ? 0
                          : getRowHeight(sheet, targetRow) * zoom / 100
                      }).reduce((sum, height) => sum + height, 0)
                    : undefined

                  return (
                    <div
                      data-cell={row + ':' + col}
                      key={row + ':' + col}
                      className={
                        'cell ' +
                        (selected ? 'selected ' : '') +
                        (active ? 'active ' : '') +
                        frozenClass +
                        (data.format?.wrap ? ' wrap ' : '') +
                        (validationRule ? ' validation-cell ' : '') +
                        (tableStyleInfo.header ? ' table-header-cell ' : '') +
                        (note ? ' has-note ' : '') +
                        (mergeOrigin ? ' merged-origin ' : '') +
                        (mergeCovered ? ' merged-covered ' : '') +
                        (data.hyperlink ? ' hyperlink-cell ' : '')
                      }
                      style={{
                        fontWeight: data.format?.bold ? 700 : tableStyleInfo.bold ? 700 : 400,
                        fontStyle: data.format?.italic ? 'italic' : 'normal',
                        textDecoration: decoration || 'none',
                        textAlign: data.format?.align || 'left',
                        color: conditionalStyle.color || data.format?.color || tableStyleInfo.color,
                        background: conditionalStyle.background || data.format?.background || tableStyleInfo.background,
                        fontSize: data.format?.fontSize ? `${data.format.fontSize}px` : undefined,
                        gridColumnStart: col + 2,
                        gridRowStart: row + 2,
                        visibility: mergeCovered ? 'hidden' : undefined,
                        width: mergeWidth ? `${mergeWidth}px` : undefined,
                        height: mergeHeight ? `${mergeHeight}px` : undefined,
                        zIndex: mergeOrigin ? 18 : undefined,
                        borderTop: data.format?.borderTop ? `1px solid ${data.format.borderColor || '#808080'}` : undefined,
                        borderRight: data.format?.borderRight ? `1px solid ${data.format.borderColor || '#808080'}` : undefined,
                        borderBottom: data.format?.borderBottom ? `1px solid ${data.format.borderColor || '#808080'}` : undefined,
                        borderLeft: data.format?.borderLeft ? `1px solid ${data.format.borderColor || '#808080'}` : undefined,
                      }}
                      onMouseDown={(e) => {
                        e.preventDefault()
                        if (editing) endEdit(true)
                        const p = { row, col }

                        if (validationRule?.inputMessage) {
                          setNotice((validationRule.inputTitle ? validationRule.inputTitle + ': ' : '') + validationRule.inputMessage)
                        }

                        if (formatPainter && !sheet.protected) {
                          const copiedFormat = structuredClone(formatPainter)
                          mutate((next) => {
                            const s = activeSheet(next)
                            const key = cellKey(row, col)
                            const old = s.cells[key] || { value: '' }
                            s.cells[key] = { ...old, format: copiedFormat }
                          })
                          setFormatPainter(null)
                          setNotice('Format applied')
                        }

                        setDragging(true)
                        setSelection((current) => ({
                          anchor: e.shiftKey ? current.anchor : p,
                          focus: p,
                        }))
                        gridRef.current?.focus()
                      }}
                      onMouseEnter={() => {
                        if (dragging) {
                          setSelection((current) => ({
                            ...current,
                            focus: { row, col },
                          }))
                        }
                      }}
                      onDoubleClick={() => beginEdit()}
                      onContextMenu={(e) => openContextMenu(e, 'cell', row * COLS + col, { row, col })}
                    >
                      {conditionalStyle.dataBarPercent !== undefined && (
                        <span
                          className="conditional-data-bar"
                          style={{
                            width: `${Math.max(0, Math.min(100, conditionalStyle.dataBarPercent))}%`,
                            background: conditionalStyle.dataBarColor,
                          }}
                        />
                      )}
                      {active && editing ? (
                        <input
                          className="cell-input"
                          autoFocus
                          value={draft}
                          onChange={(e) => setDraft(e.target.value)}
                          onBlur={() => endEdit(true)}
                          onKeyDown={(e) => {
                            if (e.key === 'Enter') {
                              e.preventDefault()
                              endEdit(true)
                              move(1, 0)
                            }
                            if (e.key === 'Escape') endEdit(false)
                          }}
                        />
                      ) : data.checkbox ? (
                        <label
                          className="cell-checkbox"
                          title="Checkbox"
                          onMouseDown={(e) => {
                            e.stopPropagation()
                          }}
                        >
                          <input
                            type="checkbox"
                            checked={data.value.toUpperCase() === 'TRUE'}
                            disabled={Boolean(sheet.protected)}
                            onChange={(e) => {
                              e.stopPropagation()
                              setValue({ row, col }, e.target.checked ? 'TRUE' : 'FALSE')
                            }}
                          />
                        </label>
                      ) : data.hyperlink ? (
                        <a
                          className="cell-hyperlink"
                          href={data.hyperlink}
                          target="_blank"
                          rel="noreferrer"
                          onMouseDown={(e) => e.stopPropagation()}
                          onClick={(e) => e.stopPropagation()}
                        >
                          {display || data.hyperlink}
                        </a>
                      ) : (
                        <span>{display}</span>
                      )}
                      {note && (
                        <button
                          className="cell-note-marker"
                          title={note.text}
                          onMouseDown={(e) => {
                            e.preventDefault()
                            e.stopPropagation()
                          }}
                          onClick={(e) => {
                            e.preventDefault()
                            e.stopPropagation()
                            openNote({ row, col })
                          }}
                        />
                      )}
                      {validationRule && (validationRule.type || 'list') === 'list' && !editing && (
                        <button
                          className={'cell-validation-button ' + (filterHeader ? 'with-filter' : '')}
                          title="Choose an allowed value"
                          onMouseDown={(e) => {
                            e.preventDefault()
                            e.stopPropagation()
                          }}
                          onClick={(e) => {
                            e.preventDefault()
                            e.stopPropagation()
                            const rect = e.currentTarget.getBoundingClientRect()
                            setValidationPicker({
                              row,
                              col,
                              x: Math.min(window.innerWidth - 190, rect.right - 170),
                              y: Math.min(window.innerHeight - 220, rect.bottom + 3),
                              ruleId: validationRule.id,
                            })
                          }}
                        >
                          ▼
                        </button>
                      )}
                      {filterHeader && (
                        <button
                          className={'cell-filter-button ' + (filterActive ? 'is-active' : '')}
                          title={filterActive ? 'Filter active' : 'Filter this column'}
                          onMouseDown={(e) => {
                            e.preventDefault()
                            e.stopPropagation()
                          }}
                          onClick={(e) => {
                            e.preventDefault()
                            e.stopPropagation()
                            openFilterEditor(col)
                          }}
                        >
                          {filterActive ? '●' : '▼'}
                        </button>
                      )}
                    </div>
                  )
                })}
              </div>
            ))}

            {(sheet.objects || []).map((object) => (
              <div
                className={'sheet-object sheet-object-' + object.type}
                key={object.id}
                style={{
                  gridColumnStart: object.col + 2,
                  gridRowStart: object.row + 2,
                  width: `${Math.max(40, object.width * zoom / 100)}px`,
                  height: `${Math.max(30, object.height * zoom / 100)}px`,
                  background: object.fill,
                  borderColor: object.border,
                }}
                onMouseDown={(e) => e.stopPropagation()}
                onDoubleClick={() => editSheetObject(object.id)}
                title={object.type === 'image' ? 'Image' : 'Double-click to edit'}
              >
                {object.type === 'image' && object.src ? (
                  <img src={object.src} alt="" />
                ) : object.type === 'chart' && object.chartRange ? (
                  <MiniChart
                    data={chartDataForRange(sheet, object.chartRange, book.sheets)}
                    type={object.chartType || 'bar'}
                  />
                ) : (
                  <div className="sheet-object-text">{object.text}</div>
                )}
                {!sheet.protected && (
                  <button
                    className="sheet-object-delete"
                    title="Delete object"
                    onClick={(e) => {
                      e.preventDefault()
                      e.stopPropagation()
                      removeSheetObject(object.id)
                    }}
                  >
                    ×
                  </button>
                )}
              </div>
            ))}
          </div>
        </div>
      </main>

      <div className="sheetbar">
        <button className="plus" title="New sheet" onClick={addSheet}>＋</button>

        <div className="sheet-tabs">
          {book.sheets.filter((item) => !item.hidden).map((item) => (
            <button
              key={item.id}
              className={item.id === book.activeSheetId ? 'active' : ''}
              onClick={() => {
                setBook((current) => ({ ...current, activeSheetId: item.id }))
                selectPoint({ row: 0, col: 0 })
              }}
              onDoubleClick={() => renameSheet(item)}
              title="Double-click to rename"
            >
              {item.name}
            </button>
          ))}
        </div>

        <div className="sheet-actions">
          <button onClick={() => renameSheet(sheet)}>Rename</button>
          <button onClick={duplicateSheet}>Duplicate</button>
          <button className="danger" onClick={deleteSheet} disabled={book.sheets.length === 1}>Delete</button>
        </div>
      </div>

      <footer className="status">
        <span className="ready-pill">Ready</span>
        <span>{selectionToAddress(selection)}</span>
        <span className="spacer" />
        {stats.count > 0 && (
          <>
            <span>Average: {stats.avg.toLocaleString(undefined, { maximumFractionDigits: 4 })}</span>
            <span>Count: {stats.count}</span>
            <span>Sum: {stats.sum.toLocaleString(undefined, { maximumFractionDigits: 4 })}</span>
          </>
        )}
        <span className="status-view">{sheet.protected ? '🔒 Protected' : sheet.freezeTopRow || sheet.freezeFirstColumn ? 'Frozen panes' : 'Normal view'}</span>
        <span>{zoom}%</span>
      </footer>

      <style>{`@media print { @page { size: ${printPageSize}; margin: ${printMargin}; } }`}</style>
      <div
        className={'print-area ' + (pageLayout.printGridlines ? 'print-gridlines' : '')}
        style={{ fontSize: `${10 * ((pageLayout.scalePercent || 100) / 100)}px` }}
      >
        <div className="print-header-text">{pageLayout.headerText || book.title}</div>
        <div className="print-title">{book.title} — {sheet.name}</div>
        <table>
          <tbody>
            {Array.from({ length: printBounds.bottom - printBounds.top + 1 }, (_, rowOffset) => {
              const row = printBounds.top + rowOffset
              return (
                <tr key={row}>
                  {Array.from({ length: printBounds.right - printBounds.left + 1 }, (_, colOffset) => {
                    const col = printBounds.left + colOffset
                    const data = getCell(sheet, row, col)
                    const tableStyleInfo = tableStyleForCell(sheet, row, col)
                    const conditionalStyle = conditionalStyleForCell(sheet, row, col, formatted(data, sheet, book.sheets), conditionalMetrics)
                    return (
                      <td
                        key={col}
                        style={{
                          fontWeight: data.format?.bold || tableStyleInfo.bold ? 700 : 400,
                          fontStyle: data.format?.italic ? 'italic' : 'normal',
                          textAlign: data.format?.align || 'left',
                          color: conditionalStyle.color || data.format?.color || tableStyleInfo.color,
                          background: conditionalStyle.background || data.format?.background || tableStyleInfo.background,
                        }}
                      >
                        {formatted(data, sheet, book.sheets)}
                      </td>
                    )
                  })}
                </tr>
              )
            })}
          </tbody>
        </table>
        <div className="print-footer-text">{pageLayout.footerText || sheet.name}</div>
      </div>

      {findOpen && (
        <div className="find-box find-replace-box">
          <div className="find-title">
            <strong>Find & Replace</strong>
            <button onClick={() => setFindOpen(false)}>×</button>
          </div>
          <input
            autoFocus
            placeholder="Find in this sheet"
            value={find}
            onChange={(e) => setFind(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') findNext()
              if (e.key === 'Escape') setFindOpen(false)
            }}
          />
          <button className="find-next" onClick={findNext}>Find next</button>
          <input
            placeholder="Replace with"
            value={replaceText}
            onChange={(e) => setReplaceText(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') replaceMatch(false)
              if (e.key === 'Escape') setFindOpen(false)
            }}
          />
          <div className="replace-actions">
            <button onClick={() => replaceMatch(false)}>Replace</button>
            <button onClick={() => replaceMatch(true)}>Replace all</button>
          </div>
        </div>
      )}

      {chartOpen && (
        <div className="overlay" onMouseDown={() => setChartOpen(false)}>
          <div className="chart-card" onMouseDown={(e) => e.stopPropagation()}>
            <div className="chart-title">
              <div>
                <span className="eyebrow">INSERTED FROM {selectionToAddress(selection)}</span>
                <h2>Quick {chartType} chart</h2>
                <p>Select one numeric column, or labels + values in two columns.</p>
              </div>
              <button onClick={() => setChartOpen(false)}>×</button>
            </div>
            <div className="chart-switcher">
              {(['bar', 'line', 'pie'] as ChartType[]).map((type) => (
                <button
                  key={type}
                  className={chartType === type ? 'active' : ''}
                  onClick={() => setChartType(type)}
                >
                  {type === 'bar' ? '▥ Bar' : type === 'line' ? '⌁ Line' : '◔ Pie'}
                </button>
              ))}
              <span className="chart-switcher-spacer" />
              <button className="chart-insert-button" onClick={insertCurrentChart}>＋ Add to worksheet</button>
            </div>
            <Chart data={chart} type={chartType} />
          </div>
        </div>
      )}

      {tableOpen && (
        <div className="overlay panel-overlay" onMouseDown={() => setTableOpen(false)}>
          <div className="rule-card table-card" onMouseDown={(e) => e.stopPropagation()}>
            <div className="rule-card-header">
              <div>
                <span className="eyebrow">FORMAT AS TABLE</span>
                <h2>Create table</h2>
                <p>Turn {selectionToAddress(selection)} into a structured, filterable data table.</p>
              </div>
              <button onClick={() => setTableOpen(false)}>×</button>
            </div>

            <div className="table-form">
              <label>
                Table name
                <input value={tableName} onChange={(e) => setTableName(e.target.value)} />
              </label>

              <div className="table-style-field">
                <span>Style</span>
                <div className="table-style-grid">
                  {(Object.keys(TABLE_PALETTES) as TableStyle[]).map((style) => (
                    <button
                      key={style}
                      className={'table-style-option ' + (tableStyle === style ? 'active' : '')}
                      onClick={() => setTableStyle(style)}
                    >
                      <span style={{ background: TABLE_PALETTES[style].header }} />
                      <span style={{ background: TABLE_PALETTES[style].band }} />
                      <strong>{style}</strong>
                    </button>
                  ))}
                </div>
              </div>

              <label className="validation-checkbox">
                <input
                  type="checkbox"
                  checked={tableBandedRows}
                  onChange={(e) => setTableBandedRows(e.target.checked)}
                />
                Banded rows
              </label>
            </div>

            {(sheet.tables || []).length > 0 && (
              <div className="table-list">
                <strong>Tables on this sheet</strong>
                {(sheet.tables || []).map((table) => (
                  <div className="table-list-item" key={table.id}>
                    <span className="table-list-swatch" style={{ background: TABLE_PALETTES[table.style].header }} />
                    <button
                      onClick={() => {
                        setSelection({
                          anchor: { row: table.top, col: table.left },
                          focus: { row: table.bottom, col: table.right },
                        })
                        setTableOpen(false)
                      }}
                    >
                      {table.name}
                    </button>
                    <span>
                      {selectionToAddress({
                        anchor: { row: table.top, col: table.left },
                        focus: { row: table.bottom, col: table.right },
                      })}
                    </span>
                    <button className="remove-table" onClick={() => removeTable(table.id)}>Remove</button>
                  </div>
                ))}
              </div>
            )}

            <div className="rule-actions">
              <span className="named-range-tip">The first row becomes the table header and gets filter controls.</span>
              <span className="spacer" />
              <button className="secondary" onClick={() => setTableOpen(false)}>Cancel</button>
              <button className="primary-action" onClick={addTable}>Create table</button>
            </div>
          </div>
        </div>
      )}

      {validationOpen && (
        <div className="overlay panel-overlay" onMouseDown={() => setValidationOpen(false)}>
          <div className="rule-card validation-card" onMouseDown={(e) => e.stopPropagation()}>
            <div className="rule-card-header">
              <div>
                <span className="eyebrow">DATA VALIDATION</span>
                <h2>Validation rule</h2>
                <p>Control what can be entered in {selectionToAddress(selection)}.</p>
              </div>
              <button onClick={() => setValidationOpen(false)}>×</button>
            </div>

            <div className="validation-form advanced-validation-form">
              <label>
                Allow
                <select
                  value={validationType}
                  onChange={(e) => setValidationType(e.target.value as DataValidationType)}
                >
                  <option value="list">List / dropdown</option>
                  <option value="whole">Whole number</option>
                  <option value="decimal">Decimal</option>
                  <option value="date">Date</option>
                  <option value="textLength">Text length</option>
                </select>
              </label>

              {validationType === 'list' ? (
                <label className="validation-wide">
                  Allowed values
                  <textarea
                    autoFocus
                    value={validationOptions}
                    onChange={(e) => setValidationOptions(e.target.value)}
                    placeholder={'Pending\nPaid\nCanceled'}
                  />
                  <span>Enter one value per line, or separate values with commas.</span>
                </label>
              ) : (
                <>
                  <label>
                    Condition
                    <select
                      value={validationOperator}
                      onChange={(e) => setValidationOperator(e.target.value as DataValidationOperator)}
                    >
                      <option value="between">Between</option>
                      <option value="equalTo">Equal to</option>
                      <option value="greaterThan">Greater than</option>
                      <option value="lessThan">Less than</option>
                    </select>
                  </label>

                  <label>
                    {validationOperator === 'between' ? 'Minimum' : 'Value'}
                    <input
                      type={validationType === 'date' ? 'date' : 'number'}
                      value={validationMinimum}
                      onChange={(e) => setValidationMinimum(e.target.value)}
                    />
                  </label>

                  {validationOperator === 'between' && (
                    <label>
                      Maximum
                      <input
                        type={validationType === 'date' ? 'date' : 'number'}
                        value={validationMaximum}
                        onChange={(e) => setValidationMaximum(e.target.value)}
                      />
                    </label>
                  )}
                </>
              )}

              <label>
                Input title
                <input
                  value={validationInputTitle}
                  onChange={(e) => setValidationInputTitle(e.target.value)}
                  placeholder="Optional"
                />
              </label>

              <label className="validation-wide">
                Input message
                <input
                  value={validationInputMessage}
                  onChange={(e) => setValidationInputMessage(e.target.value)}
                  placeholder="Shown when a validated cell is selected"
                />
              </label>

              <label>
                Invalid data
                <select
                  value={validationErrorStyle}
                  onChange={(e) => setValidationErrorStyle(e.target.value as DataValidationErrorStyle)}
                >
                  <option value="stop">Stop — reject value</option>
                  <option value="warning">Warning — ask before keeping</option>
                  <option value="information">Information — allow with notice</option>
                </select>
              </label>

              <label className="validation-checkbox">
                <input
                  type="checkbox"
                  checked={validationAllowBlank}
                  onChange={(e) => setValidationAllowBlank(e.target.checked)}
                />
                Allow blank cells
              </label>
            </div>

            <div className="rule-actions">
              <button className="secondary" onClick={clearDataValidation}>Clear validation from selection</button>
              <span className="spacer" />
              <button className="secondary" onClick={() => setValidationOpen(false)}>Cancel</button>
              <button className="primary-action" onClick={addDataValidation}>Apply validation</button>
            </div>
          </div>
        </div>
      )}

      {validationPicker && (() => {
        const rule = (sheet.dataValidations || []).find((item) => item.id === validationPicker.ruleId)
        if (!rule) return null
        return (
          <div
            className="validation-picker"
            style={{ left: validationPicker.x, top: validationPicker.y }}
            onMouseDown={(e) => e.stopPropagation()}
          >
            {rule.allowBlank && (
              <button onClick={() => chooseValidationValue('')}>
                <span className="validation-blank">Blank</span>
              </button>
            )}
            {(rule.options || []).map((option) => (
              <button key={option} onClick={() => chooseValidationValue(option)}>
                {option}
              </button>
            ))}
          </div>
        )
      })()}

      {filterEditor && (
        <div className="overlay panel-overlay" onMouseDown={() => setFilterEditor(null)}>
          <div className="rule-card filter-card" onMouseDown={(e) => e.stopPropagation()}>
            <div className="rule-card-header">
              <div>
                <span className="eyebrow">FILTER</span>
                <h2>Column {colToName(filterEditor.col)}</h2>
                <p>Show rows that match this condition.</p>
              </div>
              <button onClick={() => setFilterEditor(null)}>×</button>
            </div>
            <div className="rule-form">
              <label>
                Condition
                <select
                  value={filterEditor.operator}
                  onChange={(e) => setFilterEditor((current) => current && ({ ...current, operator: e.target.value as FilterOperator }))}
                >
                  <option value="contains">Contains</option>
                  <option value="equals">Equals</option>
                  <option value="greaterThan">Greater than</option>
                  <option value="lessThan">Less than</option>
                  <option value="notBlank">Is not blank</option>
                </select>
              </label>
              {filterEditor.operator !== 'notBlank' && (
                <label>
                  Value
                  <input
                    autoFocus
                    value={filterEditor.value}
                    onChange={(e) => setFilterEditor((current) => current && ({ ...current, value: e.target.value }))}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') applyFilterEditor()
                      if (e.key === 'Escape') setFilterEditor(null)
                    }}
                  />
                </label>
              )}
            </div>
            <div className="rule-actions">
              <button className="secondary" onClick={() => clearFilterColumn(filterEditor.col)}>Clear column filter</button>
              <span className="spacer" />
              <button className="secondary" onClick={() => setFilterEditor(null)}>Cancel</button>
              <button className="primary-action" onClick={applyFilterEditor}>Apply filter</button>
            </div>
          </div>
        </div>
      )}

      {namedRangeOpen && (
        <div className="overlay panel-overlay" onMouseDown={() => setNamedRangeOpen(false)}>
          <div className="rule-card" onMouseDown={(e) => e.stopPropagation()}>
            <div className="rule-card-header">
              <div>
                <span className="eyebrow">DEFINED NAMES</span>
                <h2>Named ranges</h2>
                <p>Name {selectionToAddress(selection)} so formulas can reference it directly.</p>
              </div>
              <button onClick={() => setNamedRangeOpen(false)}>×</button>
            </div>

            <div className="named-range-create">
              <label>
                Name
                <input
                  autoFocus
                  placeholder="e.g. SALES"
                  value={namedRangeName}
                  onChange={(e) => setNamedRangeName(e.target.value.toUpperCase())}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') createNamedRange()
                    if (e.key === 'Escape') setNamedRangeOpen(false)
                  }}
                />
              </label>
              <div className="named-range-target">
                <span>Refers to</span>
                <strong>{selectionToAddress(selection)}</strong>
              </div>
              <button className="primary-action" onClick={createNamedRange}>Create</button>
            </div>

            <div className="named-range-list">
              {Object.keys(sheet.namedRanges || {}).length === 0 ? (
                <div className="empty-rules">No named ranges yet.</div>
              ) : (
                Object.entries(sheet.namedRanges || {})
                  .sort(([a], [b]) => a.localeCompare(b))
                  .map(([name, named]) => (
                    <div className="named-range-item" key={name}>
                      <button
                        className="named-range-name"
                        onClick={() => {
                          setSelection({
                            anchor: { row: named.top, col: named.left },
                            focus: { row: named.bottom, col: named.right },
                          })
                          setNamedRangeOpen(false)
                        }}
                      >
                        {name}
                      </button>
                      <span>
                        {selectionToAddress({
                          anchor: { row: named.top, col: named.left },
                          focus: { row: named.bottom, col: named.right },
                        })}
                      </span>
                      <button className="remove-name" onClick={() => removeNamedRange(name)}>Remove</button>
                    </div>
                  ))
              )}
            </div>

            <div className="rule-actions">
              <span className="named-range-tip">Use names in formulas, for example =SUM(SALES)</span>
              <span className="spacer" />
              <button className="secondary" onClick={() => setNamedRangeOpen(false)}>Done</button>
            </div>
          </div>
        </div>
      )}

      {conditionalOpen && (
        <div className="overlay panel-overlay" onMouseDown={() => setConditionalOpen(false)}>
          <div className="rule-card" onMouseDown={(e) => e.stopPropagation()}>
            <div className="rule-card-header">
              <div>
                <span className="eyebrow">CONDITIONAL FORMATTING</span>
                <h2>New rule</h2>
                <p>Apply formatting to {selectionToAddress(selection)} when the condition is true.</p>
              </div>
              <button onClick={() => setConditionalOpen(false)}>×</button>
            </div>

            <div className="rule-form rule-grid">
              <label>
                Condition
                <select value={conditionalOperator} onChange={(e) => setConditionalOperator(e.target.value as ConditionalFormatOperator)}>
                  <option value="greaterThan">Greater than</option>
                  <option value="lessThan">Less than</option>
                  <option value="equals">Equals</option>
                  <option value="contains">Contains text</option>
                  <option value="notBlank">Is not blank</option>
                </select>
              </label>

              {conditionalOperator !== 'notBlank' && (
                <label>
                  Value
                  <input value={conditionalValue} onChange={(e) => setConditionalValue(e.target.value)} />
                </label>
              )}

              <label>
                Fill
                <input type="color" value={conditionalBackground} onChange={(e) => setConditionalBackground(e.target.value)} />
              </label>

              <label>
                Text
                <input type="color" value={conditionalColor} onChange={(e) => setConditionalColor(e.target.value)} />
              </label>
            </div>

            {(sheet.conditionalFormats || []).length > 0 && (
              <div className="rule-list">
                <strong>Existing rules</strong>
                {(sheet.conditionalFormats || []).map((rule) => (
                  <div className="rule-list-item" key={rule.id}>
                    <span className="rule-swatch" style={{ background: rule.background, color: rule.color }}>Aa</span>
                    <span>
                      {selectionToAddress({
                        anchor: { row: rule.top, col: rule.left },
                        focus: { row: rule.bottom, col: rule.right },
                      })} · {rule.operator}{rule.value ? ` ${rule.value}` : ''}
                    </span>
                    <button onClick={() => removeConditionalFormat(rule.id)}>Remove</button>
                  </div>
                ))}
              </div>
            )}

            <div className="rule-actions">
              <button className="secondary" onClick={clearConditionalFormats}>Clear all rules</button>
              <span className="spacer" />
              <button className="secondary" onClick={() => setConditionalOpen(false)}>Cancel</button>
              <button className="primary-action" onClick={addConditionalFormat}>Add rule</button>
            </div>
          </div>
        </div>
      )}

      {versionOpen && (
        <div className="overlay panel-overlay" onMouseDown={() => setVersionOpen(false)}>
          <div className="rule-card version-card" onMouseDown={(e) => e.stopPropagation()}>
            <div className="rule-card-header">
              <div>
                <span className="eyebrow">VERSION HISTORY</span>
                <h2>Workbook versions</h2>
                <p>Restore a local snapshot of the entire workbook.</p>
              </div>
              <button onClick={() => setVersionOpen(false)}>×</button>
            </div>

            <div className="version-toolbar">
              <button className="primary-action" onClick={() => saveVersion()}>Save current version</button>
              <span>Up to 12 versions are kept in this browser.</span>
            </div>

            <div className="version-list">
              {versions.length === 0 ? (
                <div className="empty-rules">No saved versions yet.</div>
              ) : versions.map((version) => (
                <div className="version-item" key={version.id}>
                  <div>
                    <strong>{version.label}</strong>
                    <span>{new Date(version.savedAt).toLocaleString()}</span>
                    <small>{version.workbook.sheets.length} sheet{version.workbook.sheets.length === 1 ? '' : 's'} · {version.workbook.title}</small>
                  </div>
                  <button onClick={() => restoreVersion(version)}>Restore</button>
                  <button className="remove-version" onClick={() => deleteVersion(version.id)}>Delete</button>
                </div>
              ))}
            </div>

            <div className="rule-actions">
              <span className="named-range-tip">Versions are stored locally in this browser.</span>
              <span className="spacer" />
              <button className="secondary" onClick={() => setVersionOpen(false)}>Done</button>
            </div>
          </div>
        </div>
      )}

      {noteOpen && (
        <div className="overlay panel-overlay" onMouseDown={() => setNoteOpen(false)}>
          <div className="rule-card note-card" onMouseDown={(e) => e.stopPropagation()}>
            <div className="rule-card-header">
              <div>
                <span className="eyebrow">CELL NOTE</span>
                <h2>{pointToAddress(noteTarget)}</h2>
                <p>Add a note that stays attached to this cell.</p>
              </div>
              <button onClick={() => setNoteOpen(false)}>×</button>
            </div>
            <div className="note-editor">
              <textarea
                autoFocus
                value={noteDraft}
                onChange={(e) => setNoteDraft(e.target.value)}
                placeholder="Write a note…"
                onKeyDown={(e) => {
                  if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') saveNote()
                  if (e.key === 'Escape') setNoteOpen(false)
                }}
              />
              <span>Ctrl/Cmd + Enter to save</span>
            </div>
            <div className="rule-actions">
              {sheet.notes?.[cellKey(noteTarget.row, noteTarget.col)] && (
                <button className="secondary danger-text" onClick={removeNote}>Delete note</button>
              )}
              <span className="spacer" />
              <button className="secondary" onClick={() => setNoteOpen(false)}>Cancel</button>
              <button className="primary-action" onClick={saveNote}>Save note</button>
            </div>
          </div>
        </div>
      )}

      {pivotOpen && (
        <div className="overlay panel-overlay" onMouseDown={() => setPivotOpen(false)}>
          <div className="rule-card pivot-card" onMouseDown={(e) => e.stopPropagation()}>
            <div className="rule-card-header">
              <div>
                <span className="eyebrow">PIVOTTABLE</span>
                <h2>Create summary</h2>
                <p>Summarize {selectionToAddress({
                  anchor: { row: pivotSource.top, col: pivotSource.left },
                  focus: { row: pivotSource.bottom, col: pivotSource.right },
                })} on a new sheet.</p>
              </div>
              <button onClick={() => setPivotOpen(false)}>×</button>
            </div>

            <div className="pivot-form">
              <label>
                Rows
                <select value={pivotRowField} onChange={(e) => setPivotRowField(Number(e.target.value))}>
                  {Array.from({ length: pivotSource.right - pivotSource.left + 1 }, (_, offset) => {
                    const col = pivotSource.left + offset
                    return (
                      <option key={col} value={col}>
                        {formatted(getCell(sheet, pivotSource.top, col), sheet, book.sheets) || colToName(col)}
                      </option>
                    )
                  })}
                </select>
              </label>

              <label>
                Values
                <select value={pivotValueField} onChange={(e) => setPivotValueField(Number(e.target.value))}>
                  {Array.from({ length: pivotSource.right - pivotSource.left + 1 }, (_, offset) => {
                    const col = pivotSource.left + offset
                    return (
                      <option key={col} value={col}>
                        {formatted(getCell(sheet, pivotSource.top, col), sheet, book.sheets) || colToName(col)}
                      </option>
                    )
                  })}
                </select>
              </label>

              <label>
                Calculation
                <select value={pivotAggregator} onChange={(e) => setPivotAggregator(e.target.value as 'sum' | 'count' | 'average')}>
                  <option value="sum">Sum</option>
                  <option value="count">Count</option>
                  <option value="average">Average</option>
                </select>
              </label>
            </div>

            <div className="rule-actions">
              <span className="named-range-tip">Excel Micro creates the summary on a new worksheet.</span>
              <span className="spacer" />
              <button className="secondary" onClick={() => setPivotOpen(false)}>Cancel</button>
              <button className="primary-action" onClick={createPivot}>Create PivotTable</button>
            </div>
          </div>
        </div>
      )}

      {hyperlinkOpen && (
        <div className="overlay panel-overlay" onMouseDown={() => setHyperlinkOpen(false)}>
          <div className="rule-card hyperlink-card" onMouseDown={(e) => e.stopPropagation()}>
            <div className="rule-card-header">
              <div>
                <span className="eyebrow">HYPERLINK</span>
                <h2>Link {pointToAddress(hyperlinkTarget)}</h2>
                <p>Add a web address to the selected cell.</p>
              </div>
              <button onClick={() => setHyperlinkOpen(false)}>×</button>
            </div>
            <div className="rule-form">
              <label>
                Address
                <input
                  autoFocus
                  value={hyperlinkDraft}
                  onChange={(e) => setHyperlinkDraft(e.target.value)}
                  placeholder="https://example.com"
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') saveHyperlink()
                    if (e.key === 'Escape') setHyperlinkOpen(false)
                  }}
                />
              </label>
            </div>
            <div className="rule-actions">
              <button className="secondary" onClick={() => setHyperlinkDraft('')}>Remove link</button>
              <span className="spacer" />
              <button className="secondary" onClick={() => setHyperlinkOpen(false)}>Cancel</button>
              <button className="primary-action" onClick={saveHyperlink}>Save link</button>
            </div>
          </div>
        </div>
      )}

      {contextMenu && (
        <div
          className="context-menu"
          style={{ left: contextMenu.x, top: contextMenu.y }}
          onMouseDown={(e) => e.stopPropagation()}
        >
          <div className="context-menu-title">
            {contextMenu.kind === 'row'
              ? `Row ${contextMenu.index + 1}`
              : contextMenu.kind === 'col'
                ? `Column ${colToName(contextMenu.index)}`
                : selectionToAddress(selection)}
          </div>

          {contextMenu.kind === 'cell' && (
            <>
              <button onClick={() => runContextAction(() => void copySelected())}><span>⧉</span>Copy</button>
              <button
                onClick={() => runContextAction(() => openHyperlink({
                  row: Math.floor(contextMenu.index / COLS),
                  col: contextMenu.index % COLS,
                }))}
              >
                <span>↗</span>{getCell(sheet, Math.floor(contextMenu.index / COLS), contextMenu.index % COLS).hyperlink ? 'Edit hyperlink' : 'Add hyperlink'}
              </button>
              <button
                onClick={() => runContextAction(() => openNote({
                  row: Math.floor(contextMenu.index / COLS),
                  col: contextMenu.index % COLS,
                }))}
              >
                <span>▰</span>{sheet.notes?.[cellKey(Math.floor(contextMenu.index / COLS), contextMenu.index % COLS)] ? 'Edit note' : 'Add note'}
              </button>
              <button onClick={() => runContextAction(clearSelected)}><span>⌫</span>Clear contents</button>
              <button onClick={() => runContextAction(clearFormatting)}><span>Tx</span>Clear formatting</button>
              <div className="context-divider" />
              <button onClick={() => runContextAction(insertRow)}><span>＋R</span>Insert row above</button>
              <button onClick={() => runContextAction(insertColumn)}><span>＋C</span>Insert column left</button>
            </>
          )}

          {contextMenu.kind === 'row' && (
            <>
              <button onClick={() => runContextAction(insertRow)}><span>＋</span>Insert row above</button>
              <button className="context-danger" onClick={() => runContextAction(deleteRow)}><span>−</span>Delete row</button>
              <div className="context-divider" />
              <button onClick={() => runContextAction(() => autoFitRow(contextMenu.index))}><span>↕</span>AutoFit row height</button>
              <button onClick={() => runContextAction(() => adjustRowHeight(6))}><span>↕</span>Increase row height</button>
            </>
          )}

          {contextMenu.kind === 'col' && (
            <>
              <button onClick={() => runContextAction(insertColumn)}><span>＋</span>Insert column left</button>
              <button className="context-danger" onClick={() => runContextAction(deleteColumn)}><span>−</span>Delete column</button>
              <div className="context-divider" />
              <button onClick={() => runContextAction(() => autoFitColumn(contextMenu.index))}><span>↔</span>AutoFit column width</button>
              <button onClick={() => runContextAction(() => adjustColumnWidth(12))}><span>↔</span>Increase column width</button>
            </>
          )}
        </div>
      )}

      {notice && <div className="toast">{notice}</div>}

      <input
        ref={imageRef}
        className="hidden"
        type="file"
        accept="image/png,image/jpeg,image/webp,image/gif,image/svg+xml"
        onChange={(e) => {
          const file = e.target.files?.[0]
          e.target.value = ''
          if (file) handleImageFile(file)
        }}
      />

      <input
        ref={fileRef}
        className="hidden"
        type="file"
        accept=".xlsx,.xls,.csv,.json"
        onChange={(e) => {
          const file = e.target.files?.[0]
          e.target.value = ''
          if (file) {
            void importFile(file).catch((error) => {
              window.alert(error instanceof Error ? error.message : 'Import failed')
            })
          }
        }}
      />
    </div>
  )
}

function Group({ name, children }: { name: string; children: ReactNode }) {
  return (
    <div className="group">
      <div className="group-controls">{children}</div>
      <small>{name}</small>
    </div>
  )
}

function RibbonButton({
  icon,
  label,
  onClick,
  primary = false,
}: {
  icon: string
  label: string
  onClick: () => void
  primary?: boolean
}) {
  return (
    <button className={'ribbon-button ' + (primary ? 'primary' : '')} onClick={onClick}>
      <span className="ribbon-icon">{icon}</span>
      <span>{label}</span>
    </button>
  )
}

function Chart({
  data,
  type,
}: {
  data: { label: string; value: number }[]
  type: ChartType
}) {
  if (!data.length) {
    return <div className="chart-empty">Select numeric cells to chart.</div>
  }

  if (type === 'line') {
    const width = 760
    const height = 300
    const padX = 48
    const padTop = 25
    const padBottom = 48
    const values = data.map((item) => item.value)
    const min = Math.min(0, ...values)
    const max = Math.max(0, ...values)
    const span = Math.max(1, max - min)
    const xStep = data.length > 1 ? (width - padX * 2) / (data.length - 1) : 0
    const toY = (value: number) =>
      padTop + (max - value) / span * (height - padTop - padBottom)
    const points = data
      .map((item, index) => `${padX + index * xStep},${toY(item.value)}`)
      .join(' ')
    const zeroY = toY(0)

    return (
      <div className="line-chart-wrap">
        <svg className="line-chart" viewBox={`0 0 ${width} ${height}`} role="img" aria-label="Line chart">
          <line x1={padX} y1={zeroY} x2={width - padX} y2={zeroY} className="chart-axis" />
          {[0, 0.25, 0.5, 0.75, 1].map((fraction) => {
            const y = padTop + fraction * (height - padTop - padBottom)
            return <line key={fraction} x1={padX} y1={y} x2={width - padX} y2={y} className="chart-gridline" />
          })}
          <polyline points={points} className="chart-line" fill="none" />
          {data.map((item, index) => {
            const x = padX + index * xStep
            const y = toY(item.value)
            return (
              <g key={index}>
                <circle cx={x} cy={y} r="4.5" className="chart-point" />
                <text x={x} y={Math.max(12, y - 10)} className="chart-point-value" textAnchor="middle">
                  {item.value.toLocaleString()}
                </text>
                <text x={x} y={height - 18} className="chart-x-label" textAnchor="middle">
                  {item.label.slice(0, 10)}
                </text>
              </g>
            )
          })}
        </svg>
      </div>
    )
  }

  if (type === 'pie') {
    const positive = data
      .map((item) => ({ ...item, magnitude: Math.abs(item.value) }))
      .filter((item) => item.magnitude > 0)
    const total = positive.reduce((sum, item) => sum + item.magnitude, 0)

    if (!total) {
      return <div className="chart-empty">Pie charts need at least one non-zero value.</div>
    }

    let cursor = 0
    const segments = positive.map((item, index) => {
      const start = cursor
      const end = cursor + item.magnitude / total * 100
      cursor = end
      return {
        ...item,
        start,
        end,
        color: `hsl(${(index * 67 + 142) % 360} 55% 44%)`,
      }
    })
    const gradient = segments
      .map((segment) => `${segment.color} ${segment.start}% ${segment.end}%`)
      .join(', ')

    return (
      <div className="pie-chart-wrap">
        <div className="pie-chart" style={{ background: `conic-gradient(${gradient})` }}>
          <div className="pie-hole">
            <strong>{total.toLocaleString()}</strong>
            <span>Total</span>
          </div>
        </div>
        <div className="pie-legend">
          {segments.map((segment, index) => (
            <div className="pie-legend-item" key={index}>
              <span className="pie-dot" style={{ background: segment.color }} />
              <span className="pie-label">{segment.label}</span>
              <strong>{segment.value.toLocaleString()}</strong>
              <span className="pie-percent">{((segment.magnitude / total) * 100).toFixed(1)}%</span>
            </div>
          ))}
        </div>
      </div>
    )
  }

  const max = Math.max(...data.map((item) => Math.abs(item.value)), 1)

  return (
    <div className="bars">
      {data.map((item, index) => (
        <div className="bar-item" key={index}>
          <span className="bar-value">{item.value.toLocaleString()}</span>
          <div
            className="bar"
            style={{ height: Math.max(4, Math.abs(item.value) / max * 240) }}
          />
          <span className="bar-label">{item.label.slice(0, 14)}</span>
        </div>
      ))}
    </div>
  )
}

function MiniChart({
  data,
  type,
}: {
  data: { label: string; value: number }[]
  type: ChartType
}) {
  if (!data.length) return <div className="mini-chart-empty">No numeric data</div>

  const sample = data.slice(0, 16)
  const max = Math.max(...sample.map((item) => Math.abs(item.value)), 1)
  const min = Math.min(0, ...sample.map((item) => item.value))
  const maxValue = Math.max(0, ...sample.map((item) => item.value))
  const span = Math.max(1, maxValue - min)

  if (type === 'pie') {
    const positive = sample
      .map((item) => ({ ...item, magnitude: Math.abs(item.value) }))
      .filter((item) => item.magnitude > 0)
    const total = positive.reduce((sum, item) => sum + item.magnitude, 0)
    if (!total) return <div className="mini-chart-empty">No non-zero values</div>

    let cursor = 0
    const gradient = positive.map((item, index) => {
      const start = cursor
      const end = cursor + item.magnitude / total * 100
      cursor = end
      const color = `hsl(${(index * 67 + 142) % 360} 55% 44%)`
      return `${color} ${start}% ${end}%`
    }).join(', ')

    return (
      <div className="mini-chart mini-pie-chart">
        <div className="mini-pie" style={{ background: `conic-gradient(${gradient})` }} />
        <div className="mini-chart-caption">{positive.length} categories</div>
      </div>
    )
  }

  if (type === 'line') {
    const width = 360
    const height = 180
    const padding = 22
    const xStep = sample.length > 1 ? (width - padding * 2) / (sample.length - 1) : 0
    const points = sample.map((item, index) => {
      const x = padding + index * xStep
      const y = padding + (maxValue - item.value) / span * (height - padding * 2)
      return `${x},${y}`
    }).join(' ')

    return (
      <div className="mini-chart">
        <svg viewBox={`0 0 ${width} ${height}`} className="mini-line-svg" aria-label="Worksheet line chart">
          <polyline points={points} className="mini-line" fill="none" />
          {sample.map((item, index) => {
            const x = padding + index * xStep
            const y = padding + (maxValue - item.value) / span * (height - padding * 2)
            return <circle key={index} cx={x} cy={y} r="3" className="mini-line-point" />
          })}
        </svg>
      </div>
    )
  }

  return (
    <div className="mini-chart mini-bars">
      {sample.map((item, index) => (
        <div className="mini-bar-column" key={index} title={`${item.label}: ${item.value}`}>
          <div className="mini-bar" style={{ height: `${Math.max(3, Math.abs(item.value) / max * 100)}%` }} />
        </div>
      ))}
    </div>
  )
}

