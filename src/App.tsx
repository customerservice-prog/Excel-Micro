import { useEffect, useMemo, useRef, useState, type CSSProperties, type ReactNode } from 'react'
import * as XLSX from 'xlsx'
import { displayValue } from './formulas'
import {
  COLS,
  DEFAULT_COLUMN_WIDTH,
  ROWS,
  cellKey,
  colToName,
  createBlankSheet,
  getCell,
  getColumnWidth,
  getSelectedPoints,
  normalizeSelection,
  parseCsv,
  parseSelectionAddress,
  pointToAddress,
  selectionToAddress,
  shiftFormulaReferences,
  toCsv,
} from './spreadsheet'
import type { CellData, CellFormat, NumberFormat, Point, Selection, SheetData, WorkbookData } from './types'

const STORAGE_KEY = 'excel-micro-workbook-v1'
type Tab = 'Home' | 'Insert' | 'Formulas' | 'Data' | 'View'

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

function formatted(cell: CellData, sheet: SheetData) {
  const value = displayValue(cell.value, sheet)
  const n = Number(value)
  const format = cell.format?.numberFormat || 'general'
  const decimals = cell.format?.decimals

  if (!Number.isFinite(n) || value === '') return value

  if (format === 'currency') {
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

  if (format === 'number') {
    return n.toLocaleString(undefined, {
      minimumFractionDigits: decimals ?? 0,
      maximumFractionDigits: decimals ?? 4,
    })
  }

  return value
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
  const [find, setFind] = useState('')
  const [findOpen, setFindOpen] = useState(false)
  const [chartOpen, setChartOpen] = useState(false)
  const [zoom, setZoom] = useState(100)
  const [nameBox, setNameBox] = useState('A1')
  const [saveStatus, setSaveStatus] = useState('Saved')
  const [notice, setNotice] = useState('')
  const fileRef = useRef<HTMLInputElement>(null)
  const gridRef = useRef<HTMLDivElement>(null)
  const columnResizeRef = useRef<{ col: number; startX: number; startWidth: number } | null>(null)

  const sheet = useMemo(
    () => book.sheets.find((item) => item.id === book.activeSheetId) || book.sheets[0],
    [book],
  )
  const point = selection.focus
  const cell = getCell(sheet, point.row, point.col)
  const range = normalizeSelection(selection)

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
    const stop = () => setDragging(false)
    window.addEventListener('mouseup', stop)
    return () => window.removeEventListener('mouseup', stop)
  }, [])

  useEffect(() => {
    const onMove = (event: MouseEvent) => {
      const resize = columnResizeRef.current
      if (!resize) return

      const width = Math.max(48, Math.min(360, resize.startWidth + event.clientX - resize.startX))

      setBook((current) => {
        const next = structuredClone(current)
        const target = next.sheets.find((item) => item.id === next.activeSheetId) || next.sheets[0]
        target.columnWidths ||= {}
        target.columnWidths[String(resize.col)] = Math.round(width)
        next.updatedAt = Date.now()
        return next
      })
    }

    const onUp = () => {
      if (columnResizeRef.current) setNotice('Column width updated')
      columnResizeRef.current = null
      document.body.classList.remove('is-column-resizing')
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

  function activeSheet(next: WorkbookData) {
    return next.sheets.find((item) => item.id === next.activeSheetId) || next.sheets[0]
  }

  function mutate(change: (next: WorkbookData) => void) {
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
    setDraft(seed === undefined ? cell.value : seed)
    setEditing(true)
  }

  function endEdit(save = true) {
    if (save) setValue(point, draft)
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

  function applyFormat(patch: Partial<CellFormat>) {
    const points = getSelectedPoints(selection)

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
    const points = getSelectedPoints(selection)

    mutate((next) => {
      const s = activeSheet(next)

      points.forEach((p) => {
        const key = cellKey(p.row, p.col)
        const old = s.cells[key]
        if (!old) return
        if (old.format) s.cells[key] = { ...old, value: '' }
        else delete s.cells[key]
      })
    })
  }

  function clearFormatting() {
    const points = getSelectedPoints(selection)

    mutate((next) => {
      const s = activeSheet(next)

      points.forEach((p) => {
        const key = cellKey(p.row, p.col)
        const old = s.cells[key]
        if (!old) return
        if (old.value) s.cells[key] = { value: old.value }
        else delete s.cells[key]
      })
    })
  }

  function undo() {
    const previous = history.at(-1)
    if (!previous) return

    setFuture((items) => [book, ...items].slice(0, 75))
    setHistory((items) => items.slice(0, -1))
    setBook(previous)
  }

  function redo() {
    const next = future[0]
    if (!next) return

    setHistory((items) => [...items, book].slice(-75))
    setFuture((items) => items.slice(1))
    setBook(next)
  }

  async function copySelected() {
    const r = normalizeSelection(selection)
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

    mutate((next) => {
      const s = activeSheet(next)

      rows.forEach((values, ro) => {
        values.forEach((value, co) => {
          const row = point.row + ro
          const col = point.col + co
          if (row >= ROWS || col >= COLS) return

          const key = cellKey(row, col)
          s.cells[key] = {
            ...(s.cells[key] || { value: '' }),
            value,
          }
        })
      })
    })
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
        const av = displayValue(a[0].value, s)
        const bv = displayValue(b[0].value, s)
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
    const parsed = parseSelectionAddress(nameBox)

    if (!parsed) {
      setNameBox(selectionToAddress(selection))
      setNotice('Enter a cell like B12 or range like A1:D20')
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
      const value = formatted(getCell(sheet, row, col), sheet)
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
    const nums = getSelectedPoints(selection)
      .map((p) => Number(displayValue(getCell(sheet, p.row, p.col).value, sheet)))
      .filter(Number.isFinite)

    return {
      count: nums.length,
      sum: nums.reduce((a, b) => a + b, 0),
      avg: nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0,
    }
  }, [selection, sheet])

  const chart = useMemo(() => {
    const data: { label: string; value: number }[] = []

    for (let row = range.top; row <= range.bottom; row += 1) {
      const label = range.right > range.left
        ? formatted(getCell(sheet, row, range.left), sheet)
        : String(row + 1)
      const valueCol = range.right > range.left ? range.left + 1 : range.left
      const value = Number(displayValue(getCell(sheet, row, valueCol).value, sheet))
      if (Number.isFinite(value)) data.push({ label: label || String(row + 1), value })
    }

    return data.slice(0, 24)
  }, [range.bottom, range.left, range.right, range.top, sheet])

  const columnTemplate = useMemo(() => {
    const widths = Array.from({ length: COLS }, (_, col) => {
      const width = getColumnWidth(sheet, col) * zoom / 100
      return `${Math.round(width)}px`
    })
    return `46px ${widths.join(' ')}`
  }, [sheet, zoom])

  const gridStyle = {
    '--row-height': `${Math.max(20, Math.round(25 * zoom / 100))}px`,
    gridTemplateColumns: columnTemplate,
    fontSize: `${Math.max(10, 12 * zoom / 100)}px`,
  } as CSSProperties

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
        {(['Home', 'Insert', 'Formulas', 'Data', 'View'] as Tab[]).map((item) => (
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
            </Group>

            <Group name="Number">
              <select
                value={cell.format?.numberFormat || 'general'}
                onChange={(e) => applyFormat({ numberFormat: e.target.value as NumberFormat })}
              >
                <option value="general">General</option>
                <option value="number">Number</option>
                <option value="currency">Currency</option>
                <option value="percent">Percent</option>
              </select>
              <button className="compact" title="Decrease decimals" onClick={() => adjustDecimals(-1)}>.0←</button>
              <button className="compact" title="Increase decimals" onClick={() => adjustDecimals(1)}>→.00</button>
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

            <Group name="Visuals">
              <RibbonButton icon="▥" label="Bar chart" onClick={() => setChartOpen(true)} />
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
            </Group>

            <div className="formula-help">
              <strong>Formula examples</strong>
              <span>=SUM(A1:A10)</span>
              <span>=AVERAGE(B2:B20)</span>
              <span>=IF(SUM(A1:A5)&gt;100,"Over","OK")</span>
              <span>=ROUND(A1*B1,2)</span>
            </div>
          </>
        )}

        {tab === 'Data' && (
          <>
            <Group name="Sort">
              <RibbonButton icon="A↓" label="A → Z" onClick={() => sortSelected(1)} />
              <RibbonButton icon="Z↓" label="Z → A" onClick={() => sortSelected(-1)} />
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

            {Array.from({ length: COLS }, (_, col) => (
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

            {Array.from({ length: ROWS }, (_, row) => (
              <div className="grid-row" key={'r' + row}>
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
                >
                  {row + 1}
                </div>

                {Array.from({ length: COLS }, (_, col) => {
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

                  return (
                    <div
                      data-cell={row + ':' + col}
                      key={row + ':' + col}
                      className={
                        'cell ' +
                        (selected ? 'selected ' : '') +
                        (active ? 'active ' : '') +
                        frozenClass +
                        (data.format?.wrap ? ' wrap ' : '')
                      }
                      style={{
                        fontWeight: data.format?.bold ? 700 : 400,
                        fontStyle: data.format?.italic ? 'italic' : 'normal',
                        textDecoration: decoration || 'none',
                        textAlign: data.format?.align || 'left',
                        color: data.format?.color,
                        background: data.format?.background,
                        fontSize: data.format?.fontSize ? `${data.format.fontSize}px` : undefined,
                      }}
                      onMouseDown={(e) => {
                        e.preventDefault()
                        if (editing) endEdit(true)
                        const p = { row, col }
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
                    >
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
                      ) : (
                        <span>{formatted(data, sheet)}</span>
                      )}
                    </div>
                  )
                })}
              </div>
            ))}
          </div>
        </div>
      </main>

      <div className="sheetbar">
        <button className="plus" title="New sheet" onClick={addSheet}>＋</button>

        <div className="sheet-tabs">
          {book.sheets.map((item) => (
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
        <span className="status-view">{sheet.freezeTopRow || sheet.freezeFirstColumn ? 'Frozen panes' : 'Normal view'}</span>
        <span>{zoom}%</span>
      </footer>

      {findOpen && (
        <div className="find-box">
          <div className="find-title">
            <strong>Find</strong>
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
        </div>
      )}

      {chartOpen && (
        <div className="overlay" onMouseDown={() => setChartOpen(false)}>
          <div className="chart-card" onMouseDown={(e) => e.stopPropagation()}>
            <div className="chart-title">
              <div>
                <span className="eyebrow">INSERTED FROM {selectionToAddress(selection)}</span>
                <h2>Quick bar chart</h2>
                <p>Select one numeric column, or labels + values in two columns.</p>
              </div>
              <button onClick={() => setChartOpen(false)}>×</button>
            </div>
            <Chart data={chart} />
          </div>
        </div>
      )}

      {notice && <div className="toast">{notice}</div>}

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

function Chart({ data }: { data: { label: string; value: number }[] }) {
  if (!data.length) {
    return <div className="chart-empty">Select numeric cells to chart.</div>
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
