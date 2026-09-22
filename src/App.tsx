import { useEffect, useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'
import { displayValue } from './formulas'
import { COLS, ROWS, cellKey, colToName, createBlankSheet, getCell, getSelectedPoints, normalizeSelection, parseCsv, pointToAddress, toCsv } from './spreadsheet'
import type { CellData, CellFormat, NumberFormat, Point, Selection, SheetData, WorkbookData } from './types'

const STORAGE_KEY = 'excel-micro-workbook-v1'
type Tab = 'Home' | 'Insert' | 'Data' | 'View'

function newWorkbook(): WorkbookData {
  const sheet = createBlankSheet(1)
  return { id: crypto.randomUUID(), title: 'Book 1', activeSheetId: sheet.id, sheets: [sheet], updatedAt: Date.now() }
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
  if (!Number.isFinite(n) || value === '') return value
  if (format === 'currency') return n.toLocaleString(undefined, { style: 'currency', currency: 'USD' })
  if (format === 'percent') return n.toLocaleString(undefined, { style: 'percent', maximumFractionDigits: 2 })
  if (format === 'number') return n.toLocaleString(undefined, { maximumFractionDigits: 4 })
  return value
}

export default function App() {
  const [book, setBook] = useState<WorkbookData>(loadWorkbook)
  const [selection, setSelection] = useState<Selection>({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } })
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
  const fileRef = useRef<HTMLInputElement>(null)
  const gridRef = useRef<HTMLDivElement>(null)

  const sheet = useMemo(() => book.sheets.find((item) => item.id === book.activeSheetId) || book.sheets[0], [book])
  const point = selection.focus
  const cell = getCell(sheet, point.row, point.col)
  const range = normalizeSelection(selection)

  useEffect(() => {
    const timer = window.setTimeout(() => localStorage.setItem(STORAGE_KEY, JSON.stringify(book)), 200)
    return () => window.clearTimeout(timer)
  }, [book])

  useEffect(() => {
    const stop = () => setDragging(false)
    window.addEventListener('mouseup', stop)
    return () => window.removeEventListener('mouseup', stop)
  }, [])

  function activeSheet(next: WorkbookData) {
    return next.sheets.find((item) => item.id === next.activeSheetId) || next.sheets[0]
  }

  function mutate(change: (next: WorkbookData) => void) {
    setHistory((items) => [...items, book].slice(-50))
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
    const next = { row: Math.max(0, Math.min(ROWS - 1, point.row + dr)), col: Math.max(0, Math.min(COLS - 1, point.col + dc)) }
    setSelection((current) => ({ anchor: extend ? current.anchor : next, focus: next }))
  }

  function applyFormat(patch: Partial<CellFormat>) {
    const points = getSelectedPoints(selection)
    mutate((next) => {
      const s = activeSheet(next)
      points.forEach((p) => {
        const key = cellKey(p.row, p.col)
        const old = s.cells[key] || { value: '' }
        s.cells[key] = { ...old, format: { ...(old.format || {}), ...patch } }
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

  function undo() {
    const previous = history.at(-1)
    if (!previous) return
    setFuture((items) => [book, ...items].slice(0, 50))
    setHistory((items) => items.slice(0, -1))
    setBook(previous)
  }

  function redo() {
    const next = future[0]
    if (!next) return
    setHistory((items) => [...items, book].slice(-50))
    setFuture((items) => items.slice(1))
    setBook(next)
  }

  async function copySelected() {
    const r = normalizeSelection(selection)
    const rows: string[] = []
    for (let row = r.top; row <= r.bottom; row += 1) {
      const values: string[] = []
      for (let col = r.left; col <= r.right; col += 1) values.push(getCell(sheet, row, col).value)
      rows.push(values.join('\t'))
    }
    await navigator.clipboard.writeText(rows.join('\n'))
  }

  function pasteText(text: string) {
    const rows = text.replace(/\r/g, '').split('\n').map((row) => row.split('\t'))
    mutate((next) => {
      const s = activeSheet(next)
      rows.forEach((values, ro) => values.forEach((value, co) => {
        const row = point.row + ro
        const col = point.col + co
        if (row >= ROWS || col >= COLS) return
        const key = cellKey(row, col)
        s.cells[key] = { ...(s.cells[key] || { value: '' }), value }
      }))
    })
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
    if (cmd && event.key.toLowerCase() === 'z') { event.preventDefault(); event.shiftKey ? redo() : undo(); return }
    if (cmd && event.key.toLowerCase() === 'y') { event.preventDefault(); redo(); return }
    if (cmd && event.key.toLowerCase() === 'c') { event.preventDefault(); void copySelected(); return }
    if (cmd && event.key.toLowerCase() === 'f') { event.preventDefault(); setFindOpen(true); return }
    if (event.key === 'Delete' || event.key === 'Backspace') { event.preventDefault(); clearSelected(); return }
    if (event.key === 'Enter' || event.key === 'F2') { event.preventDefault(); beginEdit(); return }
    if (event.key === 'Tab') { event.preventDefault(); move(0, event.shiftKey ? -1 : 1); return }
    const dirs: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }
    if (dirs[event.key]) { event.preventDefault(); const d = dirs[event.key]; move(d[0], d[1], event.shiftKey); return }
    if (!cmd && !event.altKey && event.key.length === 1) { event.preventDefault(); beginEdit(event.key) }
  }

  function addSheet() {
    mutate((next) => {
      const s = createBlankSheet(next.sheets.length + 1)
      next.sheets.push(s)
      next.activeSheetId = s.id
    })
    setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } })
  }

  function renameSheet(item: SheetData) {
    const name = window.prompt('Rename sheet', item.name)?.trim()
    if (!name) return
    mutate((next) => { const target = next.sheets.find((s) => s.id === item.id); if (target) target.name = name.slice(0, 31) })
  }

  function deleteSheet() {
    if (book.sheets.length === 1) return
    mutate((next) => {
      const index = next.sheets.findIndex((s) => s.id === next.activeSheetId)
      next.sheets.splice(index, 1)
      next.activeSheetId = next.sheets[Math.max(0, index - 1)].id
    })
  }

  function sortSelected(direction: 1 | -1) {
    if (range.top === range.bottom) return
    mutate((next) => {
      const s = activeSheet(next)
      const rows: CellData[][] = []
      for (let row = range.top; row <= range.bottom; row += 1) {
        const values: CellData[] = []
        for (let col = range.left; col <= range.right; col += 1) values.push(structuredClone(getCell(s, row, col)))
        rows.push(values)
      }
      rows.sort((a, b) => {
        const av = displayValue(a[0].value, s)
        const bv = displayValue(b[0].value, s)
        const an = Number(av), bn = Number(bv)
        const result = Number.isFinite(an) && Number.isFinite(bn) ? an - bn : av.localeCompare(bv, undefined, { numeric: true })
        return result * direction
      })
      rows.forEach((values, ro) => values.forEach((value, co) => { s.cells[cellKey(range.top + ro, range.left + co)] = value }))
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
        setSelection({ anchor: p, focus: p })
        document.querySelector('[data-cell="' + p.row + ':' + p.col + '"]')?.scrollIntoView({ block: 'center', inline: 'center' })
        break
      }
    }
  }

  function exportXlsx() {
    const out = XLSX.utils.book_new()
    book.sheets.forEach((s) => {
      let maxRow = 0, maxCol = 0
      Object.keys(s.cells).forEach((key) => { const bits = key.split(':').map(Number); maxRow = Math.max(maxRow, bits[0]); maxCol = Math.max(maxCol, bits[1]) })
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
      for (let row = 0; row <= maxRow; row += 1) for (let col = 0; col <= maxCol; col += 1) {
        const raw = getCell(s, row, col).value
        if (!raw.startsWith('=')) continue
        const address = XLSX.utils.encode_cell({ r: row, c: col })
        ws[address] = { t: 'n', f: raw.slice(1) }
      }
      XLSX.utils.book_append_sheet(out, ws, s.name.slice(0, 31))
    })
    XLSX.writeFile(out, safeName(book.title) + '.xlsx')
  }

  function fromRows(rows: unknown[][], name: string): SheetData {
    const s: SheetData = { id: crypto.randomUUID(), name: name.slice(0, 31) || 'Sheet1', cells: {} }
    rows.slice(0, ROWS).forEach((row, r) => row.slice(0, COLS).forEach((value, c) => { if (value !== '' && value != null) s.cells[cellKey(r, c)] = { value: String(value) } }))
    return s
  }

  async function importFile(file: File) {
    const ext = file.name.split('.').pop()?.toLowerCase()
    if (ext === 'json') {
      const parsed = JSON.parse(await file.text()) as WorkbookData
      if (!parsed.sheets?.length) throw new Error('Invalid Excel Micro backup')
      setHistory((items) => [...items, book].slice(-50)); setBook(parsed); return
    }
    if (ext === 'csv') {
      const s = fromRows(parseCsv(await file.text()), file.name.replace(/\.csv$/i, ''))
      mutate((next) => { next.sheets.push(s); next.activeSheetId = s.id }); return
    }
    const source = XLSX.read(await file.arrayBuffer(), { type: 'array', cellFormula: true })
    const sheets = source.SheetNames.map((name) => {
      const ws = source.Sheets[name]
      const ref = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1')
      const rows: string[][] = []
      for (let r = ref.s.r; r <= Math.min(ref.e.r, ROWS - 1); r += 1) {
        const values: string[] = []
        for (let c = ref.s.c; c <= Math.min(ref.e.c, COLS - 1); c += 1) {
          const sourceCell = ws[XLSX.utils.encode_cell({ r, c })]
          values.push(sourceCell?.f ? '=' + sourceCell.f : sourceCell?.v == null ? '' : String(sourceCell.v))
        }
        rows.push(values)
      }
      return fromRows(rows, name)
    })
    if (!sheets.length) return
    setHistory((items) => [...items, book].slice(-50))
    setBook({ id: crypto.randomUUID(), title: file.name.replace(/\.[^.]+$/, ''), sheets, activeSheetId: sheets[0].id, updatedAt: Date.now() })
  }

  const stats = useMemo(() => {
    const nums = getSelectedPoints(selection).map((p) => Number(displayValue(getCell(sheet, p.row, p.col).value, sheet))).filter(Number.isFinite)
    return { count: nums.length, sum: nums.reduce((a, b) => a + b, 0), avg: nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0 }
  }, [selection, sheet])

  const chart = useMemo(() => {
    const data: { label: string; value: number }[] = []
    for (let row = range.top; row <= range.bottom; row += 1) {
      const label = range.right > range.left ? formatted(getCell(sheet, row, range.left), sheet) : String(row + 1)
      const valueCol = range.right > range.left ? range.left + 1 : range.left
      const value = Number(displayValue(getCell(sheet, row, valueCol).value, sheet))
      if (Number.isFinite(value)) data.push({ label: label || String(row + 1), value })
    }
    return data.slice(0, 20)
  }, [selection, sheet])

  return (
    <div className="app-shell">
      <header className="titlebar">
        <div className="logo-box">X</div>
        <strong>Excel Micro</strong>
        <input className="book-title" value={book.title} onChange={(e) => setBook((current) => ({ ...current, title: e.target.value }))} />
        <span className="saved">Saved locally</span>
        <div className="title-actions">
          <button onClick={() => fileRef.current?.click()}>Open</button>
          <button className="download-btn" onClick={exportXlsx}>Download .xlsx</button>
        </div>
      </header>

      <div className="quickbar">
        <button onClick={undo} disabled={!history.length}>↶</button><button onClick={redo} disabled={!future.length}>↷</button>
        <span className="divider" /><button onClick={() => void copySelected()}>Copy</button><button onClick={() => setFindOpen(true)}>Find</button>
      </div>

      <nav className="tabs">{(['Home', 'Insert', 'Data', 'View'] as Tab[]).map((item) => <button key={item} className={tab === item ? 'active' : ''} onClick={() => setTab(item)}>{item}</button>)}</nav>

      <section className="ribbon">
        {tab === 'Home' && <>
          <Group name="Font"><button className={cell.format?.bold ? 'on' : ''} onClick={() => applyFormat({ bold: !cell.format?.bold })}><b>B</b></button><button className={cell.format?.italic ? 'on' : ''} onClick={() => applyFormat({ italic: !cell.format?.italic })}><i>I</i></button><button className={cell.format?.underline ? 'on' : ''} onClick={() => applyFormat({ underline: !cell.format?.underline })}><u>U</u></button><label className="color-label">Text<input type="color" value={cell.format?.color || '#202020'} onChange={(e) => applyFormat({ color: e.target.value })} /></label><label className="color-label">Fill<input type="color" value={cell.format?.background || '#ffffff'} onChange={(e) => applyFormat({ background: e.target.value })} /></label></Group>
          <Group name="Alignment"><button onClick={() => applyFormat({ align: 'left' })}>Left</button><button onClick={() => applyFormat({ align: 'center' })}>Center</button><button onClick={() => applyFormat({ align: 'right' })}>Right</button></Group>
          <Group name="Number"><select value={cell.format?.numberFormat || 'general'} onChange={(e) => applyFormat({ numberFormat: e.target.value as NumberFormat })}><option value="general">General</option><option value="number">Number</option><option value="currency">Currency</option><option value="percent">Percent</option></select></Group>
        </>}
        {tab === 'Insert' && <><Group name="Functions"><button onClick={() => beginEdit('=SUM(')}>Σ AutoSum</button><button onClick={() => beginEdit('=AVERAGE(')}>fx Average</button></Group><Group name="Chart"><button onClick={() => setChartOpen(true)}>Bar chart</button></Group></>}
        {tab === 'Data' && <><Group name="Sort"><button onClick={() => sortSelected(1)}>A → Z</button><button onClick={() => sortSelected(-1)}>Z → A</button></Group><Group name="Files"><button onClick={() => fileRef.current?.click()}>Import</button><button onClick={() => download(safeName(sheet.name) + '.csv', toCsv(sheet), 'text/csv')}>CSV</button><button onClick={() => download(safeName(book.title) + '.excelmicro.json', JSON.stringify(book, null, 2), 'application/json')}>Backup</button></Group></>}
        {tab === 'View' && <Group name="Zoom"><button onClick={() => setZoom((z) => Math.max(60, z - 10))}>−</button><span className="zoom-readout">{zoom}%</span><button onClick={() => setZoom((z) => Math.min(160, z + 10))}>+</button></Group>}
      </section>

      <div className="formula-row"><div className="name-box">{pointToAddress(point)}</div><div className="fx">fx</div><input value={editing ? draft : cell.value} onFocus={() => { if (!editing) beginEdit() }} onChange={(e) => setDraft(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') endEdit(true); if (e.key === 'Escape') endEdit(false) }} onBlur={() => { if (editing) endEdit(true) }} /></div>

      <main className="workspace">
        <div className="grid-viewport" ref={gridRef} tabIndex={0} onKeyDown={onGridKey} style={{ '--zoom': zoom / 100 } as React.CSSProperties}>
          <div className="sheet-grid">
            <div className="corner" />
            {Array.from({ length: COLS }, (_, col) => <div key={'h' + col} className="col-head">{colToName(col)}</div>)}
            {Array.from({ length: ROWS }, (_, row) => <div className="grid-row" key={'r' + row}>
              <div className="row-head">{row + 1}</div>
              {Array.from({ length: COLS }, (_, col) => {
                const data = getCell(sheet, row, col)
                const active = point.row === row && point.col === col
                const selected = row >= range.top && row <= range.bottom && col >= range.left && col <= range.right
                return <div data-cell={row + ':' + col} key={row + ':' + col} className={'cell ' + (selected ? 'selected ' : '') + (active ? 'active ' : '')} style={{ fontWeight: data.format?.bold ? 700 : 400, fontStyle: data.format?.italic ? 'italic' : 'normal', textDecoration: data.format?.underline ? 'underline' : 'none', textAlign: data.format?.align || 'left', color: data.format?.color, background: data.format?.background }} onMouseDown={(e) => { e.preventDefault(); if (editing) endEdit(true); const p = { row, col }; setDragging(true); setSelection((current) => ({ anchor: e.shiftKey ? current.anchor : p, focus: p })); gridRef.current?.focus() }} onMouseEnter={() => { if (dragging) setSelection((current) => ({ ...current, focus: { row, col } })) }} onDoubleClick={() => beginEdit()}>
                  {active && editing ? <input className="cell-input" autoFocus value={draft} onChange={(e) => setDraft(e.target.value)} onBlur={() => endEdit(true)} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); endEdit(true); move(1, 0) } if (e.key === 'Escape') endEdit(false) }} /> : <span>{formatted(data, sheet)}</span>}
                </div>
              })}
            </div>)}
          </div>
        </div>
      </main>

      <div className="sheetbar"><button className="plus" onClick={addSheet}>＋</button><div className="sheet-tabs">{book.sheets.map((item) => <button key={item.id} className={item.id === book.activeSheetId ? 'active' : ''} onClick={() => { setBook((current) => ({ ...current, activeSheetId: item.id })); setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } }) }} onDoubleClick={() => renameSheet(item)}>{item.name}</button>)}</div><button onClick={deleteSheet} disabled={book.sheets.length === 1}>Delete</button></div>
      <footer className="status"><span>Ready</span><span className="spacer" />{stats.count > 0 && <><span>Average: {stats.avg.toLocaleString(undefined, { maximumFractionDigits: 4 })}</span><span>Count: {stats.count}</span><span>Sum: {stats.sum.toLocaleString(undefined, { maximumFractionDigits: 4 })}</span></>}<span>{zoom}%</span></footer>

      {findOpen && <div className="find-box"><input autoFocus placeholder="Find in sheet" value={find} onChange={(e) => setFind(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') findNext(); if (e.key === 'Escape') setFindOpen(false) }} /><button onClick={findNext}>Find next</button><button onClick={() => setFindOpen(false)}>×</button></div>}
      {chartOpen && <div className="overlay" onMouseDown={() => setChartOpen(false)}><div className="chart-card" onMouseDown={(e) => e.stopPropagation()}><div className="chart-title"><div><h2>Quick chart</h2><p>Select one numeric column, or labels + values in two columns.</p></div><button onClick={() => setChartOpen(false)}>×</button></div><Chart data={chart} /></div></div>}
      <input ref={fileRef} className="hidden" type="file" accept=".xlsx,.xls,.csv,.json" onChange={(e) => { const file = e.target.files?.[0]; if (file) void importFile(file).catch((error) => window.alert(error instanceof Error ? error.message : 'Import failed')) }} />
    </div>
  )
}

function Group({ name, children }: { name: string; children: React.ReactNode }) {
  return <div className="group"><div className="group-controls">{children}</div><small>{name}</small></div>
}

function Chart({ data }: { data: { label: string; value: number }[] }) {
  if (!data.length) return <div className="chart-empty">Select numeric cells to chart.</div>
  const max = Math.max(...data.map((item) => Math.abs(item.value)), 1)
  return <div className="bars">{data.map((item, index) => <div className="bar-item" key={index}><span className="bar-value">{item.value}</span><div className="bar" style={{ height: Math.max(4, Math.abs(item.value) / max * 240) }} /><span className="bar-label">{item.label.slice(0, 12)}</span></div>)}</div>
}
