import type { SheetData } from './types'
import { cellKey, parseAddress } from './spreadsheet'

type Resolver = (address: string) => number

function splitArgs(source: string): string[] {
  const args: string[] = []
  let depth = 0
  let token = ''
  for (const char of source) {
    if (char === '(') depth += 1
    if (char === ')') depth -= 1
    if (char === ',' && depth === 0) {
      args.push(token.trim())
      token = ''
    } else token += char
  }
  args.push(token.trim())
  return args
}

function rangeValues(token: string, sheet: SheetData, stack: Set<string>): number[] | null {
  const match = /^([A-Z]+\d+):([A-Z]+\d+)$/i.exec(token.trim())
  if (!match) return null
  const from = parseAddress(match[1])
  const to = parseAddress(match[2])
  if (!from || !to) return []
  const values: number[] = []
  for (let row = Math.min(from.row, to.row); row <= Math.max(from.row, to.row); row += 1) {
    for (let col = Math.min(from.col, to.col); col <= Math.max(from.col, to.col); col += 1) {
      const raw = sheet.cells[cellKey(row, col)]?.value ?? ''
      values.push(toNumber(raw, sheet, stack))
    }
  }
  return values
}

function toNumber(raw: string, sheet: SheetData, stack: Set<string>): number {
  if (!raw) return 0
  if (raw.startsWith('=')) {
    const evaluated = evaluateFormula(raw, sheet, stack)
    return typeof evaluated === 'number' && Number.isFinite(evaluated) ? evaluated : Number(evaluated) || 0
  }
  const normalized = raw.replace(/[$,%]/g, '')
  const numeric = Number(normalized)
  return Number.isFinite(numeric) ? numeric : 0
}

function evalComparison(source: string, resolver: Resolver): boolean {
  const operators = ['>=', '<=', '<>', '!=', '=', '>', '<']
  for (const op of operators) {
    const index = source.indexOf(op)
    if (index > -1) {
      const left = evaluateExpression(source.slice(0, index), resolver)
      const right = evaluateExpression(source.slice(index + op.length), resolver)
      if (op === '>=') return left >= right
      if (op === '<=') return left <= right
      if (op === '>') return left > right
      if (op === '<') return left < right
      if (op === '=') return left === right
      return left !== right
    }
  }
  return evaluateExpression(source, resolver) !== 0
}

function evaluateExpression(source: string, resolver: Resolver): number {
  const withRefs = source.replace(/\b([A-Z]+\d+)\b/gi, (match) => String(resolver(match)))
  const safe = withRefs.replace(/\s+/g, '')
  if (!safe || !/^[0-9+\-*/().%]+$/.test(safe)) return Number(safe) || 0
  try {
    const result = Function(`"use strict"; return (${safe})`)() as unknown
    return typeof result === 'number' && Number.isFinite(result) ? result : 0
  } catch {
    return 0
  }
}

export function evaluateFormula(raw: string, sheet: SheetData, stack = new Set<string>()): number | string {
  if (!raw.startsWith('=')) return raw
  let formula = raw.slice(1).trim()
  const resolver: Resolver = (address) => {
    const point = parseAddress(address)
    if (!point) return 0
    const key = cellKey(point.row, point.col)
    if (stack.has(key)) return 0
    const nextStack = new Set(stack)
    nextStack.add(key)
    return toNumber(sheet.cells[key]?.value ?? '', sheet, nextStack)
  }

  const functionPattern = /(SUM|AVERAGE|AVG|MIN|MAX|COUNT)\(([^()]*)\)/i
  let guard = 0
  while (functionPattern.test(formula) && guard < 50) {
    guard += 1
    formula = formula.replace(functionPattern, (_full, fnName: string, body: string) => {
      const values = splitArgs(body).flatMap((arg) => rangeValues(arg, sheet, stack) ?? [evaluateExpression(arg, resolver)])
      const fn = fnName.toUpperCase()
      if (fn === 'SUM') return String(values.reduce((sum, value) => sum + value, 0))
      if (fn === 'AVERAGE' || fn === 'AVG') return String(values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : 0)
      if (fn === 'MIN') return String(values.length ? Math.min(...values) : 0)
      if (fn === 'MAX') return String(values.length ? Math.max(...values) : 0)
      return String(values.filter((value) => Number.isFinite(value)).length)
    })
  }

  const ifMatch = /^IF\((.*)\)$/i.exec(formula)
  if (ifMatch) {
    const args = splitArgs(ifMatch[1])
    if (args.length >= 3) {
      const branch = evalComparison(args[0], resolver) ? args[1] : args[2]
      if (/^".*"$/.test(branch.trim())) return branch.trim().slice(1, -1)
      return evaluateExpression(branch, resolver)
    }
  }

  return evaluateExpression(formula, resolver)
}

export function displayValue(raw: string, sheet: SheetData): string {
  if (!raw.startsWith('=')) return raw
  const result = evaluateFormula(raw, sheet)
  return typeof result === 'number' ? String(Math.round((result + Number.EPSILON) * 1e10) / 1e10) : String(result)
}
