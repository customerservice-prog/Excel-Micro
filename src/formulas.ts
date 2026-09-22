import type { SheetData } from './types'
import { cellKey, parseAddress } from './spreadsheet'

type FormulaValue = number | string | boolean | FormulaValue[]
export type FormulaResult = number | string | boolean

type TokenType =
  | 'number'
  | 'string'
  | 'identifier'
  | 'operator'
  | 'lparen'
  | 'rparen'
  | 'comma'
  | 'colon'
  | 'eof'

interface Token {
  type: TokenType
  value: string
}

function isNumeric(value: FormulaValue): boolean {
  if (Array.isArray(value)) return false
  if (typeof value === 'number') return Number.isFinite(value)
  if (typeof value === 'boolean') return true
  if (typeof value === 'string' && value.trim() !== '') return Number.isFinite(Number(value))
  return false
}

function toNumber(value: FormulaValue): number {
  if (Array.isArray(value)) return value.length ? toNumber(value[0]) : 0
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0
  if (typeof value === 'boolean') return value ? 1 : 0
  if (typeof value === 'string') {
    const normalized = value.replace(/[$,%]/g, '').trim()
    const numeric = Number(normalized)
    return Number.isFinite(numeric) ? numeric : 0
  }
  return 0
}

function toText(value: FormulaValue): string {
  if (Array.isArray(value)) return value.map(toText).join('')
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  return String(value)
}

function truthy(value: FormulaValue): boolean {
  if (Array.isArray(value)) return value.some(truthy)
  if (typeof value === 'boolean') return value
  if (typeof value === 'number') return value !== 0
  const normalized = value.trim().toUpperCase()
  if (!normalized || normalized === 'FALSE' || normalized === '0') return false
  return true
}

function flatten(values: FormulaValue[]): FormulaValue[] {
  return values.flatMap((value) => Array.isArray(value) ? flatten(value) : [value])
}

function numericValues(values: FormulaValue[]) {
  return flatten(values).filter(isNumeric).map(toNumber)
}

function compare(left: FormulaValue, right: FormulaValue, operator: string): boolean {
  const numeric = isNumeric(left) && isNumeric(right)
  const a = numeric ? toNumber(left) : toText(left).toLowerCase()
  const b = numeric ? toNumber(right) : toText(right).toLowerCase()

  if (operator === '=') return a === b
  if (operator === '<>' || operator === '!=') return a !== b
  if (operator === '>') return a > b
  if (operator === '<') return a < b
  if (operator === '>=') return a >= b
  if (operator === '<=') return a <= b
  return false
}

function matchesCriteria(value: FormulaValue, criterion: FormulaValue): boolean {
  const text = toText(criterion)
  const operatorMatch = /^(>=|<=|<>|!=|>|<|=)(.*)$/.exec(text)

  if (operatorMatch) {
    return compare(value, operatorMatch[2], operatorMatch[1])
  }

  if (text.includes('*') || text.includes('?')) {
    const escaped = text.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&')
    const pattern = '^' + escaped.replace(/\*/g, '.*').replace(/\?/g, '.') + '$'
    return new RegExp(pattern, 'i').test(toText(value))
  }

  return compare(value, criterion, '=')
}

function asList(value: FormulaValue | undefined): FormulaValue[] {
  if (value === undefined) return []
  return Array.isArray(value) ? flatten(value) : [value]
}

function isErrorValue(value: FormulaValue): boolean {
  return !Array.isArray(value) && typeof value === 'string' && value.startsWith('#')
}

const EXCEL_EPOCH = Date.UTC(1899, 11, 30)
const DAY_MS = 86_400_000

function excelSerialFromDate(date: Date): number {
  return (Date.UTC(
    date.getUTCFullYear(),
    date.getUTCMonth(),
    date.getUTCDate(),
    date.getUTCHours(),
    date.getUTCMinutes(),
    date.getUTCSeconds(),
    date.getUTCMilliseconds(),
  ) - EXCEL_EPOCH) / DAY_MS
}

function dateFromFormulaValue(value: FormulaValue): Date | null {
  if (isNumeric(value)) {
    const serial = toNumber(value)
    const date = new Date(EXCEL_EPOCH + serial * DAY_MS)
    return Number.isNaN(date.getTime()) ? null : date
  }

  const text = toText(value).trim()
  if (!text) return null
  const date = new Date(text)
  return Number.isNaN(date.getTime()) ? null : date
}

class Lexer {
  private index = 0

  constructor(private readonly source: string) {}

  tokenize(): Token[] {
    const tokens: Token[] = []

    while (this.index < this.source.length) {
      const char = this.source[this.index]

      if (/\s/.test(char)) {
        this.index += 1
        continue
      }

      if (char === '"') {
        tokens.push({ type: 'string', value: this.readString() })
        continue
      }

      if (char === "'") {
        tokens.push({ type: 'identifier', value: this.readQuotedSheetReference() })
        continue
      }

      if (/\d/.test(char) || (char === '.' && /\d/.test(this.source[this.index + 1] || ''))) {
        tokens.push({ type: 'number', value: this.readNumber() })
        continue
      }

      if (/[A-Za-z_$]/.test(char)) {
        tokens.push({ type: 'identifier', value: this.readIdentifier() })
        continue
      }

      if (char === '(') {
        tokens.push({ type: 'lparen', value: char })
        this.index += 1
        continue
      }

      if (char === ')') {
        tokens.push({ type: 'rparen', value: char })
        this.index += 1
        continue
      }

      if (char === ',') {
        tokens.push({ type: 'comma', value: char })
        this.index += 1
        continue
      }

      if (char === ':') {
        tokens.push({ type: 'colon', value: char })
        this.index += 1
        continue
      }

      const pair = this.source.slice(this.index, this.index + 2)
      if (['>=', '<=', '<>', '!='].includes(pair)) {
        tokens.push({ type: 'operator', value: pair })
        this.index += 2
        continue
      }

      if ('+-*/^%&=<>'.includes(char)) {
        tokens.push({ type: 'operator', value: char })
        this.index += 1
        continue
      }

      this.index += 1
    }

    tokens.push({ type: 'eof', value: '' })
    return tokens
  }

  private readString() {
    this.index += 1
    let value = ''

    while (this.index < this.source.length) {
      const char = this.source[this.index]

      if (char === '"' && this.source[this.index + 1] === '"') {
        value += '"'
        this.index += 2
        continue
      }

      if (char === '"') {
        this.index += 1
        break
      }

      value += char
      this.index += 1
    }

    return value
  }

  private readNumber() {
    const start = this.index

    while (this.index < this.source.length && /[0-9.]/.test(this.source[this.index])) {
      this.index += 1
    }

    return this.source.slice(start, this.index)
  }

  private readIdentifier() {
    const start = this.index

    while (
      this.index < this.source.length &&
      /[A-Za-z0-9_.$!]/.test(this.source[this.index])
    ) {
      this.index += 1
    }

    return this.source.slice(start, this.index)
  }

  private readQuotedSheetReference() {
    this.index += 1
    let sheetName = ''

    while (this.index < this.source.length && this.source[this.index] !== "'") {
      sheetName += this.source[this.index]
      this.index += 1
    }

    if (this.source[this.index] === "'") this.index += 1
    if (this.source[this.index] !== '!') return sheetName
    this.index += 1

    const start = this.index
    while (
      this.index < this.source.length &&
      /[A-Za-z0-9_.$]/.test(this.source[this.index])
    ) {
      this.index += 1
    }

    return sheetName + '!' + this.source.slice(start, this.index)
  }
}

class Parser {
  private index = 0

  constructor(
    private readonly tokens: Token[],
    private readonly sheet: SheetData,
    private readonly stack: Set<string>,
    private readonly sheets: SheetData[],
  ) {}

  parse(): FormulaValue {
    return this.parseComparison()
  }

  private current() {
    return this.tokens[this.index]
  }

  private next() {
    const token = this.tokens[this.index]
    this.index += 1
    return token
  }

  private match(type: TokenType, value?: string) {
    const token = this.current()
    if (token.type !== type) return false
    if (value !== undefined && token.value.toUpperCase() !== value.toUpperCase()) return false
    this.index += 1
    return true
  }

  private parseComparison(): FormulaValue {
    let left = this.parseConcat()

    while (
      this.current().type === 'operator' &&
      ['=', '<>', '!=', '>', '<', '>=', '<='].includes(this.current().value)
    ) {
      const operator = this.next().value
      const right = this.parseConcat()
      left = compare(left, right, operator)
    }

    return left
  }

  private parseConcat(): FormulaValue {
    let value = this.parseAdditive()

    while (this.match('operator', '&')) {
      value = toText(value) + toText(this.parseAdditive())
    }

    return value
  }

  private parseAdditive(): FormulaValue {
    let value = this.parseMultiplicative()

    while (
      this.current().type === 'operator' &&
      ['+', '-'].includes(this.current().value)
    ) {
      const operator = this.next().value
      const right = this.parseMultiplicative()
      value = operator === '+'
        ? toNumber(value) + toNumber(right)
        : toNumber(value) - toNumber(right)
    }

    return value
  }

  private parseMultiplicative(): FormulaValue {
    let value = this.parsePower()

    while (
      this.current().type === 'operator' &&
      ['*', '/'].includes(this.current().value)
    ) {
      const operator = this.next().value
      const right = this.parsePower()
      const divisor = toNumber(right)

      value = operator === '*'
        ? toNumber(value) * divisor
        : divisor === 0
          ? '#DIV/0!'
          : toNumber(value) / divisor
    }

    return value
  }

  private parsePower(): FormulaValue {
    let value = this.parseUnary()

    if (this.match('operator', '^')) {
      value = toNumber(value) ** toNumber(this.parsePower())
    }

    return value
  }

  private parseUnary(): FormulaValue {
    if (this.match('operator', '+')) return toNumber(this.parseUnary())
    if (this.match('operator', '-')) return -toNumber(this.parseUnary())

    let value = this.parsePrimary()

    while (this.match('operator', '%')) {
      value = toNumber(value) / 100
    }

    return value
  }

  private parsePrimary(): FormulaValue {
    const token = this.current()

    if (token.type === 'number') {
      this.next()
      return Number(token.value)
    }

    if (token.type === 'string') {
      this.next()
      return token.value
    }

    if (this.match('lparen')) {
      const value = this.parseComparison()
      this.match('rparen')
      return value
    }

    if (token.type === 'identifier') {
      this.next()
      const identifier = token.value
      const upper = identifier.toUpperCase()

      if (upper === 'TRUE') return true
      if (upper === 'FALSE') return false

      if (this.match('lparen')) {
        const args: FormulaValue[] = []

        if (!this.match('rparen')) {
          do {
            args.push(this.parseComparison())
          } while (this.match('comma'))

          this.match('rparen')
        }

        return this.callFunction(upper, args)
      }

      const crossSheet = this.crossSheetReference(identifier)
      if (crossSheet) {
        if (this.match('colon')) {
          const end = this.next()
          if (end.type !== 'identifier') return []
          const endCrossSheet = this.crossSheetReference(end.value)
          const endAddress = endCrossSheet ? endCrossSheet.address : end.value
          if (!this.isCellReference(endAddress)) return []
          if (endCrossSheet && endCrossSheet.sheet.id !== crossSheet.sheet.id) return '#REF!'
          return this.rangeValues(crossSheet.address, endAddress, crossSheet.sheet)
        }

        return this.cellValue(crossSheet.address, crossSheet.sheet)
      }

      if (this.isCellReference(identifier)) {
        if (this.match('colon')) {
          const end = this.next()
          if (end.type === 'identifier' && this.isCellReference(end.value)) {
            return this.rangeValues(identifier, end.value)
          }
          return []
        }

        return this.cellValue(identifier)
      }

      const namedRange = this.sheet.namedRanges?.[upper]
      if (namedRange) {
        const values: FormulaValue[] = []
        for (let row = namedRange.top; row <= namedRange.bottom; row += 1) {
          for (let col = namedRange.left; col <= namedRange.right; col += 1) {
            values.push(this.cellByPoint(row, col))
          }
        }
        return values
      }

      return 0
    }

    this.next()
    return 0
  }

  private isCellReference(value: string) {
    return /^\$?[A-Z]+\$?\d+$/i.test(value)
  }

  private crossSheetReference(value: string) {
    const bang = value.lastIndexOf('!')
    if (bang <= 0) return null
    const sheetName = value.slice(0, bang)
    const address = value.slice(bang + 1)
    if (!this.isCellReference(address)) return null
    const targetSheet = this.sheets.find((item) => item.name.toLowerCase() === sheetName.toLowerCase())
    if (!targetSheet) return null
    return { sheet: targetSheet, address }
  }

  private normalizedAddress(value: string) {
    return value.replace(/\$/g, '').toUpperCase()
  }

  private cellValue(address: string, targetSheet = this.sheet): FormulaValue {
    const normalized = this.normalizedAddress(address)
    const point = parseAddress(normalized)
    if (!point) return 0

    const cellAddress = cellKey(point.row, point.col)
    const stackKey = targetSheet.id + ':' + cellAddress
    if (this.stack.has(stackKey)) return '#CIRC!'

    const raw = targetSheet.cells[cellAddress]?.value ?? ''
    if (!raw) return ''

    if (raw.startsWith('=')) {
      const nextStack = new Set(this.stack)
      nextStack.add(stackKey)
      return evaluateFormula(raw, targetSheet, nextStack, this.sheets)
    }

    const numeric = Number(raw.replace(/[$,%]/g, ''))
    return Number.isFinite(numeric) && raw.trim() !== '' ? numeric : raw
  }

  private rangeValues(startAddress: string, endAddress: string, targetSheet = this.sheet): FormulaValue[] {
    const start = parseAddress(this.normalizedAddress(startAddress))
    const end = parseAddress(this.normalizedAddress(endAddress))
    if (!start || !end) return []

    const values: FormulaValue[] = []

    for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row += 1) {
      for (let col = Math.min(start.col, end.col); col <= Math.max(start.col, end.col); col += 1) {
        values.push(this.cellByPoint(row, col, targetSheet))
      }
    }

    return values
  }

  private cellByPoint(row: number, col: number, targetSheet = this.sheet): FormulaValue {
    const cellAddress = cellKey(row, col)
    const stackKey = targetSheet.id + ':' + cellAddress
    if (this.stack.has(stackKey)) return '#CIRC!'

    const raw = targetSheet.cells[cellAddress]?.value ?? ''
    if (!raw) return ''

    if (raw.startsWith('=')) {
      const nextStack = new Set(this.stack)
      nextStack.add(stackKey)
      return evaluateFormula(raw, targetSheet, nextStack, this.sheets)
    }

    const numeric = Number(raw.replace(/[$,%]/g, ''))
    return Number.isFinite(numeric) && raw.trim() !== '' ? numeric : raw
  }

  private callFunction(name: string, args: FormulaValue[]): FormulaValue {
    const values = flatten(args)
    const nums = numericValues(args)

    if (name === 'SUM') return nums.reduce((sum, value) => sum + value, 0)
    if (name === 'AVERAGE' || name === 'AVG') return nums.length ? nums.reduce((sum, value) => sum + value, 0) / nums.length : 0
    if (name === 'MIN') return nums.length ? Math.min(...nums) : 0
    if (name === 'MAX') return nums.length ? Math.max(...nums) : 0
    if (name === 'COUNT') return nums.length
    if (name === 'COUNTA') return values.filter((value) => toText(value) !== '').length
    if (name === 'PRODUCT') return nums.length ? nums.reduce((product, value) => product * value, 1) : 0

    if (name === 'COUNTIF') {
      const source = asList(args[0])
      const criterion = args[1] ?? ''
      return source.filter((value) => matchesCriteria(value, criterion)).length
    }

    if (name === 'COUNTIFS') {
      const pairs: Array<{ source: FormulaValue[]; criterion: FormulaValue }> = []
      for (let index = 0; index + 1 < args.length; index += 2) {
        pairs.push({
          source: asList(args[index]),
          criterion: args[index + 1] ?? '',
        })
      }

      const length = Math.max(0, ...pairs.map((pair) => pair.source.length))
      let count = 0
      for (let row = 0; row < length; row += 1) {
        if (pairs.every((pair) => matchesCriteria(pair.source[row] ?? '', pair.criterion))) count += 1
      }
      return count
    }

    if (name === 'SUMIF' || name === 'AVERAGEIF') {
      const source = asList(args[0])
      const criterion = args[1] ?? ''
      const sumSource = asList(args[2] ?? args[0])
      const matches: number[] = []

      source.forEach((value, index) => {
        if (!matchesCriteria(value, criterion)) return
        const candidate = sumSource[index] ?? 0
        if (isNumeric(candidate)) matches.push(toNumber(candidate))
      })

      if (name === 'AVERAGEIF') {
        return matches.length
          ? matches.reduce((sum, value) => sum + value, 0) / matches.length
          : 0
      }

      return matches.reduce((sum, value) => sum + value, 0)
    }

    if (name === 'SUMIFS' || name === 'AVERAGEIFS') {
      const sumSource = asList(args[0])
      const pairs: Array<{ source: FormulaValue[]; criterion: FormulaValue }> = []

      for (let index = 1; index + 1 < args.length; index += 2) {
        pairs.push({
          source: asList(args[index]),
          criterion: args[index + 1] ?? '',
        })
      }

      const matches: number[] = []
      sumSource.forEach((candidate, row) => {
        if (!pairs.every((pair) => matchesCriteria(pair.source[row] ?? '', pair.criterion))) return
        if (isNumeric(candidate)) matches.push(toNumber(candidate))
      })

      if (name === 'AVERAGEIFS') {
        return matches.length
          ? matches.reduce((sum, value) => sum + value, 0) / matches.length
          : 0
      }

      return matches.reduce((sum, value) => sum + value, 0)
    }

    if (name === 'XLOOKUP') {
      const lookupValue = args[0] ?? ''
      const lookupArray = asList(args[1])
      const returnArray = asList(args[2])
      const matchIndex = lookupArray.findIndex((value) => compare(value, lookupValue, '='))
      if (matchIndex === -1) return args[3] ?? '#N/A'
      return returnArray[matchIndex] ?? args[3] ?? '#N/A'
    }

    if (name === 'MATCH') {
      const lookupValue = args[0] ?? ''
      const lookupArray = asList(args[1])
      const matchIndex = lookupArray.findIndex((value) => compare(value, lookupValue, '='))
      return matchIndex === -1 ? '#N/A' : matchIndex + 1
    }

    if (name === 'XMATCH') {
      const lookupValue = args[0] ?? ''
      const lookupArray = asList(args[1])
      const matchIndex = lookupArray.findIndex((value) => compare(value, lookupValue, '='))
      return matchIndex === -1 ? '#N/A' : matchIndex + 1
    }

    if (name === 'INDEX') {
      const source = asList(args[0])
      const index = Math.trunc(toNumber(args[1] ?? 1)) - 1
      if (index < 0 || index >= source.length) return '#REF!'
      return source[index]
    }

    if (name === 'IFERROR') {
      const candidate = args[0] ?? ''
      return isErrorValue(candidate) ? args[1] ?? '' : candidate
    }

    if (name === 'TEXTJOIN') {
      const delimiter = toText(args[0] ?? '')
      const ignoreEmpty = truthy(args[1] ?? false)
      const parts = flatten(args.slice(2))
        .map(toText)
        .filter((value) => !ignoreEmpty || value !== '')
      return parts.join(delimiter)
    }

    if (name === 'TODAY') {
      const now = new Date()
      return excelSerialFromDate(new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate())))
    }

    if (name === 'NOW') return excelSerialFromDate(new Date())

    if (name === 'DATE') {
      const year = Math.trunc(toNumber(args[0] ?? 1900))
      const month = Math.trunc(toNumber(args[1] ?? 1))
      const day = Math.trunc(toNumber(args[2] ?? 1))
      return excelSerialFromDate(new Date(Date.UTC(year, month - 1, day)))
    }

    if (['YEAR', 'MONTH', 'DAY', 'HOUR', 'MINUTE', 'SECOND'].includes(name)) {
      const date = dateFromFormulaValue(args[0] ?? '')
      if (!date) return '#VALUE!'
      if (name === 'YEAR') return date.getUTCFullYear()
      if (name === 'MONTH') return date.getUTCMonth() + 1
      if (name === 'DAY') return date.getUTCDate()
      if (name === 'HOUR') return date.getUTCHours()
      if (name === 'MINUTE') return date.getUTCMinutes()
      return date.getUTCSeconds()
    }

    if (name === 'EDATE' || name === 'EOMONTH') {
      const date = dateFromFormulaValue(args[0] ?? '')
      if (!date) return '#VALUE!'
      const months = Math.trunc(toNumber(args[1] ?? 0))
      const shifted = new Date(Date.UTC(
        date.getUTCFullYear(),
        date.getUTCMonth() + months + (name === 'EOMONTH' ? 1 : 0),
        name === 'EOMONTH' ? 0 : date.getUTCDate(),
        date.getUTCHours(),
        date.getUTCMinutes(),
        date.getUTCSeconds(),
      ))
      return excelSerialFromDate(shifted)
    }

    if (name === 'MEDIAN') {
      if (!nums.length) return 0
      const sorted = [...nums].sort((a, b) => a - b)
      const middle = Math.floor(sorted.length / 2)
      return sorted.length % 2
        ? sorted[middle]
        : (sorted[middle - 1] + sorted[middle]) / 2
    }

    if (name === 'ROUND') {
      const digits = Math.trunc(toNumber(args[1] ?? 0))
      const factor = 10 ** digits
      return Math.round(toNumber(args[0] ?? 0) * factor) / factor
    }

    if (name === 'ROUNDUP') {
      const digits = Math.trunc(toNumber(args[1] ?? 0))
      const factor = 10 ** digits
      const value = toNumber(args[0] ?? 0) * factor
      return (value >= 0 ? Math.ceil(value) : Math.floor(value)) / factor
    }

    if (name === 'ROUNDDOWN') {
      const digits = Math.trunc(toNumber(args[1] ?? 0))
      const factor = 10 ** digits
      const value = toNumber(args[0] ?? 0) * factor
      return (value >= 0 ? Math.floor(value) : Math.ceil(value)) / factor
    }

    if (name === 'ABS') return Math.abs(toNumber(args[0] ?? 0))
    if (name === 'SQRT') return Math.sqrt(Math.max(0, toNumber(args[0] ?? 0)))
    if (name === 'POWER') return toNumber(args[0] ?? 0) ** toNumber(args[1] ?? 0)

    if (name === 'MOD') {
      const divisor = toNumber(args[1] ?? 1)
      return divisor === 0 ? '#DIV/0!' : toNumber(args[0] ?? 0) % divisor
    }

    if (name === 'IF') {
      return truthy(args[0] ?? false)
        ? args[1] ?? true
        : args[2] ?? false
    }

    if (name === 'AND') return values.every(truthy)
    if (name === 'OR') return values.some(truthy)
    if (name === 'NOT') return !truthy(args[0] ?? false)

    if (name === 'LEN') return toText(args[0] ?? '').length
    if (name === 'LOWER') return toText(args[0] ?? '').toLowerCase()
    if (name === 'UPPER') return toText(args[0] ?? '').toUpperCase()
    if (name === 'TRIM') return toText(args[0] ?? '').trim().replace(/\s+/g, ' ')
    if (name === 'LEFT') return toText(args[0] ?? '').slice(0, Math.max(0, Math.trunc(toNumber(args[1] ?? 1))))
    if (name === 'RIGHT') {
      const count = Math.max(0, Math.trunc(toNumber(args[1] ?? 1)))
      return toText(args[0] ?? '').slice(-count)
    }

    if (name === 'MID') {
      const text = toText(args[0] ?? '')
      const start = Math.max(0, Math.trunc(toNumber(args[1] ?? 1)) - 1)
      const count = Math.max(0, Math.trunc(toNumber(args[2] ?? 0)))
      return text.slice(start, start + count)
    }

    if (name === 'CONCAT' || name === 'CONCATENATE') return values.map(toText).join('')

    if (name === 'SUMPRODUCT') {
      const arrays = args.map(asList)
      const length = Math.max(0, ...arrays.map((array) => array.length))
      let total = 0
      for (let index = 0; index < length; index += 1) {
        let product = 1
        for (const array of arrays) product *= toNumber(array[index] ?? 0)
        total += product
      }
      return total
    }

    if (name === 'LARGE' || name === 'SMALL') {
      const source = numericValues([args[0] ?? []]).sort((a, b) => a - b)
      const k = Math.trunc(toNumber(args[1] ?? 1))
      if (k < 1 || k > source.length) return '#NUM!'
      return name === 'LARGE' ? source[source.length - k] : source[k - 1]
    }

    if (name === 'RANK' || name === 'RANK.EQ') {
      const value = toNumber(args[0] ?? 0)
      const source = numericValues([args[1] ?? []])
      const descending = toNumber(args[2] ?? 0) === 0
      const sorted = [...source].sort((a, b) => descending ? b - a : a - b)
      const index = sorted.findIndex((candidate) => candidate === value)
      return index === -1 ? '#N/A' : index + 1
    }

    if (name === 'STDEV' || name === 'STDEV.S' || name === 'STDEV.P') {
      if (!nums.length) return '#DIV/0!'
      const mean = nums.reduce((sum, value) => sum + value, 0) / nums.length
      const divisor = name === 'STDEV.P' ? nums.length : nums.length - 1
      if (divisor <= 0) return '#DIV/0!'
      const variance = nums.reduce((sum, value) => sum + (value - mean) ** 2, 0) / divisor
      return Math.sqrt(variance)
    }

    if (name === 'VAR' || name === 'VAR.S' || name === 'VAR.P') {
      if (!nums.length) return '#DIV/0!'
      const mean = nums.reduce((sum, value) => sum + value, 0) / nums.length
      const divisor = name === 'VAR.P' ? nums.length : nums.length - 1
      if (divisor <= 0) return '#DIV/0!'
      return nums.reduce((sum, value) => sum + (value - mean) ** 2, 0) / divisor
    }

    if (name === 'PERCENTILE' || name === 'PERCENTILE.INC') {
      const source = numericValues([args[0] ?? []]).sort((a, b) => a - b)
      const percentile = toNumber(args[1] ?? 0)
      if (!source.length || percentile < 0 || percentile > 1) return '#NUM!'
      const position = (source.length - 1) * percentile
      const lower = Math.floor(position)
      const upper = Math.ceil(position)
      if (lower === upper) return source[lower]
      const ratio = position - lower
      return source[lower] + (source[upper] - source[lower]) * ratio
    }

    if (name === 'INT') return Math.floor(toNumber(args[0] ?? 0))
    if (name === 'SIGN') return Math.sign(toNumber(args[0] ?? 0))
    if (name === 'PI') return Math.PI
    if (name === 'EXP') return Math.exp(toNumber(args[0] ?? 0))
    if (name === 'LN') return Math.log(toNumber(args[0] ?? 0))
    if (name === 'LOG10') return Math.log10(toNumber(args[0] ?? 0))
    if (name === 'LOG') {
      const value = toNumber(args[0] ?? 0)
      const base = toNumber(args[1] ?? 10)
      return Math.log(value) / Math.log(base)
    }

    if (name === 'CEILING' || name === 'CEILING.MATH') {
      const value = toNumber(args[0] ?? 0)
      const significance = Math.abs(toNumber(args[1] ?? 1)) || 1
      return Math.ceil(value / significance) * significance
    }

    if (name === 'FLOOR' || name === 'FLOOR.MATH') {
      const value = toNumber(args[0] ?? 0)
      const significance = Math.abs(toNumber(args[1] ?? 1)) || 1
      return Math.floor(value / significance) * significance
    }

    if (name === 'ISNUMBER') return isNumeric(args[0] ?? '')
    if (name === 'ISTEXT') return typeof (args[0] ?? '') === 'string' && !isErrorValue(args[0] ?? '')
    if (name === 'ISERROR') return isErrorValue(args[0] ?? '')
    if (name === 'ISBLANK') return toText(args[0] ?? '') === ''
    if (name === 'VALUE') {
      const source = toText(args[0] ?? '').replace(/[$,%]/g, '')
      const parsed = Number(source)
      return Number.isFinite(parsed) ? parsed : '#VALUE!'
    }

    if (name === 'TIME') {
      const hour = toNumber(args[0] ?? 0)
      const minute = toNumber(args[1] ?? 0)
      const second = toNumber(args[2] ?? 0)
      return (hour * 3600 + minute * 60 + second) / 86400
    }

    if (name === 'DATEVALUE') {
      const date = new Date(toText(args[0] ?? ''))
      return Number.isNaN(date.getTime()) ? '#VALUE!' : excelSerialFromDate(date)
    }

    if (name === 'TIMEVALUE') {
      const text = toText(args[0] ?? '')
      const date = new Date('1970-01-01T' + text)
      if (Number.isNaN(date.getTime())) return '#VALUE!'
      return (date.getHours() * 3600 + date.getMinutes() * 60 + date.getSeconds()) / 86400
    }

    if (name === 'WEEKDAY') {
      const date = dateFromFormulaValue(args[0] ?? '')
      if (!date) return '#VALUE!'
      const returnType = Math.trunc(toNumber(args[1] ?? 1))
      const day = date.getUTCDay()
      if (returnType === 2) return day === 0 ? 7 : day
      return day + 1
    }

    if (name === 'DAYS') {
      const end = dateFromFormulaValue(args[0] ?? '')
      const start = dateFromFormulaValue(args[1] ?? '')
      if (!end || !start) return '#VALUE!'
      return Math.round((end.getTime() - start.getTime()) / DAY_MS)
    }

    if (name === 'PMT') {
      const rate = toNumber(args[0] ?? 0)
      const periods = toNumber(args[1] ?? 0)
      const presentValue = toNumber(args[2] ?? 0)
      const futureValue = toNumber(args[3] ?? 0)
      const due = toNumber(args[4] ?? 0)
      if (periods === 0) return '#DIV/0!'
      if (rate === 0) return -(presentValue + futureValue) / periods
      const factor = (1 + rate) ** periods
      return -(rate * (presentValue * factor + futureValue)) / ((1 + rate * due) * (factor - 1))
    }

    if (name === 'FV') {
      const rate = toNumber(args[0] ?? 0)
      const periods = toNumber(args[1] ?? 0)
      const payment = toNumber(args[2] ?? 0)
      const presentValue = toNumber(args[3] ?? 0)
      const due = toNumber(args[4] ?? 0)
      if (rate === 0) return -(presentValue + payment * periods)
      const factor = (1 + rate) ** periods
      return -(presentValue * factor + payment * (1 + rate * due) * (factor - 1) / rate)
    }

    if (name === 'PV') {
      const rate = toNumber(args[0] ?? 0)
      const periods = toNumber(args[1] ?? 0)
      const payment = toNumber(args[2] ?? 0)
      const futureValue = toNumber(args[3] ?? 0)
      const due = toNumber(args[4] ?? 0)
      if (rate === 0) return -futureValue - payment * periods
      const factor = (1 + rate) ** periods
      return -(futureValue + payment * (1 + rate * due) * (factor - 1) / rate) / factor
    }

    return `#NAME? ${name}`
  }
}

export function evaluateFormula(
  raw: string,
  sheet: SheetData,
  stack = new Set<string>(),
  sheets: SheetData[] = [sheet],
): FormulaResult {
  if (!raw.startsWith('=')) return raw

  try {
    const tokens = new Lexer(raw.slice(1)).tokenize()
    const result = new Parser(tokens, sheet, stack, sheets).parse()

    if (Array.isArray(result)) {
      const first = result[0]
      if (Array.isArray(first)) return ''
      return first ?? ''
    }

    return result
  } catch {
    return '#ERROR!'
  }
}

export function displayValue(raw: string, sheet: SheetData, sheets: SheetData[] = [sheet]): string {
  if (!raw.startsWith('=')) return raw

  const result = evaluateFormula(raw, sheet, new Set<string>(), sheets)

  if (typeof result === 'boolean') return result ? 'TRUE' : 'FALSE'

  if (typeof result === 'number') {
    if (!Number.isFinite(result)) return '#NUM!'
    return String(Math.round((result + Number.EPSILON) * 1e10) / 1e10)
  }

  return String(result)
}
