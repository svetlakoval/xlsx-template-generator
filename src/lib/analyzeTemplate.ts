import ExcelJS from 'exceljs'
import type { ColumnConfig, ColumnType, TemplateAnalysis } from './types'

const MAX_SCAN_ROWS = 30
const MAX_SAMPLE_ROWS = 10

const getCellPlainValue = (cell: ExcelJS.Cell): unknown => {
  const value = cell.value

  if (value === null || value === undefined) return ''
  if (value instanceof Date) return value
  if (typeof value === 'object') {
    if ('formula' in value) return `=${String(value.formula ?? '')}`
    if ('result' in value) return value.result ?? ''
    if ('text' in value) return value.text ?? ''
    if ('richText' in value && Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text ?? '').join('')
    }
    if ('hyperlink' in value && 'text' in value) return value.text ?? value.hyperlink ?? ''
  }

  return value
}

const padDatePart = (value: number) => String(value).padStart(2, '0')
const formatDate = (date: Date) => `${padDatePart(date.getDate())}.${padDatePart(date.getMonth() + 1)}.${date.getFullYear()}`

const parseDateValue = (value: unknown) => {
  if (value instanceof Date) return value
  if (typeof value !== 'string') return null

  const trimmed = value.trim()
  const ruDateMatch = trimmed.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/)
  if (ruDateMatch) {
    const [, day, month, year] = ruDateMatch
    const parsed = new Date(Number(year), Number(month) - 1, Number(day))
    return Number.isNaN(parsed.getTime()) ? null : parsed
  }

  const parsed = new Date(trimmed)
  return Number.isNaN(parsed.getTime()) ? null : parsed
}

export const valueToString = (value: unknown) => {
  if (value === null || value === undefined) return ''
  if (value instanceof Date) return formatDate(value)
  return String(value).trim()
}

const normalize = (value: string) => value.trim().toLowerCase()

const isLikelyHeaderText = (value: string) => {
  if (!value) return false
  if (/^\d+([.,]\d+)?$/.test(value)) return false
  if (/^\d{4}-\d{2}-\d{2}/.test(value)) return false
  return /[a-zа-яё_ -]/i.test(value)
}

const detectHeaderRowNumber = (worksheet: ExcelJS.Worksheet) => {
  let bestRowNumber = 1
  let bestScore = -Infinity
  const rowsToScan = Math.min(worksheet.rowCount || MAX_SCAN_ROWS, MAX_SCAN_ROWS)

  for (let rowNumber = 1; rowNumber <= rowsToScan; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber)
    const values = row.values as unknown[]
    const filledValues = values.slice(1).map(valueToString).filter(Boolean)
    const textValues = filledValues.filter(isLikelyHeaderText)
    const uniqueValues = new Set(filledValues.map(normalize))
    const nextRow = worksheet.getRow(rowNumber + 1)
    const nextFilledCount = (nextRow.values as unknown[]).slice(1).map(valueToString).filter(Boolean).length

    const score = textValues.length * 3 + uniqueValues.size + Math.min(nextFilledCount, filledValues.length) - Math.abs(filledValues.length - uniqueValues.size) * 2

    if (filledValues.length >= 2 && textValues.length >= Math.ceil(filledValues.length / 2) && score > bestScore) {
      bestScore = score
      bestRowNumber = rowNumber
    }
  }

  return bestRowNumber
}

const looksLikeEmail = (value: string) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value)
const looksLikeUuid = (value: string) => /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value)
const looksLikePhone = (value: string) => /^\+?[\d\s().-]{7,}$/.test(value) && /\d{7,}/.test(value.replace(/\D/g, ''))
const looksLikeBoolean = (value: string) => /^(true|false|yes|no|да|нет|истина|ложь|0|1)$/i.test(value)
const looksLikeFormula = (value: string) => value.startsWith('=')

const looksLikeDate = (value: unknown) => {
  if (value instanceof Date) return true
  if (typeof value !== 'string') return false
  const trimmed = value.trim()
  if (!trimmed) return false
  if (/^\d+$/.test(trimmed)) return false
  if (!/[./-]/.test(trimmed)) return false
  const parsed = parseDateValue(trimmed)
  return Boolean(parsed)
}

const looksLikeNumber = (value: string) => {
  if (!value) return false
  if (/^0\d+/.test(value)) return false
  return /^-?\d+([.,]\d+)?$/.test(value)
}

const headerHints: Array<[ColumnType, RegExp]> = [
  ['email', /(^|[_\s-])(email|e-mail|mail|почта|емейл)([_\s-]|$)/i],
  ['phone', /(phone|tel|mobile|телефон)/i],
  ['lastName', /(last.?name|surname|family.?name|фамил)/i],
  ['middleName', /(middle.?name|patronymic|отчеств)/i],
  ['firstName', /(^|[_\s-])(first.?name|given.?name|имя)([_\s-]|$)/i],
  ['fullName', /(fio|full.?name|name|фио|ф\.и\.о|клиент|пользователь)/i],
  ['gender', /(^|[_\s-])(gender|sex|пол)([_\s-]|$)/i],
  ['company', /(company|organization|org|компан|организац)/i],
  ['country', /(citizenship|nationality|country|гражданств|страна|подданств)/i],
  ['address', /(address|адрес|регистрац|пропис)/i],
  ['issuer', /(issued.?by|issuer|кем.?выдан|выдан|орган.?выдач|подразделен)/i],
  ['blankNumber', /(blank.?number|номер.?бланк|бланк)/i],
  ['city', /(city|город|населен)/i],
  ['uuid', /(uuid|guid)/i],
  ['date', /(date|created|updated|дата|создан|обновлен)/i],
  ['boolean', /(active|enabled|deleted|is_|актив|флаг|да\/нет)/i],
  ['number', /(id|ид|код|count|qty|amount|price|sum|total|номер|количество|цена|сумма)/i],
]

const semanticHeaderTypes: ColumnType[] = ['fullName', 'firstName', 'lastName', 'middleName', 'gender', 'country', 'address', 'blankNumber', 'issuer', 'city', 'company']

const inferColumnType = (header: string, rawSamples: unknown[]): ColumnType => {
  const samples = rawSamples.map((sample) => (sample instanceof Date ? sample : valueToString(sample))).filter((sample) => valueToString(sample).length > 0)
  const headerHint = headerHints.find(([, pattern]) => pattern.test(header))?.[0]

  if (!samples.length) return headerHint ?? 'text'
  if (headerHint && semanticHeaderTypes.includes(headerHint)) return headerHint
  if (samples.every((sample) => looksLikeFormula(valueToString(sample)))) return 'formula'
  if (samples.every((sample) => looksLikeUuid(valueToString(sample)))) return 'uuid'
  if (samples.every((sample) => looksLikeEmail(valueToString(sample)))) return 'email'
  if (samples.every((sample) => looksLikeBoolean(valueToString(sample)))) return 'boolean'
  if (samples.every((sample) => looksLikeDate(sample))) return 'date'
  if (samples.every((sample) => looksLikeNumber(valueToString(sample)))) return headerHint === 'phone' ? 'phone' : 'number'
  if (samples.every((sample) => looksLikePhone(valueToString(sample)))) return 'phone'

  const uniqueValues = new Set(samples.map((sample) => normalize(valueToString(sample))))
  if (uniqueValues.size > 1 && uniqueValues.size <= Math.min(8, samples.length)) return 'enum'

  return headerHint ?? 'text'
}

const inferNumberBounds = (samples: string[]) => {
  const numbers = samples.map((sample) => Number(sample.replace(',', '.'))).filter(Number.isFinite)
  if (!numbers.length) return {}
  return {
    min: Math.min(...numbers),
    max: Math.max(...numbers),
  }
}

const inferDateBounds = (samples: unknown[]) => {
  const dates = samples
    .map((sample) => {
      if (sample instanceof Date) return sample
      const parsed = parseDateValue(sample)
      return parsed
    })
    .filter((date): date is Date => Boolean(date))

  if (!dates.length) return {}
  const timestamps = dates.map((date) => date.getTime())
  return {
    dateFrom: new Date(Math.min(...timestamps)).toISOString().slice(0, 10),
    dateTo: new Date(Math.max(...timestamps)).toISOString().slice(0, 10),
  }
}

const getSampleRows = (worksheet: ExcelJS.Worksheet, headerRowNumber: number) => {
  const rows: ExcelJS.Row[] = []
  const lastRow = Math.min(worksheet.rowCount, headerRowNumber + MAX_SAMPLE_ROWS)

  for (let rowNumber = headerRowNumber + 1; rowNumber <= lastRow; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber)
    const hasValues = (row.values as unknown[]).slice(1).some((value) => valueToString(value).length > 0)
    if (hasValues) rows.push(row)
  }

  return rows
}

export const analyzeWorksheet = (workbook: ExcelJS.Workbook, sheetName: string, preferredHeaderRowNumber?: number): TemplateAnalysis => {
  const worksheet = workbook.getWorksheet(sheetName)
  if (!worksheet) {
    throw new Error(`Лист «${sheetName}» не найден`)
  }

  const headerRowNumber = preferredHeaderRowNumber ?? detectHeaderRowNumber(worksheet)
  const headerRow = worksheet.getRow(headerRowNumber)
  const sampleRows = getSampleRows(worksheet, headerRowNumber)
  const maxColumn = Math.max(headerRow.cellCount, worksheet.columnCount)
  const columns: ColumnConfig[] = []

  for (let columnIndex = 1; columnIndex <= maxColumn; columnIndex += 1) {
    const header = valueToString(getCellPlainValue(headerRow.getCell(columnIndex)))
    if (!header) continue

    const rawSamples = sampleRows.map((row) => getCellPlainValue(row.getCell(columnIndex))).filter((value) => valueToString(value).length > 0)
    const samples = rawSamples.map(valueToString)
    const type = inferColumnType(header, rawSamples)
    const numberBounds = type === 'number' ? inferNumberBounds(samples) : {}
    const dateBounds = type === 'date' ? inferDateBounds(rawSamples) : {}
    const enumValues = type === 'enum' ? Array.from(new Set(samples.map((sample) => sample.trim()).filter(Boolean))).slice(0, 20) : []
    const formulaSample = type === 'formula' ? samples.find((sample) => sample.startsWith('='))?.slice(1) : undefined

    columns.push({
      id: `${columnIndex}-${header}`,
      index: columnIndex,
      header,
      type,
      enabled: true,
      unique: ['uuid', 'email'].includes(type) || /(^id$|_id$|ид|код)/i.test(header),
      samples: samples.slice(0, 5),
      enumValues,
      formula: formulaSample,
      ...numberBounds,
      ...dateBounds,
    })
  }

  return {
    sheetNames: workbook.worksheets.map((sheet) => sheet.name),
    selectedSheetName: worksheet.name,
    headerRowNumber,
    sampleRowCount: sampleRows.length,
    columns,
  }
}

export const loadWorkbookFromFile = async (file: File) => {
  const buffer = await file.arrayBuffer()
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(buffer)
  return workbook
}
