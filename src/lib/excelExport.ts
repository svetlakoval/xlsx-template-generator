import ExcelJS from 'exceljs'
import { generateValue, type RowGenerationContext } from './generateData'
import type { ColumnConfig, GenerationConfig } from './types'

const cloneValue = <T,>(value: T | undefined): T | undefined => {
  if (value === undefined) return undefined
  return JSON.parse(JSON.stringify(value)) as T
}

const copyColumnFormatting = (sourceWorksheet: ExcelJS.Worksheet, targetWorksheet: ExcelJS.Worksheet, maxColumn: number) => {
  for (let columnIndex = 1; columnIndex <= maxColumn; columnIndex += 1) {
    const sourceColumn = sourceWorksheet.getColumn(columnIndex)
    const targetColumn = targetWorksheet.getColumn(columnIndex)

    targetColumn.width = sourceColumn.width
    targetColumn.hidden = sourceColumn.hidden
    targetColumn.outlineLevel = sourceColumn.outlineLevel
    targetColumn.style = cloneValue(sourceColumn.style) ?? {}
  }
}

const copyCellFormatting = (sourceCell: ExcelJS.Cell, targetCell: ExcelJS.Cell) => {
  targetCell.style = cloneValue(sourceCell.style) ?? {}
  targetCell.numFmt = sourceCell.numFmt

  const alignment = cloneValue(sourceCell.alignment)
  const border = cloneValue(sourceCell.border)
  const fill = cloneValue(sourceCell.fill)
  const font = cloneValue(sourceCell.font)
  const protection = cloneValue(sourceCell.protection)
  const note = cloneValue(sourceCell.note)

  if (alignment) targetCell.alignment = alignment
  if (border) targetCell.border = border
  if (fill) targetCell.fill = fill
  if (font) targetCell.font = font
  if (protection) targetCell.protection = protection
  if (note) targetCell.note = note
}

const copyRowFormatting = (sourceRow: ExcelJS.Row, targetRow: ExcelJS.Row, maxColumn: number) => {
  targetRow.height = sourceRow.height

  for (let columnIndex = 1; columnIndex <= maxColumn; columnIndex += 1) {
    copyCellFormatting(sourceRow.getCell(columnIndex), targetRow.getCell(columnIndex))
  }
}

const copyRowWithValues = (sourceRow: ExcelJS.Row, targetRow: ExcelJS.Row, maxColumn: number) => {
  copyRowFormatting(sourceRow, targetRow, maxColumn)

  for (let columnIndex = 1; columnIndex <= maxColumn; columnIndex += 1) {
    targetRow.getCell(columnIndex).value = cloneValue(sourceRow.getCell(columnIndex).value) as ExcelJS.CellValue
  }

  targetRow.commit()
}

const copyHeaderArea = (sourceWorksheet: ExcelJS.Worksheet, targetWorksheet: ExcelJS.Worksheet, headerRowNumber: number, maxColumn: number) => {
  for (let rowNumber = 1; rowNumber <= headerRowNumber; rowNumber += 1) {
    copyRowWithValues(sourceWorksheet.getRow(rowNumber), targetWorksheet.getRow(rowNumber), maxColumn)
  }
}

const copyMergedCells = (sourceWorksheet: ExcelJS.Worksheet, targetWorksheet: ExcelJS.Worksheet, headerRowNumber: number) => {
  const merges = ((sourceWorksheet as unknown as { model?: { merges?: string[] } }).model?.merges ?? [])

  merges.forEach((range) => {
    const rows = range.match(/\d+/g)?.map(Number) ?? []
    if (rows.length && Math.max(...rows) <= headerRowNumber) {
      targetWorksheet.mergeCells(range)
    }
  })
}

const setCellValue = (cell: ExcelJS.Cell, column: ColumnConfig, value: unknown) => {
  if (column.type === 'formula' && typeof value === 'string' && value.startsWith('=')) {
    cell.value = { formula: value.slice(1) }
    return
  }

  cell.value = value as ExcelJS.CellValue
}

export const generateWorkbook = async (templateWorkbook: ExcelJS.Workbook, config: GenerationConfig) => {
  const sourceWorksheet = templateWorkbook.getWorksheet(config.sheetName)
  if (!sourceWorksheet) throw new Error(`Лист «${config.sheetName}» не найден`)

  const outputWorkbook = new ExcelJS.Workbook()
  outputWorkbook.creator = templateWorkbook.creator
  outputWorkbook.lastModifiedBy = templateWorkbook.lastModifiedBy
  outputWorkbook.created = templateWorkbook.created
  outputWorkbook.modified = new Date()

  const worksheet = outputWorkbook.addWorksheet(sourceWorksheet.name, {
    properties: cloneValue(sourceWorksheet.properties),
    pageSetup: cloneValue(sourceWorksheet.pageSetup),
    views: cloneValue(sourceWorksheet.views),
  })

  const dataStartRowNumber = config.headerRowNumber + 1
  const maxColumn = Math.max(sourceWorksheet.columnCount, ...config.columns.map((column) => column.index))
  const templateRowNumber = dataStartRowNumber <= sourceWorksheet.rowCount ? dataStartRowNumber : config.headerRowNumber
  const templateRow = sourceWorksheet.getRow(templateRowNumber)

  copyColumnFormatting(sourceWorksheet, worksheet, maxColumn)
  copyHeaderArea(sourceWorksheet, worksheet, config.headerRowNumber, maxColumn)
  copyMergedCells(sourceWorksheet, worksheet, config.headerRowNumber)

  const uniqueStore = new Map<string, Set<string>>()
  const ignoredRowNumbers = new Set(config.ignoredRowNumbers)

  for (let rowIndex = 0; rowIndex < config.rowsCount; rowIndex += 1) {
    const rowNumber = dataStartRowNumber + rowIndex
    const row = worksheet.getRow(rowNumber)

    if (ignoredRowNumbers.has(rowNumber)) {
      copyRowWithValues(sourceWorksheet.getRow(rowNumber), row, maxColumn)
      continue
    }

    copyRowFormatting(templateRow, row, maxColumn)

    const rowContext: RowGenerationContext = {}

    for (const column of config.columns) {
      if (!column.enabled) continue
      const cell = row.getCell(column.index)
      const value = generateValue(column, rowIndex, uniqueStore, rowContext)
      setCellValue(cell, column, value)
    }

    row.commit()
  }

  return outputWorkbook
}

export const downloadWorkbook = async (workbook: ExcelJS.Workbook, fileName: string) => {
  const buffer = await workbook.xlsx.writeBuffer()
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = fileName
  document.body.appendChild(link)
  link.click()
  link.remove()
  URL.revokeObjectURL(url)
}
