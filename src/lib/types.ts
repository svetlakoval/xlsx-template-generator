export type ColumnType =
  | 'text'
  | 'number'
  | 'date'
  | 'email'
  | 'phone'
  | 'fullName'
  | 'firstName'
  | 'lastName'
  | 'middleName'
  | 'gender'
  | 'company'
  | 'city'
  | 'country'
  | 'address'
  | 'blankNumber'
  | 'issuer'
  | 'uuid'
  | 'boolean'
  | 'enum'
  | 'formula'

export type ColumnConfig = {
  id: string
  index: number
  header: string
  type: ColumnType
  enabled: boolean
  unique: boolean
  samples: string[]
  enumValues: string[]
  min?: number
  max?: number
  dateFrom?: string
  dateTo?: string
  formula?: string
}

export type TemplateAnalysis = {
  sheetNames: string[]
  selectedSheetName: string
  headerRowNumber: number
  sampleRowCount: number
  columns: ColumnConfig[]
}

export type GenerationConfig = {
  sheetName: string
  headerRowNumber: number
  ignoredRowNumbers: number[]
  rowsCount: number
  columns: ColumnConfig[]
}

export type PreviewRow = Record<string, string>
