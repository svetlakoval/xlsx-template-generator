import { useMemo, useState } from 'react'
import type ExcelJS from 'exceljs'
import { analyzeWorksheet, loadWorkbookFromFile } from './lib/analyzeTemplate'
import { downloadWorkbook, generateWorkbook } from './lib/excelExport'
import { generatePreviewRows } from './lib/generateData'
import type { ColumnConfig, ColumnType, TemplateAnalysis } from './lib/types'

const columnTypes: Array<{ value: ColumnType; label: string }> = [
  { value: 'text', label: 'Текст' },
  { value: 'number', label: 'Число' },
  { value: 'date', label: 'Дата' },
  { value: 'email', label: 'Email' },
  { value: 'phone', label: 'Телефон' },
  { value: 'fullName', label: 'ФИО' },
  { value: 'lastName', label: 'Фамилия' },
  { value: 'firstName', label: 'Имя' },
  { value: 'middleName', label: 'Отчество' },
  { value: 'gender', label: 'Пол' },
  { value: 'company', label: 'Компания' },
  { value: 'city', label: 'Город' },
  { value: 'country', label: 'Гражданство / страна' },
  { value: 'address', label: 'Адрес' },
  { value: 'blankNumber', label: 'Номер бланка' },
  { value: 'issuer', label: 'Кем выдан' },
  { value: 'uuid', label: 'UUID' },
  { value: 'boolean', label: 'Boolean' },
  { value: 'enum', label: 'Список' },
  { value: 'formula', label: 'Формула' },
]

const getGeneratedFileName = (fileName: string | null, rowsCount: number) => {
  const baseName = fileName?.replace(/\.xlsx$/i, '') || 'template'
  return `${baseName}-generated-${rowsCount}.xlsx`
}

const parseIgnoredRowNumbers = (value: string) => {
  const rowNumbers = new Set<number>()

  value
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean)
    .forEach((item) => {
      const rangeMatch = item.match(/^(\d+)\s*-\s*(\d+)$/)
      if (rangeMatch) {
        const from = Number(rangeMatch[1])
        const to = Number(rangeMatch[2])
        const min = Math.min(from, to)
        const max = Math.max(from, to)

        for (let rowNumber = min; rowNumber <= max; rowNumber += 1) {
          rowNumbers.add(rowNumber)
        }
        return
      }

      const rowNumber = Number(item)
      if (Number.isInteger(rowNumber) && rowNumber > 0) rowNumbers.add(rowNumber)
    })

  return Array.from(rowNumbers).sort((a, b) => a - b)
}

const updateColumnById = (
  columns: ColumnConfig[],
  id: string,
  updater: (column: ColumnConfig) => ColumnConfig,
) => columns.map((column) => (column.id === id ? updater(column) : column))

function App() {
  const [workbook, setWorkbook] = useState<ExcelJS.Workbook | null>(null)
  const [fileName, setFileName] = useState<string | null>(null)
  const [analysis, setAnalysis] = useState<TemplateAnalysis | null>(null)
  const [rowsCount, setRowsCount] = useState(100)
  const [ignoredRowsInput, setIgnoredRowsInput] = useState('')
  const [isLoading, setIsLoading] = useState(false)
  const [isGenerating, setIsGenerating] = useState(false)
  const [errorMessage, setErrorMessage] = useState('')

  const previewRows = useMemo(() => {
    if (!analysis) return []
    return generatePreviewRows(analysis.columns, Math.min(rowsCount, 5))
  }, [analysis, rowsCount])

  const enabledColumns = analysis?.columns.filter((column) => column.enabled) ?? []

  const handleFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) return

    setIsLoading(true)
    setErrorMessage('')

    try {
      if (!file.name.toLowerCase().endsWith('.xlsx')) {
        throw new Error('Пока поддерживается только формат .xlsx')
      }

      const loadedWorkbook = await loadWorkbookFromFile(file)
      const firstSheet = loadedWorkbook.worksheets[0]
      if (!firstSheet) throw new Error('В файле не найдено листов')

      setWorkbook(loadedWorkbook)
      setFileName(file.name)
      setAnalysis(analyzeWorksheet(loadedWorkbook, firstSheet.name))
    } catch (error) {
      setWorkbook(null)
      setAnalysis(null)
      setFileName(null)
      setErrorMessage(error instanceof Error ? error.message : 'Не удалось прочитать файл')
    } finally {
      setIsLoading(false)
    }
  }

  const handleSheetChange = (sheetName: string) => {
    if (!workbook) return
    setErrorMessage('')
    setAnalysis(analyzeWorksheet(workbook, sheetName))
  }

  const handleHeaderRowChange = (headerRowNumber: number) => {
    if (!workbook || !analysis) return
    if (!Number.isFinite(headerRowNumber) || headerRowNumber < 1) return
    setErrorMessage('')
    setAnalysis(analyzeWorksheet(workbook, analysis.selectedSheetName, headerRowNumber))
  }

  const handleColumnChange = (id: string, updater: (column: ColumnConfig) => ColumnConfig) => {
    setAnalysis((current) => {
      if (!current) return current
      return {
        ...current,
        columns: updateColumnById(current.columns, id, updater),
      }
    })
  }

  const handleGenerate = async () => {
    if (!workbook || !analysis) return

    setIsGenerating(true)
    setErrorMessage('')

    try {
      const outputWorkbook = await generateWorkbook(workbook, {
        sheetName: analysis.selectedSheetName,
        headerRowNumber: analysis.headerRowNumber,
        ignoredRowNumbers: parseIgnoredRowNumbers(ignoredRowsInput),
        rowsCount,
        columns: analysis.columns,
      })

      await downloadWorkbook(outputWorkbook, getGeneratedFileName(fileName, rowsCount))
    } catch (error) {
      setErrorMessage(error instanceof Error ? error.message : 'Не удалось сгенерировать файл')
    } finally {
      setIsGenerating(false)
    }
  }

  return (
    <main className="app-shell">
      <section className="hero-card">
        <div>
          <h1>XLSX Template Generator</h1>
          <p className="lead">
            Загрузи Excel-шаблон с заголовками и несколькими заполненными строками. Приложение распознает колонки,
            предложит типы данных и соберёт новый файл по образцу.
          </p>
        </div>
        <label className="upload-box">
          <input type="file" accept=".xlsx" onChange={handleFileChange} />
          <span className="upload-title">{isLoading ? 'Читаю файл…' : 'Загрузить .xlsx'}</span>
        </label>
      </section>

      {errorMessage && <div className="alert alert-error">{errorMessage}</div>}

      {!analysis && !isLoading && (
        <section className="empty-state">
          <h2>Как подготовить шаблон</h2>
          <div className="steps-grid">
            <div>
              <span className="step-number">1</span>
              <h3>Укажи номер строки с заголовками</h3>
              <p>Генератор возьмёт названия колонок из этой строки, а значения будет считать со следующей строки.</p>
            </div>
            <div>
              <span className="step-number">2</span>
              <h3>Добавь 1–10 строк-примеров</h3>
              <p>
                По ним генератор угадает типы данных и возможные значения списков.
              </p>
            </div>
            <div>
              <span className="step-number">3</span>
              <h3>Проверь настройки</h3>
              <p>Автоопределение можно поправить вручную перед генерацией.</p>
            </div>
          </div>
        </section>
      )}

      {analysis && (
        <>
          <section className="settings-card">
            <div className="settings-grid">
              <label>
                <span>Файл</span>
                <input value={fileName ?? ''} readOnly />
              </label>
              <label>
                <span>Лист</span>
                <select value={analysis.selectedSheetName} onChange={(event) => handleSheetChange(event.target.value)}>
                  {analysis.sheetNames.map((sheetName) => (
                    <option key={sheetName} value={sheetName}>
                      {sheetName}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                <span className="label-with-tooltip">
                  Номер строки с заголовками
                  <span
                    className="tooltip-icon"
                    tabIndex={0}
                    aria-label="Названия колонок берутся из этой строки, значения — со строк ниже."
                    data-tooltip="Названия колонок берутся из этой строки, значения — со строк ниже."
                  >
                    ?
                  </span>
                </span>
                <input
                  type="number"
                  min={1}
                  value={analysis.headerRowNumber}
                  onChange={(event) => handleHeaderRowChange(Number(event.target.value))}
                />
              </label>
              <label>
                <span>Игнорировать строки</span>
                <input
                  value={ignoredRowsInput}
                  onChange={(event) => setIgnoredRowsInput(event.target.value)}
                  placeholder="например: 1, 3, 10-12"
                />
              </label>
              <label>
                <span>Итоговое количество строк данных</span>
                <input
                  type="number"
                  min={1}
                  max={50000}
                  value={rowsCount}
                  onChange={(event) => setRowsCount(Math.max(1, Number(event.target.value) || 1))}
                />
              </label>
            </div>
            <p className="muted">
              Найдено колонок: {analysis.columns.length}. Заголовки взяты из строки {analysis.headerRowNumber}, значения считаются со строки{' '}
              {analysis.headerRowNumber + 1}. Существующие строки данных заменяются новыми сгенерированными строками.
            </p>
          </section>

          <section className="columns-card">
            <div className="section-header">
              <div>
                <p className="eyebrow">Шаг 2</p>
                <h2>Проверь распознанные колонки</h2>
              </div>
              <button className="primary-button" type="button" onClick={handleGenerate} disabled={isGenerating || !enabledColumns.length}>
                {isGenerating ? 'Генерирую…' : 'Скачать XLSX'}
              </button>
            </div>

            <div className="columns-table-wrap">
              <table className="columns-table">
                <thead>
                  <tr>
                    <th>Вкл.</th>
                    <th>Колонка</th>
                    <th>Тип</th>
                    <th>Параметры</th>
                    <th>Примеры</th>
                  </tr>
                </thead>
                <tbody>
                  {analysis.columns.map((column) => (
                    <tr key={column.id} className={!column.enabled ? 'disabled-row' : ''}>
                      <td>
                        <input
                          type="checkbox"
                          checked={column.enabled}
                          onChange={(event) => handleColumnChange(column.id, (current) => ({ ...current, enabled: event.target.checked }))}
                        />
                      </td>
                      <td>
                        <strong>{column.header}</strong>
                        <span className="column-index">#{column.index}</span>
                      </td>
                      <td>
                        <select
                          value={column.type}
                          onChange={(event) =>
                            handleColumnChange(column.id, (current) => ({ ...current, type: event.target.value as ColumnType }))
                          }
                        >
                          {columnTypes.map((type) => (
                            <option key={type.value} value={type.value}>
                              {type.label}
                            </option>
                          ))}
                        </select>
                      </td>
                      <td>
                        <ColumnParameters column={column} onChange={(updater) => handleColumnChange(column.id, updater)} />
                      </td>
                      <td>
                        <div className="samples-list">
                          {column.samples.length ? column.samples.map((sample, index) => <span key={`${sample}-${index}`}>{sample}</span>) : <em>нет примеров</em>}
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>

          <section className="preview-card">
            <div className="section-header">
              <div>
                <p className="eyebrow">Шаг 3</p>
                <h2>Предпросмотр</h2>
              </div>
              <span className="muted">Первые {previewRows.length} строк</span>
            </div>

            {previewRows.length ? (
              <div className="preview-table-wrap">
                <table className="preview-table">
                  <thead>
                    <tr>
                      {enabledColumns.map((column) => (
                        <th key={column.id}>{column.header}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {previewRows.map((row, rowIndex) => (
                      <tr key={rowIndex}>
                        {enabledColumns.map((column) => (
                          <td key={column.id}>{row[column.header]}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <p className="muted">Включи хотя бы одну колонку для предпросмотра.</p>
            )}
          </section>
        </>
      )}
    </main>
  )
}

function ColumnParameters({
  column,
  onChange,
}: {
  column: ColumnConfig
  onChange: (updater: (column: ColumnConfig) => ColumnConfig) => void
}) {
  if (column.type === 'number') {
    return (
      <div className="parameter-grid two-cols">
        <label>
          min
          <input
            type="number"
            value={column.min ?? ''}
            onChange={(event) => onChange((current) => ({ ...current, min: event.target.value === '' ? undefined : Number(event.target.value) }))}
          />
        </label>
        <label>
          max
          <input
            type="number"
            value={column.max ?? ''}
            onChange={(event) => onChange((current) => ({ ...current, max: event.target.value === '' ? undefined : Number(event.target.value) }))}
          />
        </label>
        <label className="checkbox-label full-width">
          <input
            type="checkbox"
            checked={column.unique}
            onChange={(event) => onChange((current) => ({ ...current, unique: event.target.checked }))}
          />
          уникальные значения
        </label>
      </div>
    )
  }

  if (column.type === 'date') {
    return (
      <div className="parameter-grid two-cols">
        <label>
          от
          <input
            type="date"
            value={column.dateFrom ?? ''}
            onChange={(event) => onChange((current) => ({ ...current, dateFrom: event.target.value || undefined }))}
          />
        </label>
        <label>
          до
          <input
            type="date"
            value={column.dateTo ?? ''}
            onChange={(event) => onChange((current) => ({ ...current, dateTo: event.target.value || undefined }))}
          />
        </label>
      </div>
    )
  }

  if (column.type === 'enum') {
    return (
      <label className="full-width">
        значения через запятую
        <input
          value={column.enumValues.join(', ')}
          onChange={(event) =>
            onChange((current) => ({
              ...current,
              enumValues: event.target.value.split(',').map((value) => value.trim()).filter(Boolean),
            }))
          }
          placeholder="active, blocked, pending"
        />
      </label>
    )
  }

  if (column.type === 'formula') {
    return (
      <label className="full-width">
        формула без =
        <input
          value={column.formula ?? ''}
          onChange={(event) => onChange((current) => ({ ...current, formula: event.target.value }))}
          placeholder="A2*B2"
        />
      </label>
    )
  }

  if (['uuid', 'email', 'phone', 'fullName', 'firstName', 'lastName', 'middleName', 'company', 'text'].includes(column.type)) {
    return (
      <label className="checkbox-label">
        <input
          type="checkbox"
          checked={column.unique}
          onChange={(event) => onChange((current) => ({ ...current, unique: event.target.checked }))}
        />
        уникальные значения
      </label>
    )
  }

  return <span className="muted">без параметров</span>
}

export default App
