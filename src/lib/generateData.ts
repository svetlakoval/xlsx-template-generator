import { fakerRU as faker } from '@faker-js/faker'
import type { ColumnConfig, PreviewRow } from './types'

const countries = [
  'Россия',
  'Казахстан',
  'Беларусь',
  'Армения',
  'Киргизия',
  'Узбекистан',
  'Таджикистан',
  'Азербайджан',
  'Грузия',
  'Турция',
  'Китай',
  'Индия',
  'Германия',
  'Франция',
  'Италия',
  'Испания',
  'Австрия',
  'Польша',
  'Сербия',
  'США',
  'Канада',
  'Бразилия',
]

const issuerPrefixes = ['ОВД', 'УФМС', 'ГУ МВД', 'УМВД', 'МВД']
const issuerDistricts = [
  'Тверского района г. Москвы',
  'Пресненского района г. Москвы',
  'Басманного района г. Москвы',
  'Хамовники г. Москвы',
  'района Арбат г. Москвы',
  'района Сокольники г. Москвы',
  'района Измайлово г. Москвы',
  'Московской области',
]

type Sex = 'male' | 'female'

type PersonNameDefinitions = {
  first_name?: Partial<Record<Sex | 'generic', string[]>>
  last_name?: Partial<Record<Sex | 'generic', string[]>>
  middle_name?: Partial<Record<Sex, string[]>>
}

const fakerDefinitions = faker as unknown as {
  definitions?: { person?: PersonNameDefinitions }
  rawDefinitions?: { person?: PersonNameDefinitions }
}

const getPersonNames = <T extends keyof PersonNameDefinitions>(
  key: T,
  sex: Sex,
  fallback: string[],
) => fakerDefinitions.definitions?.person?.[key]?.[sex] ?? fakerDefinitions.rawDefinitions?.person?.[key]?.[sex] ?? fallback

type PersonData = {
  sex: Sex
  gender: 'М' | 'Ж'
  firstName: string
  lastName: string
  middleName: string
  fullName: string
}

export type RowGenerationContext = {
  person?: PersonData
}

const firstNames: Record<Sex, string[]> = {
  male: [
    'Александр',
    'Алексей',
    'Андрей',
    'Виктор',
    'Дмитрий',
    'Иван',
    'Максим',
    'Михаил',
    'Никита',
    'Павел',
    'Панфил',
    'Роман',
    'Сергей',
    'Федор',
    'Юрий',
  ],
  female: [
    'Анна',
    'Виктория',
    'Дарья',
    'Екатерина',
    'Елена',
    'Ирина',
    'Мария',
    'Наталья',
    'Ольга',
    'Полина',
    'Светлана',
    'Татьяна',
    'Юлия',
  ],
}

const lastNames: Array<Record<Sex, string>> = [
  { male: 'Иванов', female: 'Иванова' },
  { male: 'Петров', female: 'Петрова' },
  { male: 'Сидоров', female: 'Сидорова' },
  { male: 'Смирнов', female: 'Смирнова' },
  { male: 'Кузнецов', female: 'Кузнецова' },
  { male: 'Попов', female: 'Попова' },
  { male: 'Васильев', female: 'Васильева' },
  { male: 'Соколов', female: 'Соколова' },
  { male: 'Михайлов', female: 'Михайлова' },
  { male: 'Новиков', female: 'Новикова' },
  { male: 'Федоров', female: 'Федорова' },
  { male: 'Морозов', female: 'Морозова' },
  { male: 'Волков', female: 'Волкова' },
  { male: 'Алексеев', female: 'Алексеева' },
  { male: 'Лебедев', female: 'Лебедева' },
  { male: 'Молчанов', female: 'Молчанова' },
]

const middleNames: Record<Sex, string[]> = {
  male: [
    'Александрович',
    'Алексеевич',
    'Андреевич',
    'Викторович',
    'Дмитриевич',
    'Иванович',
    'Михайлович',
    'Николаевич',
    'Павлович',
    'Сергеевич',
    'Федорович',
    'Юрьевич',
  ],
  female: [
    'Александровна',
    'Алексеевна',
    'Андреевна',
    'Викторовна',
    'Дмитриевна',
    'Ивановна',
    'Михайловна',
    'Николаевна',
    'Павловна',
    'Сергеевна',
    'Федоровна',
    'Юрьевна',
  ],
}

const normalizeHeader = (header: string) => header.trim().toLowerCase()

const isFirstNameHeader = (header: string) => /(^|[_\s-])(first.?name|given.?name|имя)([_\s-]|$)/i.test(header)
const isLastNameHeader = (header: string) => /(last.?name|surname|family.?name|фамил)/i.test(header)
const isMiddleNameHeader = (header: string) => /(middle.?name|patronymic|отчеств)/i.test(header)
const isGenderHeader = (header: string) => /(^|[_\s-])(gender|sex|пол)([_\s-]|$)/i.test(header)
const isCountryHeader = (header: string) => /(citizenship|nationality|country|гражданств|страна|подданств)/i.test(header)
const isAddressHeader = (header: string) => /(address|адрес|регистрац|пропис)/i.test(header)
const isIssuerHeader = (header: string) => /(issued.?by|issuer|кем.?выдан|выдан|орган.?выдач|подразделен)/i.test(header)
const isBlankNumberHeader = (header: string) => /(blank.?number|номер.?бланк|бланк)/i.test(header)

const pickByRow = <T,>(items: T[], rowIndex: number, salt = 0) => items[(rowIndex * 7 + salt) % items.length]

type UniqueStore = Map<string, Set<string>>

const padDatePart = (value: number) => String(value).padStart(2, '0')
const formatDate = (date: Date) => `${padDatePart(date.getDate())}.${padDatePart(date.getMonth() + 1)}.${date.getFullYear()}`

const createPerson = (rowIndex: number, sexOverride?: Sex): PersonData => {
  const sex: Sex = sexOverride ?? (rowIndex % 2 === 0 ? 'male' : 'female')
  const availableFirstNames = getPersonNames('first_name', sex, firstNames[sex])
  const availableLastNames = getPersonNames('last_name', sex, lastNames.map((name) => name[sex]))
  const availableMiddleNames = getPersonNames('middle_name', sex, middleNames[sex])
  const firstName = pickByRow(availableFirstNames, rowIndex, sex === 'male' ? 2 : 5)
  const lastName = pickByRow(availableLastNames, rowIndex, 3)
  const middleName = pickByRow(availableMiddleNames, rowIndex, 7)

  return {
    sex,
    gender: sex === 'male' ? 'М' : 'Ж',
    firstName,
    lastName,
    middleName,
    fullName: `${lastName} ${firstName} ${middleName}`,
  }
}

const getRowPerson = (rowIndex: number, rowContext: RowGenerationContext) => {
  if (rowContext.person) return rowContext.person

  rowContext.person = createPerson(rowIndex)
  return rowContext.person
}

const createTextFromSamples = (column: ColumnConfig, rowIndex: number) => {
  const samples = column.samples.map((sample) => sample.trim()).filter(Boolean)
  if (!samples.length) return null
  if (samples.some((sample) => /\d+/.test(sample))) return createSequentialValueFromSamples(column, rowIndex, samples[0])
  return samples[rowIndex % samples.length]
}

const createTextValue = (column: ColumnConfig, rowIndex: number, rowContext: RowGenerationContext) => {
  const header = normalizeHeader(column.header)

  if (isLastNameHeader(header)) return getRowPerson(rowIndex, rowContext).lastName
  if (isFirstNameHeader(header)) return getRowPerson(rowIndex, rowContext).firstName
  if (isMiddleNameHeader(header)) return getRowPerson(rowIndex, rowContext).middleName
  if (isGenderHeader(header)) return getRowPerson(rowIndex, rowContext).gender
  if (isCountryHeader(header)) return countries[rowIndex % countries.length]
  if (isAddressHeader(header)) return createAddressFromSamples(column, rowIndex)
  if (isIssuerHeader(header)) return createIssuer(rowIndex)
  if (isBlankNumberHeader(header)) return createBlankNumberFromSamples(column, rowIndex)

  return createTextFromSamples(column, rowIndex) ?? faker.lorem.words({ min: 2, max: 5 })
}

const parseDateInput = (value?: string, fallback?: Date) => {
  if (!value) return fallback ?? new Date()

  const ruDateMatch = value.trim().match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/)
  if (ruDateMatch) {
    const [, day, month, year] = ruDateMatch
    const parsed = new Date(Number(year), Number(month) - 1, Number(day))
    return Number.isNaN(parsed.getTime()) ? fallback ?? new Date() : parsed
  }

  const parsed = new Date(value)
  return Number.isNaN(parsed.getTime()) ? fallback ?? new Date() : parsed
}

const randomInteger = (min = 1, max = 1000) => {
  const safeMin = Number.isFinite(min) ? Math.floor(min) : 1
  const safeMax = Number.isFinite(max) ? Math.floor(max) : safeMin + 1000
  return faker.number.int({ min: Math.min(safeMin, safeMax), max: Math.max(safeMin, safeMax) })
}

const createPhoneNumber = () => `79${faker.string.numeric(9)}`

const createSequentialValueFromSamples = (column: ColumnConfig, rowIndex: number, fallbackPrefix: string) => {
  const samples = column.samples.map((sample) => sample.trim()).filter(Boolean)
  const numberedSample = samples.find((sample) => /\d+/.test(sample))

  if (!numberedSample) return `${fallbackPrefix}${String(rowIndex + 1).padStart(6, '0')}`

  const match = numberedSample.match(/^(.*?)(\d+)(\D*)$/)
  if (!match) return `${fallbackPrefix}${String(rowIndex + 1).padStart(6, '0')}`

  const [, prefix, numberPart, suffix] = match
  const sampleNumbers = samples
    .map((sample) => sample.match(/(\d+)(?!.*\d)/)?.[1])
    .filter((value): value is string => Boolean(value))
    .map(Number)
    .filter(Number.isFinite)
  const startNumber = sampleNumbers.length ? Math.min(...sampleNumbers) : Number(numberPart)
  const nextNumber = startNumber + rowIndex

  return `${prefix}${String(nextNumber).padStart(numberPart.length, '0')}${suffix}`
}

const createAddressFromSamples = (column: ColumnConfig, rowIndex: number) => createSequentialValueFromSamples(column, rowIndex, 'г. Москва, ул. Тарасовская, д.')

const createBlankNumberFromSamples = (column: ColumnConfig, rowIndex: number) => createSequentialValueFromSamples(column, rowIndex, 'БЛ')

const createIssuer = (rowIndex: number) => {
  const prefix = pickByRow(issuerPrefixes, rowIndex, 1)
  const district = pickByRow(issuerDistricts, rowIndex, 3)
  return `${prefix} ${district}`
}

const valueToKey = (value: unknown) => {
  if (value instanceof Date) return value.toISOString()
  return String(value ?? '')
}

const withUnique = (
  column: ColumnConfig,
  uniqueStore: UniqueStore,
  createValue: (attempt: number) => unknown,
) => {
  if (!column.unique) return createValue(0)

  const key = column.id
  const seen = uniqueStore.get(key) ?? new Set<string>()
  uniqueStore.set(key, seen)

  for (let attempt = 0; attempt < 100; attempt += 1) {
    const value = createValue(attempt)
    const valueKey = valueToKey(value)
    if (!seen.has(valueKey)) {
      seen.add(valueKey)
      return value
    }
  }

  const fallbackValue = createValue(seen.size + 100)
  seen.add(valueToKey(fallbackValue))
  return fallbackValue
}

const makeUniquePersonValue = (
  column: ColumnConfig,
  uniqueStore: UniqueStore,
  rowIndex: number,
  rowContext: RowGenerationContext,
  createValue: (person: PersonData) => string,
) => {
  const value = createValue(getRowPerson(rowIndex, rowContext))
  if (!column.unique) return value

  const seen = uniqueStore.get(column.id) ?? new Set<string>()
  uniqueStore.set(column.id, seen)
  seen.add(valueToKey(value))
  return value
}

export const generateValue = (
  column: ColumnConfig,
  rowIndex: number,
  uniqueStore: UniqueStore,
  rowContext: RowGenerationContext = {},
): unknown => {
  if (!column.enabled) return null

  switch (column.type) {
    case 'uuid':
      return withUnique(column, uniqueStore, () => faker.string.uuid())
    case 'email':
      return withUnique(column, uniqueStore, (attempt) => {
        const firstName = faker.person.firstName()
        const lastName = faker.person.lastName()
        const prefix = attempt ? `${firstName}.${lastName}.${attempt}` : `${firstName}.${lastName}`
        return faker.internet.email({ firstName: prefix, lastName })
      })
    case 'phone':
      return withUnique(column, uniqueStore, () => createPhoneNumber())
    case 'fullName':
      return column.unique
        ? makeUniquePersonValue(column, uniqueStore, rowIndex, rowContext, (person) => person.fullName)
        : getRowPerson(rowIndex, rowContext).fullName
    case 'firstName':
      return column.unique
        ? makeUniquePersonValue(column, uniqueStore, rowIndex, rowContext, (person) => person.firstName)
        : getRowPerson(rowIndex, rowContext).firstName
    case 'lastName':
      return column.unique
        ? makeUniquePersonValue(column, uniqueStore, rowIndex, rowContext, (person) => person.lastName)
        : getRowPerson(rowIndex, rowContext).lastName
    case 'middleName':
      return column.unique
        ? makeUniquePersonValue(column, uniqueStore, rowIndex, rowContext, (person) => person.middleName)
        : getRowPerson(rowIndex, rowContext).middleName
    case 'gender':
      return getRowPerson(rowIndex, rowContext).gender
    case 'company':
      return withUnique(column, uniqueStore, () => faker.company.name())
    case 'city':
      return faker.location.city()
    case 'country':
      return countries[rowIndex % countries.length]
    case 'address':
      return createAddressFromSamples(column, rowIndex)
    case 'blankNumber':
      return createBlankNumberFromSamples(column, rowIndex)
    case 'issuer':
      return createIssuer(rowIndex)
    case 'boolean':
      return faker.datatype.boolean()
    case 'number': {
      if (column.unique) {
        const start = Number.isFinite(column.min) ? Number(column.min) : 1
        return start + rowIndex
      }
      return randomInteger(column.min, column.max)
    }
    case 'date': {
      const now = new Date()
      const yearAgo = new Date(now)
      yearAgo.setFullYear(now.getFullYear() - 1)

      const from = parseDateInput(column.dateFrom, yearAgo)
      const to = parseDateInput(column.dateTo, now)
      const min = from.getTime() <= to.getTime() ? from : to
      const max = from.getTime() <= to.getTime() ? to : from
      return formatDate(faker.date.between({ from: min, to: max }))
    }
    case 'enum': {
      const values = column.enumValues.map((value) => value.trim()).filter(Boolean)
      if (!values.length) return faker.lorem.word()
      return faker.helpers.arrayElement(values)
    }
    case 'formula':
      return column.formula ? `=${column.formula}` : ''
    case 'text':
    default:
      return withUnique(column, uniqueStore, () => createTextValue(column, rowIndex, rowContext))
  }
}

export const formatPreviewValue = (value: unknown) => {
  if (value === null || value === undefined) return ''
  if (value instanceof Date) return formatDate(value)
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  return String(value)
}

export const generatePreviewRows = (columns: ColumnConfig[], count = 5): PreviewRow[] => {
  const enabledColumns = columns.filter((column) => column.enabled)
  const uniqueStore: UniqueStore = new Map()

  return Array.from({ length: count }, (_, rowIndex) => {
    const rowContext: RowGenerationContext = {}

    return enabledColumns.reduce<PreviewRow>((row, column) => {
      row[column.header] = formatPreviewValue(generateValue(column, rowIndex, uniqueStore, rowContext))
      return row
    }, {})
  })
}
