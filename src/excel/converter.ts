import type { ICellData, IColorStyle, IStyleData, IWorkbookData, IWorksheetData, LocaleType } from '@univerjs/core'
import type {
  Border,
  Borders,
  Cell,
  DefinedNamesModel,
  Table as ExcelTable,
  Workbook as ExcelWorkbook,
  Fill,
  Font,
  HeaderFooter,
  PageSetup,
  TableColumnProperties,
  TableStyleProperties,
  Worksheet,
} from 'exceljs'
import {
  BooleanNumber,
  BorderStyleTypes,
  CellValueType,
  CustomRangeType,
  HorizontalAlign,
  TextDecoration,
  VerticalAlign,
  WrapStrategy,
} from '@univerjs/core'
import ExcelJS from 'exceljs'
import { validateXlsxBuffer } from './capabilities'
import { normalizeXlsxForExcelJs } from './compatibility'
import { workbookMetrics } from './safety'

const DEFAULT_ROW_COUNT = 100
const DEFAULT_COLUMN_COUNT = 26
const DEFAULT_ROW_HEIGHT = 24
const DEFAULT_COLUMN_WIDTH = 88
const UNSUPPORTED_RESOURCE_NAMES = new Set([
  'SHEET_CONDITIONAL_FORMATTING_PLUGIN',
  'SHEET_DATA_VALIDATION_PLUGIN',
  'SHEET_FILTER_PLUGIN',
])

interface ExcelCellMetadata {
  date?: boolean
  error?: string
  hyperlink?: string
}

interface ExcelWorksheetMetadata {
  originalName: string
  originalRowCount: number
  originalColumnCount: number
  pageSetup?: Partial<PageSetup>
  headerFooter?: Partial<HeaderFooter>
  tables?: ExcelTableMetadata[]
}

interface ExcelTableMetadata {
  name: string
  displayName: string
  ref: string
  headerRow: boolean
  totalsRow: boolean
  columns: TableColumnProperties[]
  style: TableStyleProperties
}

interface ExcelTableModel {
  name: string
  displayName?: string
  tableRef: string
  headerRow?: boolean
  totalsRow?: boolean
  columns: TableColumnProperties[]
  style?: TableStyleProperties
}

interface ExcelWorkbookMetadata {
  definedNames?: DefinedNamesModel
}

function createId(prefix: string): string {
  const random = globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(16).slice(2)}`
  return `${prefix}-${random}`
}

function argbToRgb(argb?: string): string | undefined {
  if (!argb)
    return undefined
  const value = argb.length === 8 ? argb.slice(2) : argb
  return `#${value.toUpperCase()}`
}

function colorToUniver(color?: { argb?: string }): IColorStyle | undefined {
  const rgb = argbToRgb(color?.argb)
  return rgb ? { rgb } : undefined
}

function excelBorderStyle(style?: Border['style']): BorderStyleTypes {
  const styles: Partial<Record<NonNullable<Border['style']>, BorderStyleTypes>> = {
    thin: BorderStyleTypes.THIN,
    hair: BorderStyleTypes.HAIR,
    dotted: BorderStyleTypes.DOTTED,
    dashed: BorderStyleTypes.DASHED,
    dashDot: BorderStyleTypes.DASH_DOT,
    dashDotDot: BorderStyleTypes.DASH_DOT_DOT,
    double: BorderStyleTypes.DOUBLE,
    medium: BorderStyleTypes.MEDIUM,
    mediumDashed: BorderStyleTypes.MEDIUM_DASHED,
    mediumDashDot: BorderStyleTypes.MEDIUM_DASH_DOT,
    mediumDashDotDot: BorderStyleTypes.MEDIUM_DASH_DOT_DOT,
    slantDashDot: BorderStyleTypes.SLANT_DASH_DOT,
    thick: BorderStyleTypes.THICK,
  }
  return style ? styles[style] ?? BorderStyleTypes.THIN : BorderStyleTypes.NONE
}

function borderToUniver(border?: Partial<Border>) {
  if (!border?.style)
    return undefined
  return {
    s: excelBorderStyle(border.style),
    cl: colorToUniver(border.color) ?? { rgb: '#000000' },
  }
}

function excelStyleToUniver(cell: Cell): IStyleData | undefined {
  const font = cell.font
  const fill = cell.fill
  const alignment = cell.alignment
  const border = cell.border
  const style: IStyleData = {}

  if (font?.name)
    style.ff = font.name
  if (font?.size)
    style.fs = font.size
  if (font?.bold)
    style.bl = BooleanNumber.TRUE
  if (font?.italic)
    style.it = BooleanNumber.TRUE
  if (font?.underline) {
    style.ul = {
      s: BooleanNumber.TRUE,
      t: font.underline === 'double' ? TextDecoration.DOUBLE : TextDecoration.SINGLE,
    }
  }
  if (font?.strike)
    style.st = { s: BooleanNumber.TRUE }
  style.cl = colorToUniver(font?.color)

  if (fill?.type === 'pattern' && fill.pattern === 'solid')
    style.bg = colorToUniver(fill.fgColor)

  if (cell.numFmt && cell.numFmt !== 'General')
    style.n = { pattern: cell.numFmt }

  const horizontal = { left: HorizontalAlign.LEFT, center: HorizontalAlign.CENTER, right: HorizontalAlign.RIGHT, justify: HorizontalAlign.JUSTIFIED, distributed: HorizontalAlign.DISTRIBUTED }
  const vertical = { top: VerticalAlign.TOP, middle: VerticalAlign.MIDDLE, bottom: VerticalAlign.BOTTOM, justify: VerticalAlign.MIDDLE, distributed: VerticalAlign.MIDDLE }
  if (alignment?.horizontal && alignment.horizontal in horizontal)
    style.ht = horizontal[alignment.horizontal as keyof typeof horizontal]
  if (alignment?.vertical && alignment.vertical in vertical)
    style.vt = vertical[alignment.vertical as keyof typeof vertical]
  if (alignment?.wrapText)
    style.tb = WrapStrategy.WRAP
  if (typeof alignment?.textRotation === 'number')
    style.tr = { a: alignment.textRotation }

  const borders = {
    t: borderToUniver(border?.top),
    r: borderToUniver(border?.right),
    b: borderToUniver(border?.bottom),
    l: borderToUniver(border?.left),
  }
  if (Object.values(borders).some(Boolean))
    style.bd = borders

  for (const key of Object.keys(style) as Array<keyof IStyleData>) {
    if (style[key] === undefined)
      delete style[key]
  }
  return Object.keys(style).length > 0 ? style : undefined
}

function primitiveValue(value: unknown): string | number | boolean | undefined {
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean')
    return value
  if (value instanceof Date)
    return dateToExcelSerial(value)
  return undefined
}

function dateToExcelSerial(date: Date): number {
  return (date.getTime() - Date.UTC(1899, 11, 30)) / 86400000
}

function excelSerialToDate(value: number): Date {
  return new Date(Date.UTC(1899, 11, 30) + value * 86400000)
}

function excelCellToUniver(cell: Cell, styleId?: string): ICellData | undefined {
  const value = cell.value
  const data: ICellData = {}
  const metadata: ExcelCellMetadata = {}

  if (styleId)
    data.s = styleId

  if (cell.formula) {
    const result = cell.result
    data.f = `=${cell.formula}`
    data.v = primitiveValue(result)
    if (result instanceof Date)
      metadata.date = true
  }
  else if (value instanceof Date) {
    data.v = dateToExcelSerial(value)
    data.t = CellValueType.NUMBER
    metadata.date = true
  }
  else if (typeof value === 'string') {
    data.v = value
    data.t = CellValueType.STRING
  }
  else if (typeof value === 'number') {
    data.v = value
    data.t = CellValueType.NUMBER
  }
  else if (typeof value === 'boolean') {
    data.v = value
    data.t = CellValueType.BOOLEAN
  }
  else if (value && typeof value === 'object' && 'richText' in value) {
    data.v = value.richText.map(part => part.text).join('')
    data.t = CellValueType.STRING
  }
  else if (value && typeof value === 'object' && 'hyperlink' in value) {
    data.v = value.text
    data.t = CellValueType.STRING
    metadata.hyperlink = value.hyperlink
  }
  else if (value && typeof value === 'object' && 'error' in value) {
    data.v = value.error
    data.t = CellValueType.STRING
    metadata.error = value.error
  }

  if (Object.keys(metadata).length > 0)
    data.custom = { excel: metadata }

  return Object.keys(data).length > 0 ? data : undefined
}

function columnToIndex(column: string): number {
  return [...column.toUpperCase()].reduce((value, character) => value * 26 + character.charCodeAt(0) - 64, 0) - 1
}

function parseCellAddress(address: string): { row: number, column: number } {
  const match = /^\$?([A-Z]+)\$?(\d+)$/i.exec(address)
  if (!match)
    throw new Error(`Invalid Excel cell address: ${address}`)
  return { row: Number(match[2]) - 1, column: columnToIndex(match[1]) }
}

function parseRange(range: string) {
  const [start, end = start] = range.split(':')
  const startCell = parseCellAddress(start)
  const endCell = parseCellAddress(end)
  return {
    startRow: startCell.row,
    startColumn: startCell.column,
    endRow: endCell.row,
    endColumn: endCell.column,
  }
}

function worksheetTableMetadata(worksheet: Worksheet): ExcelTableMetadata[] {
  const tables = worksheet.getTables() as unknown as ExcelTable[]
  return tables.map((table) => {
    const model = (table as unknown as { model: ExcelTableModel }).model
    if (!model.tableRef)
      throw new Error(`Cannot preserve Excel table ${model.name}: its range is missing`)
    return {
      name: model.name,
      displayName: model.displayName ?? model.name,
      ref: model.tableRef,
      headerRow: model.headerRow !== false,
      totalsRow: model.totalsRow === true,
      columns: serializableClone(model.columns ?? []),
      style: serializableClone(model.style ?? {}),
    }
  })
}

function worksheetToUniver(worksheet: Worksheet, styles: Record<string, IStyleData>): Partial<IWorksheetData> {
  const cellData: Record<number, Record<number, ICellData>> = {}
  const rowData: Record<number, { h?: number, hd?: BooleanNumber }> = {}
  const columnData: Record<number, { w?: number, hd?: BooleanNumber }> = {}
  const styleIds = new Map<string, string>()

  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    if (row.height || row.hidden) {
      rowData[rowNumber - 1] = {
        h: row.height ? row.height * 96 / 72 : undefined,
        hd: row.hidden ? BooleanNumber.TRUE : undefined,
      }
    }

    row.eachCell({ includeEmpty: true }, (cell, columnNumber) => {
      const style = excelStyleToUniver(cell)
      let styleId: string | undefined
      if (style) {
        const key = JSON.stringify(style)
        styleId = styleIds.get(key)
        if (!styleId) {
          styleId = `style-${styleIds.size + 1}`
          styleIds.set(key, styleId)
          styles[styleId] = style
        }
      }

      const converted = excelCellToUniver(cell, styleId)
      if (converted) {
        cellData[rowNumber - 1] ??= {}
        cellData[rowNumber - 1][columnNumber - 1] = converted
      }
    })
  })

  const columns = worksheet.columns ?? []
  columns.forEach((column, index) => {
    if (column.width || column.hidden) {
      columnData[index] = {
        w: column.width ? Math.max(10, column.width * 7 + 5) : undefined,
        hd: column.hidden ? BooleanNumber.TRUE : undefined,
      }
    }
  })

  const views = worksheet.views ?? []
  const view = views.find(item => item.state === 'frozen')
  const xSplit = view && 'xSplit' in view ? Number(view.xSplit ?? 0) : 0
  const ySplit = view && 'ySplit' in view ? Number(view.ySplit ?? 0) : 0
  const tabColor = argbToRgb(worksheet.properties.tabColor?.argb) ?? ''
  const rowCount = Math.max(DEFAULT_ROW_COUNT, worksheet.actualRowCount)
  const columnCount = Math.max(DEFAULT_COLUMN_COUNT, worksheet.actualColumnCount)

  return {
    id: createId('sheet'),
    name: worksheet.name,
    tabColor,
    hidden: worksheet.state === 'visible' ? BooleanNumber.FALSE : BooleanNumber.TRUE,
    freeze: { xSplit, ySplit, startColumn: xSplit, startRow: ySplit },
    rowCount,
    columnCount,
    defaultColumnWidth: DEFAULT_COLUMN_WIDTH,
    defaultRowHeight: DEFAULT_ROW_HEIGHT,
    mergeData: (worksheet.model.merges ?? []).map(parseRange),
    cellData,
    rowData,
    columnData,
    rowHeader: { width: 46 },
    columnHeader: { height: 20 },
    showGridlines: views[0]?.showGridLines === false ? BooleanNumber.FALSE : BooleanNumber.TRUE,
    rightToLeft: views[0]?.rightToLeft ? BooleanNumber.TRUE : BooleanNumber.FALSE,
    custom: {
      excel: {
        originalName: worksheet.name,
        originalRowCount: rowCount,
        originalColumnCount: columnCount,
        pageSetup: serializableClone(worksheet.pageSetup),
        headerFooter: serializableClone(worksheet.headerFooter),
        tables: worksheetTableMetadata(worksheet),
      } satisfies ExcelWorksheetMetadata,
    },
  }
}

export async function importXlsx(buffer: ArrayBuffer, name: string, locale: LocaleType): Promise<IWorkbookData> {
  const workbook = new ExcelJS.Workbook()
  if (buffer.byteLength > 0) {
    const compatibleBuffer = await normalizeXlsxForExcelJs(buffer)
    await workbook.xlsx.load(new Uint8Array(compatibleBuffer) as unknown as ExcelJS.Buffer)
  }
  if (workbook.worksheets.length === 0)
    workbook.addWorksheet('Sheet1')

  const styles: Record<string, IStyleData> = {}
  const sheets: IWorkbookData['sheets'] = {}
  const sheetOrder: string[] = []

  for (const worksheet of workbook.worksheets) {
    const sheet = worksheetToUniver(worksheet, styles)
    const sheetId = sheet.id!
    sheetOrder.push(sheetId)
    sheets[sheetId] = sheet
  }

  return {
    id: createId('workbook'),
    name,
    appVersion: '0.25.1',
    locale,
    styles,
    sheetOrder,
    sheets,
    custom: {
      excel: {
        definedNames: serializableClone(workbook.definedNames.model ?? []),
      } satisfies ExcelWorkbookMetadata,
    },
  }
}

function rgbToArgb(color?: IColorStyle | null | void): string | undefined {
  const rgb = color?.rgb?.replace('#', '')
  return rgb ? `FF${rgb.toUpperCase()}` : undefined
}

function univerBorderStyle(style: BorderStyleTypes): Border['style'] {
  const styles: Partial<Record<BorderStyleTypes, Border['style']>> = {
    [BorderStyleTypes.THIN]: 'thin',
    [BorderStyleTypes.HAIR]: 'hair',
    [BorderStyleTypes.DOTTED]: 'dotted',
    [BorderStyleTypes.DASHED]: 'dashed',
    [BorderStyleTypes.DASH_DOT]: 'dashDot',
    [BorderStyleTypes.DASH_DOT_DOT]: 'dashDotDot',
    [BorderStyleTypes.DOUBLE]: 'double',
    [BorderStyleTypes.MEDIUM]: 'medium',
    [BorderStyleTypes.MEDIUM_DASHED]: 'mediumDashed',
    [BorderStyleTypes.MEDIUM_DASH_DOT]: 'mediumDashDot',
    [BorderStyleTypes.MEDIUM_DASH_DOT_DOT]: 'mediumDashDotDot',
    [BorderStyleTypes.SLANT_DASH_DOT]: 'slantDashDot',
    [BorderStyleTypes.THICK]: 'thick',
  }
  return styles[style] ?? 'thin'
}

function univerBorderToExcel(border?: { s: BorderStyleTypes, cl: IColorStyle } | null | void): Partial<Border> | undefined {
  if (!border || border.s === BorderStyleTypes.NONE)
    return undefined
  const argb = rgbToArgb(border.cl)
  return argb
    ? { style: univerBorderStyle(border.s), color: { argb } }
    : { style: univerBorderStyle(border.s) }
}

function applyUniverStyle(cell: Cell, style?: IStyleData | null): void {
  if (!style)
    return

  const font: Partial<Font> = {
    name: style.ff ?? undefined,
    size: style.fs,
    bold: style.bl === BooleanNumber.TRUE,
    italic: style.it === BooleanNumber.TRUE,
    underline: style.ul?.s === BooleanNumber.TRUE ? style.ul.t === TextDecoration.DOUBLE ? 'double' : true : undefined,
    strike: style.st?.s === BooleanNumber.TRUE,
  }
  const fontColor = rgbToArgb(style.cl)
  if (fontColor)
    font.color = { argb: fontColor }
  cell.font = font as Font

  const background = rgbToArgb(style.bg)
  if (background) {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: background },
    } as Fill
  }

  if (style.n?.pattern)
    cell.numFmt = style.n.pattern

  const horizontal = { [HorizontalAlign.LEFT]: 'left', [HorizontalAlign.CENTER]: 'center', [HorizontalAlign.RIGHT]: 'right', [HorizontalAlign.JUSTIFIED]: 'justify', [HorizontalAlign.DISTRIBUTED]: 'distributed' } as const
  const vertical = { [VerticalAlign.TOP]: 'top', [VerticalAlign.MIDDLE]: 'middle', [VerticalAlign.BOTTOM]: 'bottom' } as const
  cell.alignment = {
    horizontal: style.ht ? horizontal[style.ht as keyof typeof horizontal] : undefined,
    vertical: style.vt ? vertical[style.vt as keyof typeof vertical] : undefined,
    wrapText: style.tb === WrapStrategy.WRAP,
    textRotation: style.tr?.a,
  }

  if (style.bd) {
    cell.border = {
      top: univerBorderToExcel(style.bd.t),
      right: univerBorderToExcel(style.bd.r),
      bottom: univerBorderToExcel(style.bd.b),
      left: univerBorderToExcel(style.bd.l),
    } as Borders
  }
}

function getCellText(cell: ICellData): string | undefined {
  const dataStream = cell.p?.body?.dataStream
  return dataStream?.replace(/\r\n$/, '')
}

function applyUniverCell(cell: Cell, data: ICellData): void {
  const metadata = data.custom?.excel as ExcelCellMetadata | undefined
  const value = data.v ?? getCellText(data)
  const richTextHyperlink = data.p?.body?.customRanges?.find(range => range.rangeType === CustomRangeType.HYPERLINK)
  const richTextHyperlinkUrl = richTextHyperlink?.properties?.url

  if (data.f) {
    cell.value = {
      formula: data.f.startsWith('=') ? data.f.slice(1) : data.f,
      result: metadata?.date && typeof value === 'number' ? excelSerialToDate(value) : value ?? undefined,
    }
  }
  else if (metadata?.hyperlink && typeof value === 'string') {
    cell.value = { text: value, hyperlink: metadata.hyperlink }
  }
  else if (typeof richTextHyperlinkUrl === 'string' && typeof value === 'string') {
    cell.value = { text: value, hyperlink: richTextHyperlinkUrl }
  }
  else if (metadata?.error) {
    cell.value = { error: metadata.error as ExcelJS.CellErrorValue['error'] }
  }
  else if (metadata?.date && typeof value === 'number') {
    cell.value = excelSerialToDate(value)
  }
  else {
    cell.value = value ?? null
  }
}

function resolveStyle(workbook: IWorkbookData, cell: ICellData): IStyleData | undefined {
  if (!cell.s)
    return undefined
  if (typeof cell.s === 'string')
    return workbook.styles[cell.s] ?? undefined
  return cell.s
}

function applyWorksheetData(excelWorksheet: Worksheet, workbook: IWorkbookData, sheet: Partial<IWorksheetData>): void {
  const metadata = getExcelWorksheetMetadata(sheet)
  if (metadata?.pageSetup)
    Object.assign(excelWorksheet.pageSetup, metadata.pageSetup)
  if (metadata?.headerFooter)
    Object.assign(excelWorksheet.headerFooter, metadata.headerFooter)

  excelWorksheet.state = sheet.hidden === BooleanNumber.TRUE ? 'hidden' : 'visible'
  if (sheet.tabColor)
    excelWorksheet.properties.tabColor = { argb: `FF${sheet.tabColor.replace('#', '')}` }

  const frozen = Boolean(sheet.freeze && (sheet.freeze.xSplit > 0 || sheet.freeze.ySplit > 0))
  if (frozen && sheet.freeze) {
    excelWorksheet.views = [{
      state: 'frozen',
      xSplit: sheet.freeze.xSplit,
      ySplit: sheet.freeze.ySplit,
      topLeftCell: `${columnIndexToLetters(sheet.freeze.startColumn)}${sheet.freeze.startRow + 1}`,
      showGridLines: sheet.showGridlines !== BooleanNumber.FALSE,
      rightToLeft: sheet.rightToLeft === BooleanNumber.TRUE,
    }]
  }
  else {
    excelWorksheet.views = [{
      state: 'normal',
      showGridLines: sheet.showGridlines !== BooleanNumber.FALSE,
      rightToLeft: sheet.rightToLeft === BooleanNumber.TRUE,
    }]
  }

  for (const [rowIndex, row] of Object.entries(sheet.rowData ?? {})) {
    const excelRow = excelWorksheet.getRow(Number(rowIndex) + 1)
    if (row?.h)
      excelRow.height = row.h * 72 / 96
    excelRow.hidden = row?.hd === BooleanNumber.TRUE
  }

  for (const [columnIndex, column] of Object.entries(sheet.columnData ?? {})) {
    const excelColumn = excelWorksheet.getColumn(Number(columnIndex) + 1)
    if (column?.w)
      excelColumn.width = Math.max(1, (column.w - 5) / 7)
    excelColumn.hidden = column?.hd === BooleanNumber.TRUE
  }

  for (const [rowIndex, row] of Object.entries(sheet.cellData ?? {})) {
    for (const [columnIndex, data] of Object.entries(row ?? {})) {
      if (!data)
        continue
      const cell = excelWorksheet.getCell(Number(rowIndex) + 1, Number(columnIndex) + 1)
      applyUniverCell(cell, data)
      applyUniverStyle(cell, resolveStyle(workbook, data))
    }
  }

  for (const range of sheet.mergeData ?? [])
    excelWorksheet.mergeCells(range.startRow + 1, range.startColumn + 1, range.endRow + 1, range.endColumn + 1)

  applyExcelTables(excelWorksheet, sheet, metadata)
}

function applyExcelTables(excelWorksheet: Worksheet, sheet: Partial<IWorksheetData>, metadata?: ExcelWorksheetMetadata): void {
  if (!metadata?.tables?.length)
    return
  const tables = metadata.tables
  if (sheet.rowCount !== metadata.originalRowCount || sheet.columnCount !== metadata.originalColumnCount)
    throw new Error('Cannot save row or column structure changes while preserved Excel tables are present')

  for (const table of tables) {
    if (table.totalsRow)
      throw new Error(`Cannot safely rebuild Excel table ${table.name} because it has a totals row`)

    const range = parseRange(table.ref)
    const width = range.endColumn - range.startColumn + 1
    if (table.columns.length !== width)
      throw new Error(`Cannot safely rebuild Excel table ${table.name} because its column count changed`)

    const usedNames = new Set<string>()
    const columns = table.columns.map((column, index) => ({
      ...column,
      name: uniqueTableColumnName(
        table.headerRow ? excelWorksheet.getCell(range.startRow + 1, range.startColumn + index + 1).text : column.name,
        column.name,
        index,
        usedNames,
      ),
    }))
    const firstDataRow = range.startRow + (table.headerRow ? 1 : 0)
    const lastDataRow = range.endRow - (table.totalsRow ? 1 : 0)
    const rows: unknown[][] = []
    for (let row = firstDataRow; row <= lastDataRow; row++) {
      rows.push(columns.map((_, index) => excelWorksheet.getCell(row + 1, range.startColumn + index + 1).value))
    }

    excelWorksheet.addTable({
      name: table.name,
      displayName: table.displayName,
      ref: `${columnIndexToLetters(range.startColumn)}${range.startRow + 1}`,
      headerRow: table.headerRow,
      totalsRow: false,
      style: { ...table.style },
      columns,
      rows,
    })
  }
}

function uniqueTableColumnName(value: string, fallback: string, index: number, usedNames: Set<string>): string {
  const base = value.trim() || fallback.trim() || `Column${index + 1}`
  let name = base
  let suffix = 2
  while (usedNames.has(name.toLocaleLowerCase()))
    name = `${base}_${suffix++}`
  usedNames.add(name.toLocaleLowerCase())
  return name
}

function columnIndexToLetters(index: number): string {
  let value = index + 1
  let result = ''
  while (value > 0) {
    value--
    result = String.fromCharCode(65 + value % 26) + result
    value = Math.floor(value / 26)
  }
  return result
}

export async function exportXlsx(workbookData: IWorkbookData): Promise<ArrayBuffer> {
  assertNoUnsupportedResources(workbookData)
  const workbook: ExcelWorkbook = new ExcelJS.Workbook()
  for (const sheetId of workbookData.sheetOrder) {
    const sheet = workbookData.sheets[sheetId]
    if (!sheet)
      continue
    const worksheet = workbook.addWorksheet(sheet.name || 'Sheet')
    applyWorksheetData(worksheet, workbookData, sheet)
  }
  if (workbook.worksheets.length === 0)
    workbook.addWorksheet('Sheet1')
  applyWorkbookMetadata(workbook, workbookData)

  const output = await workbook.xlsx.writeBuffer()
  const bytes = new Uint8Array(output as ArrayBuffer)
  const buffer = bytes.slice().buffer
  await validateXlsxBuffer(buffer)
  await validateExportSemantics(workbookData, buffer)
  return buffer
}

async function validateExportSemantics(source: IWorkbookData, buffer: ArrayBuffer): Promise<void> {
  const verification = new ExcelJS.Workbook()
  await verification.xlsx.load(new Uint8Array(buffer) as unknown as ExcelJS.Buffer)
  const expected = workbookMetrics(source)
  if (verification.worksheets.length !== expected.sheetCount)
    throw new Error(`Exported ${verification.worksheets.length} worksheets, expected ${expected.sheetCount}`)

  const exportedCells = verification.worksheets.reduce((count, worksheet) => count + countPopulatedCells(worksheet), 0)
  if (expected.populatedCellCount > 0 && exportedCells === 0)
    throw new Error('The exported workbook unexpectedly contains no populated cells')

  const expectedNames = normalizedDefinedNames(definedNamesForExport(source))
  const actualNames = normalizedDefinedNames(verification.definedNames.model ?? [])
  if (JSON.stringify(actualNames) !== JSON.stringify(expectedNames))
    throw new Error('The exported workbook did not preserve its named ranges')

  for (const sheetId of source.sheetOrder) {
    const sourceSheet = source.sheets[sheetId]
    const metadata = getExcelWorksheetMetadata(sourceSheet)
    if (!sourceSheet || !metadata)
      continue
    const exportedSheet = verification.getWorksheet(sourceSheet.name ?? '')
    if (!exportedSheet)
      throw new Error(`The exported workbook is missing worksheet ${sourceSheet.name}`)
    if (metadata.pageSetup) {
      for (const key of ['printArea', 'printTitlesRow', 'printTitlesColumn'] as const) {
        if (metadata.pageSetup[key] !== undefined && exportedSheet.pageSetup[key] !== metadata.pageSetup[key])
          throw new Error(`The exported workbook did not preserve ${key}`)
      }
    }
    validateExportedTables(metadata.tables ?? [], exportedSheet)
  }
}

function assertNoUnsupportedResources(workbook: IWorkbookData): void {
  const unsupported = (workbook.resources ?? []).filter((resource) => {
    return UNSUPPORTED_RESOURCE_NAMES.has(resource.name) && hasMeaningfulResourceData(resource.data)
  })
  if (unsupported.length === 0)
    return
  const names = unsupported.map(resource => resource.name).join(', ')
  throw new Error(`Cannot save spreadsheet features that are not supported by XLSX export yet: ${names}`)
}

function hasMeaningfulResourceData(data: string): boolean {
  if (!data.trim())
    return false
  try {
    return hasMeaningfulValue(JSON.parse(data))
  }
  catch {
    return true
  }
}

function hasMeaningfulValue(value: unknown): boolean {
  if (value === null || value === undefined || value === '')
    return false
  if (Array.isArray(value))
    return value.some(hasMeaningfulValue)
  if (typeof value === 'object')
    return Object.values(value).some(hasMeaningfulValue)
  return true
}

function validateExportedTables(expected: ExcelTableMetadata[], worksheet: Worksheet): void {
  const actual = (worksheet.getTables() as unknown as ExcelTable[])
    .map(table => (table as unknown as { model: ExcelTableModel }).model)
  if (actual.length !== expected.length)
    throw new Error(`The exported worksheet did not preserve all ${expected.length} Excel tables`)

  for (const expectedTable of expected) {
    const actualTable = actual.find(table => table.name === expectedTable.name)
    if (!actualTable || actualTable.tableRef !== expectedTable.ref || actualTable.columns.length !== expectedTable.columns.length)
      throw new Error(`The exported workbook did not preserve Excel table ${expectedTable.name}`)
    if ((actualTable.style?.theme ?? '') !== (expectedTable.style.theme ?? ''))
      throw new Error(`The exported workbook did not preserve the style of Excel table ${expectedTable.name}`)
  }
}

function applyWorkbookMetadata(workbook: ExcelWorkbook, workbookData: IWorkbookData): void {
  workbook.definedNames.model = definedNamesForExport(workbookData)
}

function definedNamesForExport(workbookData: IWorkbookData): DefinedNamesModel {
  const metadata = getExcelWorkbookMetadata(workbookData)
  const definedNames = metadata?.definedNames ?? []
  const sheetNameMap = new Map<string, string>()
  const referencedSheets = new Set<string>()

  for (const sheetId of workbookData.sheetOrder) {
    const sheet = workbookData.sheets[sheetId]
    const sheetMetadata = getExcelWorksheetMetadata(sheet)
    if (sheet?.name && sheetMetadata?.originalName)
      sheetNameMap.set(sheetMetadata.originalName, sheet.name)
  }
  for (const definedName of definedNames) {
    for (const range of definedName.ranges) {
      const originalName = rangeSheetName(range)
      if (originalName)
        referencedSheets.add(originalName)
    }
  }

  for (const sheetId of workbookData.sheetOrder) {
    const sheet = workbookData.sheets[sheetId]
    const sheetMetadata = getExcelWorksheetMetadata(sheet)
    if (!sheet || !sheetMetadata)
      continue
    const hasPrintRanges = Boolean(sheetMetadata.pageSetup?.printArea || sheetMetadata.pageSetup?.printTitlesRow || sheetMetadata.pageSetup?.printTitlesColumn)
    if (!referencedSheets.has(sheetMetadata.originalName) && !hasPrintRanges)
      continue
    if (sheet.rowCount !== sheetMetadata.originalRowCount || sheet.columnCount !== sheetMetadata.originalColumnCount)
      throw new Error('Cannot save row or column structure changes while preserved named ranges or print ranges are present')
  }

  return definedNames.map(definedName => ({
    name: definedName.name,
    ranges: definedName.ranges.map(range => remapRangeSheet(range, sheetNameMap)),
  }))
}

function remapRangeSheet(range: string, sheetNameMap: Map<string, string>): string {
  const separator = range.lastIndexOf('!')
  if (separator < 0)
    throw new Error(`Cannot preserve named range without a worksheet: ${range}`)
  const originalName = unquoteSheetName(range.slice(0, separator))
  const currentName = sheetNameMap.get(originalName)
  if (!currentName)
    throw new Error(`Cannot preserve a named range because worksheet ${originalName} was removed`)
  return `${quoteSheetName(currentName)}!${range.slice(separator + 1)}`
}

function rangeSheetName(range: string): string | undefined {
  const separator = range.lastIndexOf('!')
  return separator < 0 ? undefined : unquoteSheetName(range.slice(0, separator))
}

function unquoteSheetName(name: string): string {
  return name.startsWith('\'') && name.endsWith('\'')
    ? name.slice(1, -1).replaceAll('\'\'', '\'')
    : name
}

function quoteSheetName(name: string): string {
  return `'${name.replaceAll('\'', '\'\'')}'`
}

function normalizedDefinedNames(model: DefinedNamesModel): DefinedNamesModel {
  return model
    .map(item => ({ name: item.name, ranges: [...item.ranges].sort() }))
    .sort((left, right) => left.name.localeCompare(right.name))
}

function getExcelWorkbookMetadata(workbook: IWorkbookData): ExcelWorkbookMetadata | undefined {
  return workbook.custom?.excel as ExcelWorkbookMetadata | undefined
}

function getExcelWorksheetMetadata(sheet?: Partial<IWorksheetData>): ExcelWorksheetMetadata | undefined {
  return sheet?.custom?.excel as ExcelWorksheetMetadata | undefined
}

function serializableClone<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T
}

function countPopulatedCells(worksheet: Worksheet): number {
  let count = 0
  worksheet.eachRow((row) => {
    row.eachCell((cell) => {
      if (cell.value !== null)
        count++
    })
  })
  return count
}

export async function createEmptyXlsx(): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook()
  workbook.addWorksheet('Sheet1')
  const output = await workbook.xlsx.writeBuffer()
  return new Uint8Array(output as ArrayBuffer).slice().buffer
}
