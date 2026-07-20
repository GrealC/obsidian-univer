import type { ICellData, IWorkbookData } from '@univerjs/core'
import { BooleanNumber, BorderStyleTypes, BorderType, CellValueType, CustomRangeType, HorizontalAlign, LocaleType, LogLevel, Univer } from '@univerjs/core'
import { FUniver } from '@univerjs/core/facade'
import { UniverRenderEnginePlugin, UniverSheetsPlugin } from '@univerjs/preset-sheets-core'
import ExcelJS from 'exceljs'
import { describe, expect, it } from 'vitest'
import { inspectXlsx } from './capabilities'
import { createEmptyXlsx, exportXlsx, importXlsx } from './converter'

function workbookFixture(): IWorkbookData {
  return {
    id: 'workbook-1',
    name: 'Round trip',
    appVersion: '0.25.1',
    locale: LocaleType.EN_US,
    styles: {
      heading: {
        bl: BooleanNumber.TRUE,
        ht: HorizontalAlign.CENTER,
        n: { pattern: '0.00' },
      },
    },
    sheetOrder: ['sheet-1'],
    sheets: {
      'sheet-1': {
        id: 'sheet-1',
        name: 'Data',
        rowCount: 100,
        columnCount: 26,
        cellData: {
          0: {
            0: { v: 210.6, t: CellValueType.NUMBER, s: 'heading' },
            1: { v: 382.2, t: CellValueType.NUMBER },
            2: { v: 592.8, t: CellValueType.NUMBER, f: '=SUM(A1:B1)' },
          },
          1: {
            0: { v: '00123', t: CellValueType.STRING },
            1: { v: true, t: CellValueType.BOOLEAN },
          },
        },
        mergeData: [{ startRow: 2, startColumn: 0, endRow: 2, endColumn: 1 }],
        rowData: { 0: { h: 30 } },
        columnData: { 0: { w: 120 } },
        freeze: { xSplit: 1, ySplit: 1, startColumn: 1, startRow: 1 },
        hidden: BooleanNumber.FALSE,
        showGridlines: BooleanNumber.TRUE,
      },
    },
  }
}

describe('local xlsx converter', () => {
  it('creates a valid empty workbook', async () => {
    const buffer = await createEmptyXlsx()
    const inspection = await inspectXlsx(buffer)
    const imported = await importXlsx(buffer, 'Blank', LocaleType.EN_US)
    const sheet = imported.sheets[imported.sheetOrder[0]]
    expect(inspection.sheetCount).toBe(1)
    expect(inspection.readOnly).toBe(false)
    expect(sheet.name).toBe('Sheet1')
    expect(sheet.rowCount).toBe(100)
    expect(sheet.columnCount).toBe(26)
    expect(sheet.cellData).toEqual({})
    const defaultStyle = typeof sheet.defaultStyle === 'string' ? imported.styles[sheet.defaultStyle] : sheet.defaultStyle
    expect(defaultStyle).toMatchObject({ ff: 'Calibri', fs: 11 })
  })

  it('opens legacy zero-byte workbooks as a blank sheet', async () => {
    const imported = await importXlsx(new ArrayBuffer(0), 'Blank', LocaleType.EN_US)
    expect(imported.sheetOrder).toHaveLength(1)
    expect(imported.sheets[imported.sheetOrder[0]].name).toBe('Sheet1')
  })

  it('round-trips common workbook data without numeric drift', async () => {
    const output = await exportXlsx(workbookFixture())
    const imported = await importXlsx(output, 'Round trip', LocaleType.EN_US)
    const sheet = imported.sheets[imported.sheetOrder[0]]

    expect(sheet.name).toBe('Data')
    expect(sheet.cellData?.[0]?.[0]?.v).toBe(210.6)
    expect(sheet.cellData?.[0]?.[1]?.v).toBe(382.2)
    expect(sheet.cellData?.[0]?.[2]?.f).toBe('=SUM(A1:B1)')
    expect(sheet.cellData?.[1]?.[0]?.v).toBe('00123')
    expect(sheet.mergeData).toEqual([{ startRow: 2, startColumn: 0, endRow: 2, endColumn: 1 }])
  })

  it('applies the Univer worksheet default font to exported cells', async () => {
    const source = workbookFixture()
    source.styles.default = { ff: 'Microsoft YaHei', fs: 12 }
    source.sheets['sheet-1'].defaultStyle = 'default'

    const output = await exportXlsx(source)
    const verification = new ExcelJS.Workbook()
    await verification.xlsx.load(new Uint8Array(output) as unknown as ExcelJS.Buffer)
    expect(verification.getWorksheet('Data')?.getCell('B1').font).toMatchObject({ name: 'Microsoft YaHei', size: 12 })
  })

  it('imports translated shared formulas instead of dropping follower cells', async () => {
    const excel = new ExcelJS.Workbook()
    const sheet = excel.addWorksheet('Shared formulas')
    sheet.getCell('A1').value = 1
    sheet.getCell('B1').value = 2
    sheet.getCell('A2').value = 3
    sheet.getCell('B2').value = 4
    sheet.fillFormula('C1:C2', 'A1+B1', [3, 7])

    const raw = await excel.xlsx.writeBuffer()
    const imported = await importXlsx(new Uint8Array(raw as ArrayBuffer).slice().buffer, 'Shared formulas', LocaleType.EN_US)
    const importedSheet = imported.sheets[imported.sheetOrder[0]]

    expect(importedSheet.cellData?.[0]?.[2]?.f).toBe('=A1+B1')
    expect(importedSheet.cellData?.[1]?.[2]?.f).toBe('=A2+B2')
  })

  it('round-trips global named ranges and print settings', async () => {
    const excel = new ExcelJS.Workbook()
    const sheet = excel.addWorksheet('Data Sheet')
    sheet.getCell('A1').value = 'one'
    sheet.getCell('A2').value = 'two'
    excel.definedNames.add('\'Data Sheet\'!$A$1:$A$2', 'Items')
    sheet.pageSetup.printArea = 'A1:B3'
    sheet.pageSetup.printTitlesRow = '1:1'
    sheet.headerFooter.oddHeader = '&CResearch notes'

    const source = new Uint8Array(await excel.xlsx.writeBuffer() as ArrayBuffer).slice().buffer
    expect((await inspectXlsx(source)).readOnly).toBe(false)
    const imported = await importXlsx(source, 'Named ranges', LocaleType.EN_US)
    const output = await exportXlsx(imported)
    const verification = new ExcelJS.Workbook()
    await verification.xlsx.load(new Uint8Array(output) as unknown as ExcelJS.Buffer)

    expect(verification.definedNames.model).toEqual([{ name: 'Items', ranges: ['\'Data Sheet\'!$A$1:$A$2'] }])
    expect(verification.getWorksheet('Data Sheet')?.pageSetup.printArea).toBe('A1:B3')
    expect(verification.getWorksheet('Data Sheet')?.pageSetup.printTitlesRow).toBe('1:1')
    expect(verification.getWorksheet('Data Sheet')?.headerFooter.oddHeader).toBe('&CResearch notes')
  })

  it('round-trips editable basic tables and their current cell values', async () => {
    const excel = new ExcelJS.Workbook()
    const sheet = excel.addWorksheet('Research')
    sheet.addTable({
      name: 'ResearchJobsTable',
      ref: 'A2',
      headerRow: true,
      totalsRow: false,
      style: { theme: 'TableStyleMedium2', showRowStripes: true },
      columns: [{ name: 'Company' }, { name: 'Role' }],
      rows: [['Company A', 'Researcher'], ['Company B', 'Engineer']],
    })

    const source = new Uint8Array(await excel.xlsx.writeBuffer() as ArrayBuffer).slice().buffer
    expect(await inspectXlsx(source)).toEqual({ readOnly: false, reasons: [], sheetCount: 1 })
    const imported = await importXlsx(source, 'Tables', LocaleType.EN_US)
    const importedSheet = imported.sheets[imported.sheetOrder[0]]
    expect((importedSheet.custom?.excel as { tables?: unknown[] }).tables).toHaveLength(1)

    importedSheet.cellData![2]![1]!.v = 'Senior Researcher'
    const output = await exportXlsx(imported)
    const verification = new ExcelJS.Workbook()
    await verification.xlsx.load(new Uint8Array(output) as unknown as ExcelJS.Buffer)
    const exportedSheet = verification.getWorksheet('Research')!
    const tableModel = (exportedSheet.getTables() as unknown as Array<{ model: { name: string, tableRef: string, style: { theme?: string } } }>)[0].model
    expect(tableModel.name).toBe('ResearchJobsTable')
    expect(tableModel.tableRef).toBe('A2:B4')
    expect(tableModel.style.theme).toBe('TableStyleMedium2')
    expect(exportedSheet.getCell('B3').value).toBe('Senior Researcher')
  })

  it('round-trips gridline and right-to-left view settings', async () => {
    const excel = new ExcelJS.Workbook()
    const sheet = excel.addWorksheet('View')
    sheet.views = [{ state: 'normal', showGridLines: false, rightToLeft: true }]
    sheet.getCell('A1').value = 'value'

    const source = new Uint8Array(await excel.xlsx.writeBuffer() as ArrayBuffer).slice().buffer
    const imported = await importXlsx(source, 'View settings', LocaleType.EN_US)
    const importedSheet = imported.sheets[imported.sheetOrder[0]]
    expect(importedSheet.showGridlines).toBe(BooleanNumber.FALSE)
    expect(importedSheet.rightToLeft).toBe(BooleanNumber.TRUE)

    const output = await exportXlsx(imported)
    const verification = new ExcelJS.Workbook()
    await verification.xlsx.load(new Uint8Array(output) as unknown as ExcelJS.Buffer)
    expect(verification.getWorksheet('View')?.views[0]?.showGridLines).toBe(false)
    expect(verification.getWorksheet('View')?.views[0]?.rightToLeft).toBe(true)
  })

  it('keeps cleared borders absent when the saved workbook is reopened', async () => {
    const excel = new ExcelJS.Workbook()
    const sheet = excel.addWorksheet('Borders')
    for (let row = 1; row <= 4; row++) {
      for (let column = 1; column <= 4; column++) {
        const cell = sheet.getCell(row, column)
        cell.value = row === 1 ? `Heading ${column}` : null
        cell.border = {
          top: { style: 'thin' },
          right: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
        }
      }
    }

    const source = new Uint8Array(await excel.xlsx.writeBuffer() as ArrayBuffer).slice().buffer
    const imported = await importXlsx(source, 'Borders', LocaleType.EN_US)
    const importedSheet = imported.sheets[imported.sheetOrder[0]]
    const importedRows = importedSheet.cellData as Record<number, Record<number, ICellData | undefined>>
    for (const row of Object.values(importedRows)) {
      for (const cell of Object.values(row)) {
        if (!cell)
          continue
        const style = typeof cell.s === 'string' ? imported.styles[cell.s] : cell.s
        cell.s = {
          ...style,
          bd: {
            t: { s: BorderStyleTypes.NONE, cl: { rgb: '#000000' } },
            r: { s: BorderStyleTypes.NONE, cl: { rgb: '#000000' } },
            b: { s: BorderStyleTypes.NONE, cl: { rgb: '#000000' } },
            l: { s: BorderStyleTypes.NONE, cl: { rgb: '#000000' } },
          },
        }
      }
    }

    const output = await exportXlsx(imported)
    const verification = new ExcelJS.Workbook()
    await verification.xlsx.load(new Uint8Array(output) as unknown as ExcelJS.Buffer)
    expect(verification.getWorksheet('Borders')?.getCell('A1').border.top?.style).toBeUndefined()

    const reopened = await importXlsx(output, 'Borders', LocaleType.EN_US)
    expect(Object.values(reopened.styles).every(style => !style?.bd)).toBe(true)
    expect(reopened.sheets[reopened.sheetOrder[0]].gridlinesColor).toBe('#D6D8DB')
  })

  it('round-trips formatting produced by real Univer range commands', async () => {
    const excel = new ExcelJS.Workbook()
    const sheet = excel.addWorksheet('Formatting')
    for (let row = 1; row <= 2; row++) {
      for (let column = 1; column <= 2; column++) {
        const cell = sheet.getCell(row, column)
        cell.value = `${row},${column}`
        cell.font = { name: 'Arial' }
        cell.alignment = { horizontal: 'left' }
        cell.border = {
          top: { style: 'thin' },
          right: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
        }
      }
    }

    const source = new Uint8Array(await excel.xlsx.writeBuffer() as ArrayBuffer).slice().buffer
    const imported = await importXlsx(source, 'Formatting', LocaleType.EN_US)
    const univer = new Univer({
      locale: LocaleType.EN_US,
      locales: { [LocaleType.EN_US]: {} },
      logLevel: LogLevel.SILENT,
    })
    univer.registerPlugin(UniverRenderEnginePlugin)
    univer.registerPlugin(UniverSheetsPlugin)
    const univerAPI = FUniver.newAPI(univer)

    try {
      const workbook = univerAPI.createWorkbook(imported)
      const range = workbook.getActiveSheet().getRange('A1:B2')
      range.setBorder(BorderType.NONE, BorderStyleTypes.NONE)
      range.setFontFamily('Microsoft YaHei')
      range.setHorizontalAlignment('center')

      const snapshot = workbook.save()
      const output = await exportXlsx(snapshot)
      const verification = new ExcelJS.Workbook()
      await verification.xlsx.load(new Uint8Array(output) as unknown as ExcelJS.Buffer)
      const reopened = verification.getWorksheet('Formatting')!

      expect(reopened.getCell('A1').border.top?.style).toBeUndefined()
      expect(reopened.getCell('A1').font.name).toBe('Microsoft YaHei')
      expect(reopened.getCell('A1').alignment.horizontal).toBe('center')

      const reimported = await importXlsx(output, 'Formatting', LocaleType.EN_US)
      const reopenedSheet = reimported.sheets[reimported.sheetOrder[0]]
      const reopenedCell = reopenedSheet.cellData?.[0]?.[0]
      const reopenedStyle = typeof reopenedCell?.s === 'string' ? reimported.styles[reopenedCell.s] : reopenedCell?.s
      expect(reopenedStyle?.bd).toBeUndefined()
      expect(reopenedStyle?.ff).toBe('Microsoft YaHei')
      expect(reopenedStyle?.ht).toBe(HorizontalAlign.CENTER)
    }
    finally {
      univer.dispose()
    }
  })

  it('keeps workbook style ids unique across multiple worksheets', async () => {
    const excel = new ExcelJS.Workbook()
    const first = excel.addWorksheet('First')
    first.getCell('A1').value = 'first'
    first.getCell('A1').font = { name: 'Arial', bold: true }
    const second = excel.addWorksheet('Second')
    second.getCell('A1').value = 'second'
    second.getCell('A1').font = { name: 'Microsoft YaHei', italic: true }

    const source = new Uint8Array(await excel.xlsx.writeBuffer() as ArrayBuffer).slice().buffer
    const imported = await importXlsx(source, 'Multiple sheets', LocaleType.EN_US)
    const firstCell = imported.sheets[imported.sheetOrder[0]].cellData?.[0]?.[0]
    const secondCell = imported.sheets[imported.sheetOrder[1]].cellData?.[0]?.[0]
    const firstStyle = typeof firstCell?.s === 'string' ? imported.styles[firstCell.s] : firstCell?.s
    const secondStyle = typeof secondCell?.s === 'string' ? imported.styles[secondCell.s] : secondCell?.s

    expect(firstCell?.s).not.toBe(secondCell?.s)
    expect(firstStyle).toMatchObject({ ff: 'Arial', bl: BooleanNumber.TRUE })
    expect(secondStyle).toMatchObject({ ff: 'Microsoft YaHei', it: BooleanNumber.TRUE })
  })

  it('round-trips in-cell rich text fonts instead of flattening them', async () => {
    const excel = new ExcelJS.Workbook()
    const sheet = excel.addWorksheet('Rich text')
    sheet.getCell('A1').value = {
      richText: [
        { text: '中文', font: { name: 'Microsoft YaHei', bold: true } },
        { text: ' and ' },
        { text: 'English', font: { name: 'Arial', italic: true } },
      ],
    }

    const source = new Uint8Array(await excel.xlsx.writeBuffer() as ArrayBuffer).slice().buffer
    const imported = await importXlsx(source, 'Rich text', LocaleType.EN_US)
    const importedCell = imported.sheets[imported.sheetOrder[0]].cellData?.[0]?.[0]
    expect(importedCell?.p?.body?.textRuns).toEqual([
      { st: 0, ed: 2, ts: expect.objectContaining({ ff: 'Microsoft YaHei', bl: BooleanNumber.TRUE }) },
      { st: 7, ed: 14, ts: expect.objectContaining({ ff: 'Arial', it: BooleanNumber.TRUE }) },
    ])

    const output = await exportXlsx(imported)
    const verification = new ExcelJS.Workbook()
    await verification.xlsx.load(new Uint8Array(output) as unknown as ExcelJS.Buffer)
    const value = verification.getWorksheet('Rich text')?.getCell('A1').value
    expect(value).toMatchObject({
      richText: [
        { text: '中文', font: expect.objectContaining({ name: 'Microsoft YaHei', bold: true }) },
        { text: ' and ' },
        { text: 'English', font: expect.objectContaining({ name: 'Arial', italic: true }) },
      ],
    })
  })

  it('preserves row and column styles carried outside cell data', async () => {
    const excel = new ExcelJS.Workbook()
    const sheet = excel.addWorksheet('Dimensions')
    sheet.getCell('A2').value = 'row styled'
    sheet.getCell('C1').value = 'column styled'
    sheet.getRow(2).font = { name: 'Microsoft YaHei', bold: true }
    sheet.getColumn(3).alignment = { horizontal: 'center' }

    const source = new Uint8Array(await excel.xlsx.writeBuffer() as ArrayBuffer).slice().buffer
    const imported = await importXlsx(source, 'Dimensions', LocaleType.EN_US)
    const importedSheet = imported.sheets[imported.sheetOrder[0]]
    expect(importedSheet.rowData?.[1]?.s).toBeDefined()
    expect(importedSheet.columnData?.[2]?.s).toBeDefined()

    const output = await exportXlsx(imported)
    const verification = new ExcelJS.Workbook()
    await verification.xlsx.load(new Uint8Array(output) as unknown as ExcelJS.Buffer)
    const reopened = verification.getWorksheet('Dimensions')!
    expect(reopened.getRow(2).font).toMatchObject({ name: 'Microsoft YaHei', bold: true })
    expect(reopened.getColumn(3).alignment?.horizontal).toBe('center')
  })

  it('exports hyperlinks created through the Univer hyperlink UI', async () => {
    const source = workbookFixture()
    source.sheets['sheet-1'].cellData![3] = {
      0: {
        t: CellValueType.STRING,
        p: {
          id: 'link-document',
          documentStyle: {},
          body: {
            dataStream: 'OpenAI\r\n',
            customRanges: [{
              rangeType: CustomRangeType.HYPERLINK,
              rangeId: 'link-1',
              startIndex: 0,
              endIndex: 5,
              properties: { url: 'https://openai.com' },
            }],
          },
        },
      },
    }

    const output = await exportXlsx(source)
    const verification = new ExcelJS.Workbook()
    await verification.xlsx.load(new Uint8Array(output) as unknown as ExcelJS.Buffer)
    expect(verification.getWorksheet('Data')?.getCell('A4').value).toEqual({
      text: 'OpenAI',
      hyperlink: 'https://openai.com',
    })
  })

  it('refuses to silently discard unsupported Univer plugin resources', async () => {
    const source = workbookFixture()
    source.resources = [
      { name: 'SHEET_FILTER_PLUGIN', data: '{}' },
      { name: 'SHEET_DATA_VALIDATION_PLUGIN', data: '{"sheet-1":[{"uid":"rule-1"}]}' },
    ]

    await expect(exportXlsx(source)).rejects.toThrow('SHEET_DATA_VALIDATION_PLUGIN')
  })
})
