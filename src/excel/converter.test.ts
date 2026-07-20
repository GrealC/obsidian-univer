import type { IWorkbookData } from '@univerjs/core'
import { BooleanNumber, CellValueType, CustomRangeType, HorizontalAlign, LocaleType } from '@univerjs/core'
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
