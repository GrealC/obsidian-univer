import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { inspectXlsx, validateXlsxBuffer } from './capabilities'

async function workbookZip(worksheetXml = '<worksheet><sheetData /></worksheet>', extra?: (zip: JSZip) => void): Promise<ArrayBuffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', '<Types />')
  zip.file('xl/workbook.xml', '<workbook />')
  zip.file('xl/worksheets/sheet1.xml', worksheetXml)
  extra?.(zip)
  return zip.generateAsync({ type: 'arraybuffer' })
}

describe('xlsx capability inspection', () => {
  it('allows a basic workbook', async () => {
    const buffer = await workbookZip()
    await expect(validateXlsxBuffer(buffer)).resolves.toBeUndefined()
    await expect(inspectXlsx(buffer)).resolves.toEqual({
      readOnly: false,
      reasons: [],
      sheetCount: 1,
    })
  })

  it('makes workbooks with charts read-only', async () => {
    const buffer = await workbookZip(undefined, zip => zip.file('xl/charts/chart1.xml', '<chart />'))
    const inspection = await inspectXlsx(buffer)
    expect(inspection.readOnly).toBe(true)
    expect(inspection.reasons).toContain('charts')
  })

  it('makes unsupported worksheet features read-only', async () => {
    const buffer = await workbookZip('<worksheet><conditionalFormatting /><dataValidations /></worksheet>')
    const inspection = await inspectXlsx(buffer)
    expect(inspection.reasons).toEqual(['conditional formatting', 'data validation rules'])
  })

  it('allows basic structured tables', async () => {
    const buffer = await workbookZip('<worksheet><sheetData /><tableParts count="1"><tablePart r:id="rId1" /></tableParts></worksheet>', (zip) => {
      zip.file('xl/tables/table1.xml', `<table name="Jobs" displayName="Jobs" ref="A1:B2" headerRowCount="1" totalsRowShown="0">
        <autoFilter ref="A1:B2" />
        <tableColumns count="2"><tableColumn id="1" name="Company" /><tableColumn id="2" name="Role" /></tableColumns>
        <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0" />
      </table>`)
    })

    expect(await inspectXlsx(buffer)).toEqual({ readOnly: false, reasons: [], sheetCount: 1 })
  })

  it('protects table formulas and totals rows that cannot be rebuilt safely', async () => {
    const buffer = await workbookZip(undefined, (zip) => {
      zip.file('xl/tables/table1.xml', `<table name="Jobs" displayName="Jobs" ref="A1:B3" headerRowCount="1" totalsRowCount="1">
        <tableColumns count="2">
          <tableColumn id="1" name="Company" totalsRowLabel="Total" />
          <tableColumn id="2" name="Salary"><calculatedColumnFormula>1+1</calculatedColumnFormula></tableColumn>
        </tableColumns>
      </table>`)
    })

    expect((await inspectXlsx(buffer)).reasons).toEqual(['advanced table formulas', 'table totals rows'])
  })

  it('allows global A1 named ranges and print ranges', async () => {
    const buffer = await workbookZip(undefined, (zip) => {
      zip.file('xl/workbook.xml', `<workbook><definedNames>
        <definedName name="Items">'Data Sheet'!$A$1:$A$20</definedName>
        <definedName name="_xlnm.Print_Area" localSheetId="0">'Data Sheet'!$A$1:$E$20</definedName>
      </definedNames></workbook>`)
    })
    expect((await inspectXlsx(buffer)).readOnly).toBe(false)
  })

  it('protects named ranges ExcelJS cannot preserve safely', async () => {
    const buffer = await workbookZip(undefined, (zip) => {
      zip.file('xl/workbook.xml', `<workbook><definedNames>
        <definedName name="TaxRate">0.2</definedName>
        <definedName name="LocalItems" localSheetId="0">Sheet1!$A$1:$A$2</definedName>
        <definedName name="HiddenItems" hidden="1">Sheet1!$A$1:$A$2</definedName>
      </definedNames></workbook>`)
    })
    expect((await inspectXlsx(buffer)).reasons).toEqual([
      'advanced named ranges',
      'formula-based named ranges',
      'sheet-scoped named ranges',
    ])
  })

  it('rejects a malformed exported workbook', async () => {
    const zip = new JSZip()
    zip.file('readme.txt', 'not a workbook')
    await expect(validateXlsxBuffer(await zip.generateAsync({ type: 'arraybuffer' }))).rejects.toThrow('required XLSX metadata')
  })
})
