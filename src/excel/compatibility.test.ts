import { LocaleType } from '@univerjs/core'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { normalizeXlsxForExcelJs } from './compatibility'
import { importXlsx } from './converter'

const SPREADSHEETML_NAMESPACE = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

describe('excelJS OOXML compatibility', () => {
  it('normalizes a legal SpreadsheetML element prefix before ExcelJS parses the workbook', async () => {
    const source = await workbookWithTables()
    const prefixed = await prefixSpreadsheetMlElements(source)
    const excel = new ExcelJS.Workbook()

    await expect(excel.xlsx.load(asExcelBuffer(prefixed))).rejects.toThrow()

    const normalized = await normalizeXlsxForExcelJs(prefixed)
    await expect(excel.xlsx.load(asExcelBuffer(normalized))).resolves.toBeDefined()
    expect(excel.worksheets).toHaveLength(2)
  })

  it('normalizes absolute internal table relationship targets', async () => {
    const source = await workbookWithTables()
    const absoluteTargets = await makeTableRelationshipTargetsAbsolute(source)
    const excel = new ExcelJS.Workbook()

    await expect(excel.xlsx.load(asExcelBuffer(absoluteTargets))).rejects.toThrow()

    const normalized = await normalizeXlsxForExcelJs(absoluteTargets)
    await excel.xlsx.load(asExcelBuffer(normalized))
    expect(excel.worksheets.map(sheet => sheet.getTables().length)).toEqual([1, 1])
  })

  it('imports workbook and table data after both compatibility normalizations', async () => {
    const source = await workbookWithTables()
    const prefixed = await prefixSpreadsheetMlElements(source)
    const reportedShape = await makeTableRelationshipTargetsAbsolute(prefixed)
    const workbook = await importXlsx(reportedShape, 'Recruitment research', LocaleType.ZH_CN)

    expect(workbook.sheetOrder).toHaveLength(2)
    expect(workbook.sheetOrder.every(sheetId => workbook.sheets[sheetId]?.id === sheetId)).toBe(true)
    const first = workbook.sheets[workbook.sheetOrder[0]]
    const second = workbook.sheets[workbook.sheetOrder[1]]
    expect(first.name).toBe('Research')
    expect(first.cellData?.[1]?.[0]?.v).toBe('Company A')
    expect(second.name).toBe('Field guide')
    expect(second.cellData?.[1]?.[1]?.v).toBe('Description')
  })
})

async function workbookWithTables(): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook()
  const research = workbook.addWorksheet('Research')
  research.addTable({
    name: 'ResearchJobsTable',
    ref: 'A1:B1',
    headerRow: true,
    totalsRow: false,
    style: { theme: 'TableStyleMedium2', showRowStripes: true },
    columns: [{ name: 'Company' }, { name: 'Role' }],
    rows: [['Company A', 'Researcher']],
  })
  const guide = workbook.addWorksheet('Field guide')
  guide.addTable({
    name: 'FieldGuideTable',
    ref: 'A1:B1',
    headerRow: true,
    totalsRow: false,
    style: { theme: 'TableStyleMedium2', showRowStripes: true },
    columns: [{ name: 'Field' }, { name: 'Meaning' }],
    rows: [['Role', 'Description']],
  })
  const output = await workbook.xlsx.writeBuffer()
  return new Uint8Array(output as ArrayBuffer).slice().buffer
}

async function prefixSpreadsheetMlElements(buffer: ArrayBuffer): Promise<ArrayBuffer> {
  return mutateZip(buffer, async (name, source) => {
    if (!name.toLowerCase().endsWith('.xml') || !source.includes(`xmlns="${SPREADSHEETML_NAMESPACE}"`))
      return source
    return source
      .replace(`xmlns="${SPREADSHEETML_NAMESPACE}"`, `xmlns:x="${SPREADSHEETML_NAMESPACE}"`)
      .replace(/(<\/?)([a-z_][\w.-]*)(?=[\s/>])/gi, '$1x:$2')
  })
}

async function makeTableRelationshipTargetsAbsolute(buffer: ArrayBuffer): Promise<ArrayBuffer> {
  return mutateZip(buffer, async (name, source) => {
    if (!/^xl\/worksheets\/_rels\/.*\.rels$/i.test(name))
      return source
    return source.replace(/Target="\.\.\/tables\/(table\d+\.xml)"/g, 'Target="/xl/tables/$1"')
  })
}

async function mutateZip(buffer: ArrayBuffer, mutate: (name: string, source: string) => Promise<string>): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(buffer)
  for (const entry of Object.values(zip.files)) {
    if (entry.dir || (!entry.name.toLowerCase().endsWith('.xml') && !entry.name.toLowerCase().endsWith('.rels')))
      continue
    zip.file(entry.name, await mutate(entry.name, await entry.async('text')))
  }
  return zip.generateAsync({ type: 'arraybuffer' })
}

function asExcelBuffer(buffer: ArrayBuffer): ExcelJS.Buffer {
  return new Uint8Array(buffer) as unknown as ExcelJS.Buffer
}
