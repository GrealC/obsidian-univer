import JSZip from 'jszip'

export interface XlsxInspection {
  readOnly: boolean
  reasons: string[]
  sheetCount: number
}

const unsupportedParts: Array<[RegExp, string]> = [
  [/^xl\/charts\//i, 'charts'],
  [/^xl\/pivotTables\//i, 'pivot tables'],
  [/^xl\/pivotCache\//i, 'pivot caches'],
  [/^xl\/externalLinks\//i, 'external links'],
  [/^xl\/drawings\//i, 'drawings or images'],
  [/^xl\/(threadedComments|comments)\d*\.xml$/i, 'comments'],
  [/^xl\/queryTables\//i, 'query tables'],
  [/^xl\/embeddings\//i, 'embedded objects'],
  [/^xl\/activeX\//i, 'ActiveX controls'],
  [/^xl\/ctrlProps\//i, 'form controls'],
  [/^xl\/slicers\//i, 'slicers'],
  [/^xl\/connections\.xml$/i, 'data connections'],
  [/^xl\/metadata\.xml$/i, 'advanced formula metadata'],
  [/^xl\/vbaProject\.bin$/i, 'VBA macros'],
]

const unsupportedWorksheetMarkup: Array<[RegExp, string]> = [
  [/<conditionalFormatting\b/i, 'conditional formatting'],
  [/<dataValidations\b/i, 'data validation rules'],
  [/<autoFilter\b/i, 'auto filters'],
  [/<sheetProtection\b/i, 'worksheet protection'],
  [/<(?:rowBreaks|colBreaks)\b/i, 'manual page breaks'],
  [/<legacyDrawing\b/i, 'legacy comments or drawings'],
]

export async function inspectXlsx(buffer: ArrayBuffer): Promise<XlsxInspection> {
  if (buffer.byteLength === 0)
    return { readOnly: false, reasons: [], sheetCount: 0 }

  const zip = await JSZip.loadAsync(buffer)
  const fileNames = Object.keys(zip.files)
  const reasons = new Set<string>()

  for (const fileName of fileNames) {
    for (const [pattern, reason] of unsupportedParts) {
      if (pattern.test(fileName))
        reasons.add(reason)
    }
  }

  const worksheetNames = fileNames.filter(name => /^xl\/worksheets\/sheet\d+\.xml$/i.test(name))
  await Promise.all(worksheetNames.map(async (name) => {
    const xml = await zip.file(name)?.async('text')
    if (!xml)
      return

    for (const [pattern, reason] of unsupportedWorksheetMarkup) {
      if (pattern.test(xml))
        reasons.add(reason)
    }
  }))

  const tableNames = fileNames.filter(name => /^xl\/tables\/[^/]+\.xml$/i.test(name))
  await Promise.all(tableNames.map(async (name) => {
    const xml = await zip.file(name)?.async('text')
    if (xml)
      inspectTable(xml, reasons)
  }))

  const workbookXml = await zip.file('xl/workbook.xml')?.async('text')
  if (workbookXml)
    inspectDefinedNames(workbookXml, reasons)

  return {
    readOnly: reasons.size > 0,
    reasons: [...reasons].sort(),
    sheetCount: worksheetNames.length,
  }
}

function inspectTable(tableXml: string, reasons: Set<string>): void {
  if (/<(?:[a-z_][\w.-]*:)?(?:calculatedColumnFormula|totalsRowFormula)\b/i.test(tableXml))
    reasons.add('advanced table formulas')
  const tableTag = /<(?:[a-z_][\w.-]*:)?table(?:\s[^>]*)?>/i.exec(tableXml)?.[0] ?? ''
  if (/\btotalsRowCount\s*=\s*(["'])[1-9]\d*\1/i.test(tableTag))
    reasons.add('table totals rows')
  if (/<(?:[a-z_][\w.-]*:)?(?:extLst|xmlColumnPr)\b/i.test(tableXml))
    reasons.add('advanced table extensions')
}

function inspectDefinedNames(workbookXml: string, reasons: Set<string>): void {
  const definedNamePattern = /<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/gi
  for (const match of workbookXml.matchAll(definedNamePattern)) {
    const attributes = match[1]
    const name = xmlAttribute(attributes, 'name')
    const expression = decodeXmlEntities(match[2].trim())

    if (name === '_xlnm.Print_Area' || name === '_xlnm.Print_Titles')
      continue
    if (!name || xmlAttribute(attributes, 'localSheetId') !== undefined) {
      reasons.add('sheet-scoped named ranges')
      continue
    }

    const extraAttributes = attributes.replace(/\bname\s*=\s*"[^"]*"/i, '').trim()
    if (extraAttributes) {
      reasons.add('advanced named ranges')
      continue
    }
    if (!isGlobalA1RangeExpression(expression))
      reasons.add('formula-based named ranges')
  }
}

function xmlAttribute(attributes: string, name: string): string | undefined {
  const match = new RegExp(`\\b${name}\\s*=\\s*"([^"]*)"`, 'i').exec(attributes)
  return match ? decodeXmlEntities(match[1]) : undefined
}

function decodeXmlEntities(value: string): string {
  return value
    .replace(/&#x([0-9a-f]+);/gi, (_, hex: string) => String.fromCodePoint(Number.parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, decimal: string) => String.fromCodePoint(Number.parseInt(decimal, 10)))
    .replaceAll('&quot;', '"')
    .replaceAll('&apos;', '\'')
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&amp;', '&')
}

function isGlobalA1RangeExpression(expression: string): boolean {
  const ranges = splitRangeList(expression)
  const sheet = String.raw`(?:'(?:[^']|'')+'|[^'!,]+)!`
  const cellRange = String.raw`\$[A-Z]{1,3}\$\d+(?::\$[A-Z]{1,3}\$\d+)?`
  const columnRange = String.raw`\$[A-Z]{1,3}:\$[A-Z]{1,3}`
  const rowRange = String.raw`\$\d+:\$\d+`
  const pattern = new RegExp(`^${sheet}(?:${cellRange}|${columnRange}|${rowRange})$`, 'i')
  return ranges.length > 0 && ranges.every(range => pattern.test(range))
}

function splitRangeList(expression: string): string[] {
  const ranges: string[] = []
  let current = ''
  let quoted = false
  for (let index = 0; index < expression.length; index++) {
    const character = expression[index]
    if (character === '\'' && expression[index + 1] === '\'') {
      current += '\'\''
      index++
      continue
    }
    if (character === '\'')
      quoted = !quoted
    if (character === ',' && !quoted) {
      ranges.push(current.trim())
      current = ''
    }
    else {
      current += character
    }
  }
  if (current.trim())
    ranges.push(current.trim())
  return ranges
}

export async function validateXlsxBuffer(buffer: ArrayBuffer): Promise<void> {
  if (buffer.byteLength === 0)
    throw new Error('The exported workbook is empty')

  const zip = await JSZip.loadAsync(buffer)
  if (!zip.file('[Content_Types].xml') || !zip.file('xl/workbook.xml'))
    throw new Error('The exported workbook is missing required XLSX metadata')

  const worksheetCount = Object.keys(zip.files).filter(name => /^xl\/worksheets\/sheet\d+\.xml$/i.test(name)).length
  if (worksheetCount === 0)
    throw new Error('The exported workbook does not contain a worksheet')
}
