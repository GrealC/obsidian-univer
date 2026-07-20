import type { IDocumentBody, IDocumentData, IParagraphStyle, ITextRun, ITextStyle } from '@univerjs/core'
import type { Element as XmlElement } from '@xmldom/xmldom'
import { BaselineOffset, BooleanNumber, HorizontalAlign, NamedStyleType, RichTextBuilder, TextDecoration } from '@univerjs/core'
import { DOMParser } from '@xmldom/xmldom'
import JSZip from 'jszip'
import { validateDocxBuffer } from './capabilities'

const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const A4_WIDTH_TWIPS = 11906
const A4_HEIGHT_TWIPS = 16838
const DEFAULT_MARGIN_TWIPS = 1440

const HIGHLIGHT_COLORS: Record<string, string> = {
  black: '#000000',
  blue: '#0000FF',
  cyan: '#00FFFF',
  darkBlue: '#000080',
  darkCyan: '#008080',
  darkGray: '#808080',
  darkGreen: '#008000',
  darkMagenta: '#800080',
  darkRed: '#800000',
  darkYellow: '#808000',
  green: '#00FF00',
  lightGray: '#C0C0C0',
  magenta: '#FF00FF',
  red: '#FF0000',
  white: '#FFFFFF',
  yellow: '#FFFF00',
}

export async function createEmptyDocx(title = 'Untitled'): Promise<ArrayBuffer> {
  const snapshot = createEmptyDocument(title)
  return exportDocx(snapshot)
}

export async function importDocx(buffer: ArrayBuffer, title: string): Promise<IDocumentData> {
  await validateDocxBuffer(buffer)
  const zip = await JSZip.loadAsync(buffer)
  const xml = await zip.file('word/document.xml')?.async('text')
  if (!xml)
    throw new Error('The DOCX document body is missing')

  const parsed = new DOMParser().parseFromString(xml, 'application/xml')
  if (parsed.getElementsByTagName('parsererror').length > 0)
    throw new Error('The DOCX document XML is invalid')

  const bodyElement = firstElement(parsed.documentElement, 'body')
  if (!bodyElement)
    throw new Error('The DOCX document body is invalid')

  const documentData = createEmptyDocument(title)
  const paragraphs = descendants(bodyElement, 'p')
  let dataStream = ''
  const textRuns: ITextRun[] = []
  const paragraphData: NonNullable<IDocumentBody['paragraphs']> = []

  for (const paragraph of paragraphs) {
    for (const run of descendants(paragraph, 'r')) {
      const text = readRunText(run)
      if (!text)
        continue
      const start = dataStream.length
      dataStream += text
      const style = readRunStyle(run)
      if (Object.keys(style).length > 0)
        textRuns.push({ st: start, ed: dataStream.length, ts: style })
    }
    dataStream += '\r'
    paragraphData.push({
      startIndex: dataStream.length - 1,
      paragraphStyle: readParagraphStyle(paragraph),
    })
  }

  if (paragraphData.length === 0) {
    dataStream = '\r'
    paragraphData.push({ startIndex: 0 })
  }
  dataStream += '\n'

  documentData.body = {
    dataStream,
    textRuns,
    paragraphs: paragraphData,
    sectionBreaks: [{ startIndex: dataStream.length - 1 }],
    customBlocks: [],
    tables: [],
  }
  documentData.documentStyle = {
    ...documentData.documentStyle,
    ...readSectionStyle(bodyElement),
  }
  return documentData
}

function createEmptyDocument(title: string): IDocumentData {
  const documentData = RichTextBuilder.newEmptyData()
  documentData.id = crypto.randomUUID()
  documentData.title = title
  documentData.documentStyle = {
    ...documentData.documentStyle,
    pageSize: { width: A4_WIDTH_TWIPS / 15, height: A4_HEIGHT_TWIPS / 15 },
    marginTop: DEFAULT_MARGIN_TWIPS / 15,
    marginBottom: DEFAULT_MARGIN_TWIPS / 15,
    marginLeft: DEFAULT_MARGIN_TWIPS / 15,
    marginRight: DEFAULT_MARGIN_TWIPS / 15,
  }
  return documentData
}

export async function exportDocx(documentData: IDocumentData, sourceBuffer?: ArrayBuffer): Promise<ArrayBuffer> {
  assertExportable(documentData)
  const zip = sourceBuffer ? await JSZip.loadAsync(sourceBuffer) : createDocxPackage(documentData.title ?? 'Untitled')
  zip.file('word/document.xml', buildDocumentXml(documentData))
  const output = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' })
  await validateDocxBuffer(output)
  return output
}

function assertExportable(documentData: IDocumentData): void {
  const body = documentData.body
  if ((body?.customBlocks?.length ?? 0) > 0 || (body?.tables?.length ?? 0) > 0)
    throw new Error('The document contains blocks that cannot be exported to DOCX safely')
  if ((body?.customRanges?.length ?? 0) > 0)
    throw new Error('The document contains links or ranges that cannot be exported to DOCX safely')
  if (documentData.drawings && Object.keys(documentData.drawings).length > 0)
    throw new Error('The document contains drawings that cannot be exported to DOCX safely')
}

function readRunText(run: XmlElement): string {
  let value = ''
  for (const node of childElements(run)) {
    if (node.localName === 't' || node.localName === 'delText')
      value += node.textContent ?? ''
    else if (node.localName === 'tab')
      value += '\t'
    else if (node.localName === 'br')
      value += attribute(node, 'type') === 'page' ? '\f' : '\n'
    else if (node.localName === 'noBreakHyphen')
      value += '\u2011'
    else if (node.localName === 'softHyphen')
      value += '\u00AD'
  }
  return value
}

function readRunStyle(run: XmlElement): ITextStyle {
  const properties = firstElement(run, 'rPr')
  if (!properties)
    return {}

  const style: ITextStyle = {}
  const fonts = firstElement(properties, 'rFonts')
  const family = fonts && (attribute(fonts, 'eastAsia') || attribute(fonts, 'ascii') || attribute(fonts, 'hAnsi'))
  if (family)
    style.ff = family
  const size = numericAttribute(firstElement(properties, 'sz'), 'val')
  if (size !== undefined)
    style.fs = size / 2
  if (enabled(firstElement(properties, 'b')))
    style.bl = BooleanNumber.TRUE
  if (enabled(firstElement(properties, 'i')))
    style.it = BooleanNumber.TRUE
  if (enabled(firstElement(properties, 'strike')))
    style.st = { s: BooleanNumber.TRUE, t: TextDecoration.SINGLE }

  const underline = firstElement(properties, 'u')
  if (underline && attribute(underline, 'val') !== 'none') {
    style.ul = {
      s: BooleanNumber.TRUE,
      t: attribute(underline, 'val') === 'double' ? TextDecoration.DOUBLE : TextDecoration.SINGLE,
    }
  }
  const color = normalizeColor(attribute(firstElement(properties, 'color'), 'val'))
  if (color)
    style.cl = { rgb: color }
  const highlight = attribute(firstElement(properties, 'highlight'), 'val')
  if (highlight && HIGHLIGHT_COLORS[highlight])
    style.bg = { rgb: HIGHLIGHT_COLORS[highlight] }
  const verticalAlign = attribute(firstElement(properties, 'vertAlign'), 'val')
  if (verticalAlign === 'superscript')
    style.va = BaselineOffset.SUPERSCRIPT
  else if (verticalAlign === 'subscript')
    style.va = BaselineOffset.SUBSCRIPT
  return style
}

function readParagraphStyle(paragraph: XmlElement): IParagraphStyle {
  const properties = firstElement(paragraph, 'pPr')
  if (!properties)
    return {}
  const style: IParagraphStyle = {}
  const namedStyle = namedStyleType(attribute(firstElement(properties, 'pStyle'), 'val'))
  if (namedStyle !== undefined)
    style.namedStyleType = namedStyle

  const alignment = attribute(firstElement(properties, 'jc'), 'val')
  const alignments: Record<string, HorizontalAlign> = {
    left: HorizontalAlign.LEFT,
    start: HorizontalAlign.LEFT,
    center: HorizontalAlign.CENTER,
    right: HorizontalAlign.RIGHT,
    end: HorizontalAlign.RIGHT,
    both: HorizontalAlign.JUSTIFIED,
    justify: HorizontalAlign.JUSTIFIED,
    distribute: HorizontalAlign.DISTRIBUTED,
  }
  if (alignment && alignments[alignment] !== undefined)
    style.horizontalAlign = alignments[alignment]

  const spacing = firstElement(properties, 'spacing')
  const before = numericAttribute(spacing, 'before')
  const after = numericAttribute(spacing, 'after')
  const line = numericAttribute(spacing, 'line')
  if (before !== undefined)
    style.spaceAbove = { v: before / 20 }
  if (after !== undefined)
    style.spaceBelow = { v: after / 20 }
  if (line !== undefined)
    style.lineSpacing = line / 240

  const indentation = firstElement(properties, 'ind')
  const left = numericAttribute(indentation, 'left') ?? numericAttribute(indentation, 'start')
  const right = numericAttribute(indentation, 'right') ?? numericAttribute(indentation, 'end')
  const firstLine = numericAttribute(indentation, 'firstLine')
  const hanging = numericAttribute(indentation, 'hanging')
  if (left !== undefined)
    style.indentStart = { v: left / 20 }
  if (right !== undefined)
    style.indentEnd = { v: right / 20 }
  if (firstLine !== undefined)
    style.indentFirstLine = { v: firstLine / 20 }
  if (hanging !== undefined)
    style.hanging = { v: hanging / 20 }
  if (firstElement(properties, 'keepNext'))
    style.keepNext = BooleanNumber.TRUE
  if (firstElement(properties, 'keepLines'))
    style.keepLines = BooleanNumber.TRUE
  return style
}

function readSectionStyle(body: XmlElement): IDocumentData['documentStyle'] {
  const section = descendants(body, 'sectPr').at(-1)
  if (!section)
    return {}
  const pageSize = firstElement(section, 'pgSz')
  const pageMargin = firstElement(section, 'pgMar')
  const width = numericAttribute(pageSize, 'w')
  const height = numericAttribute(pageSize, 'h')
  return {
    pageSize: width && height ? { width: width / 15, height: height / 15 } : undefined,
    marginTop: twipsToPixels(numericAttribute(pageMargin, 'top')),
    marginBottom: twipsToPixels(numericAttribute(pageMargin, 'bottom')),
    marginLeft: twipsToPixels(numericAttribute(pageMargin, 'left')),
    marginRight: twipsToPixels(numericAttribute(pageMargin, 'right')),
  }
}

function buildDocumentXml(documentData: IDocumentData): string {
  const body = documentData.body
  const dataStream = body?.dataStream ?? '\r\n'
  const paragraphs = [...(body?.paragraphs ?? [{ startIndex: Math.max(0, dataStream.indexOf('\r')) }])]
    .sort((left, right) => left.startIndex - right.startIndex)
  const textRuns = [...(body?.textRuns ?? [])].sort((left, right) => left.st - right.st)
  let cursor = 0
  const paragraphXml: string[] = []

  for (const paragraph of paragraphs) {
    const end = Math.max(cursor, Math.min(paragraph.startIndex, dataStream.length))
    const text = dataStream.slice(cursor, end)
    paragraphXml.push(`<w:p>${writeParagraphProperties(paragraph.paragraphStyle)}${writeRuns(text, cursor, textRuns)}</w:p>`)
    cursor = end + 1
  }
  if (paragraphXml.length === 0)
    paragraphXml.push('<w:p/>')

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${WORD_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>${paragraphXml.join('')}${writeSectionProperties(documentData)}</w:body>
</w:document>`
}

function writeRuns(text: string, absoluteStart: number, textRuns: ITextRun[]): string {
  if (!text)
    return ''
  const boundaries = new Set([0, text.length])
  for (const run of textRuns) {
    const start = Math.max(0, run.st - absoluteStart)
    const end = Math.min(text.length, run.ed - absoluteStart)
    if (start < end) {
      boundaries.add(start)
      boundaries.add(end)
    }
  }
  const points = [...boundaries].sort((left, right) => left - right)
  let xml = ''
  for (let index = 0; index < points.length - 1; index++) {
    const start = points[index]
    const end = points[index + 1]
    const value = text.slice(start, end)
    const style = textRuns.find(run => run.st <= absoluteStart + start && run.ed >= absoluteStart + end)?.ts
    xml += `<w:r>${writeRunProperties(style)}${writeRunContent(value)}</w:r>`
  }
  return xml
}

function writeRunContent(value: string): string {
  let xml = ''
  let text = ''
  const flush = () => {
    if (!text)
      return
    xml += `<w:t xml:space="preserve">${escapeXml(text)}</w:t>`
    text = ''
  }
  for (const character of value) {
    if (character === '\t') {
      flush()
      xml += '<w:tab/>'
    }
    else if (character === '\f') {
      flush()
      xml += '<w:br w:type="page"/>'
    }
    else if (character === '\n' || character === '\v') {
      flush()
      xml += '<w:br/>'
    }
    else {
      text += character
    }
  }
  flush()
  return xml
}

function writeRunProperties(style?: ITextStyle): string {
  if (!style)
    return ''
  const values: string[] = []
  if (style.ff)
    values.push(`<w:rFonts w:ascii="${escapeAttribute(style.ff)}" w:hAnsi="${escapeAttribute(style.ff)}" w:eastAsia="${escapeAttribute(style.ff)}"/>`)
  if (style.fs)
    values.push(`<w:sz w:val="${Math.round(style.fs * 2)}"/><w:szCs w:val="${Math.round(style.fs * 2)}"/>`)
  if (style.bl === BooleanNumber.TRUE)
    values.push('<w:b/>')
  if (style.it === BooleanNumber.TRUE)
    values.push('<w:i/>')
  if (style.ul?.s === BooleanNumber.TRUE)
    values.push(`<w:u w:val="${style.ul.t === TextDecoration.DOUBLE ? 'double' : 'single'}"/>`)
  if (style.st?.s === BooleanNumber.TRUE)
    values.push('<w:strike/>')
  const color = style.cl?.rgb?.replace('#', '')
  if (color)
    values.push(`<w:color w:val="${escapeAttribute(color.toUpperCase())}"/>`)
  if (style.va === BaselineOffset.SUPERSCRIPT)
    values.push('<w:vertAlign w:val="superscript"/>')
  else if (style.va === BaselineOffset.SUBSCRIPT)
    values.push('<w:vertAlign w:val="subscript"/>')
  return values.length > 0 ? `<w:rPr>${values.join('')}</w:rPr>` : ''
}

function writeParagraphProperties(style?: IParagraphStyle): string {
  if (!style)
    return ''
  const values: string[] = []
  const styleName = namedStyleName(style.namedStyleType)
  if (styleName)
    values.push(`<w:pStyle w:val="${styleName}"/>`)
  const alignments: Partial<Record<HorizontalAlign, string>> = {
    [HorizontalAlign.LEFT]: 'left',
    [HorizontalAlign.CENTER]: 'center',
    [HorizontalAlign.RIGHT]: 'right',
    [HorizontalAlign.JUSTIFIED]: 'both',
    [HorizontalAlign.BOTH]: 'both',
    [HorizontalAlign.DISTRIBUTED]: 'distribute',
  }
  const alignment = style.horizontalAlign && alignments[style.horizontalAlign]
  if (alignment)
    values.push(`<w:jc w:val="${alignment}"/>`)
  const before = pointValue(style.spaceAbove)
  const after = pointValue(style.spaceBelow)
  const line = style.lineSpacing
  if (before !== undefined || after !== undefined || line !== undefined)
    values.push(`<w:spacing${before !== undefined ? ` w:before="${Math.round(before * 20)}"` : ''}${after !== undefined ? ` w:after="${Math.round(after * 20)}"` : ''}${line !== undefined ? ` w:line="${Math.round(line * 240)}" w:lineRule="auto"` : ''}/>`)
  const indentAttributes = [
    ['left', pointValue(style.indentStart)],
    ['right', pointValue(style.indentEnd)],
    ['firstLine', pointValue(style.indentFirstLine)],
    ['hanging', pointValue(style.hanging)],
  ].filter((item): item is [string, number] => item[1] !== undefined)
  if (indentAttributes.length > 0)
    values.push(`<w:ind ${indentAttributes.map(([name, value]) => `w:${name}="${Math.round(value * 20)}"`).join(' ')}/>`)
  if (style.keepNext === BooleanNumber.TRUE)
    values.push('<w:keepNext/>')
  if (style.keepLines === BooleanNumber.TRUE)
    values.push('<w:keepLines/>')
  return values.length > 0 ? `<w:pPr>${values.join('')}</w:pPr>` : ''
}

function writeSectionProperties(documentData: IDocumentData): string {
  const style = documentData.documentStyle
  const width = Math.round((style.pageSize?.width ?? A4_WIDTH_TWIPS / 15) * 15)
  const height = Math.round((style.pageSize?.height ?? A4_HEIGHT_TWIPS / 15) * 15)
  const margin = (value: number | undefined) => Math.round((value ?? DEFAULT_MARGIN_TWIPS / 15) * 15)
  return `<w:sectPr><w:pgSz w:w="${width}" w:h="${height}"/><w:pgMar w:top="${margin(style.marginTop)}" w:right="${margin(style.marginRight)}" w:bottom="${margin(style.marginBottom)}" w:left="${margin(style.marginLeft)}" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`
}

function createDocxPackage(title: string): JSZip {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`)
  zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`)
  zip.file('word/styles.xml', defaultStylesXml())
  zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${escapeXml(title)}</dc:title><dc:creator>univer-plus</dc:creator></cp:coreProperties>`)
  zip.file('docProps/app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>univer-plus</Application></Properties>`)
  return zip
}

function defaultStylesXml(): string {
  const headings = [1, 2, 3, 4, 5].map(level => `<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:rPr><w:b/><w:sz w:val="${Math.max(24, 36 - level * 2)}"/></w:rPr></w:style>`).join('')
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${WORD_NS}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:qFormat/><w:rPr><w:b/><w:sz w:val="52"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:qFormat/><w:rPr><w:sz w:val="30"/></w:rPr></w:style>${headings}</w:styles>`
}

function namedStyleType(value?: string): NamedStyleType | undefined {
  const normalized = value?.replace(/[\s_-]/g, '').toLowerCase()
  const values: Record<string, NamedStyleType> = {
    normal: NamedStyleType.NORMAL_TEXT,
    title: NamedStyleType.TITLE,
    subtitle: NamedStyleType.SUBTITLE,
    heading1: NamedStyleType.HEADING_1,
    heading2: NamedStyleType.HEADING_2,
    heading3: NamedStyleType.HEADING_3,
    heading4: NamedStyleType.HEADING_4,
    heading5: NamedStyleType.HEADING_5,
  }
  return normalized ? values[normalized] : undefined
}

function namedStyleName(value?: NamedStyleType): string | undefined {
  const names: Partial<Record<NamedStyleType, string>> = {
    [NamedStyleType.NORMAL_TEXT]: 'Normal',
    [NamedStyleType.TITLE]: 'Title',
    [NamedStyleType.SUBTITLE]: 'Subtitle',
    [NamedStyleType.HEADING_1]: 'Heading1',
    [NamedStyleType.HEADING_2]: 'Heading2',
    [NamedStyleType.HEADING_3]: 'Heading3',
    [NamedStyleType.HEADING_4]: 'Heading4',
    [NamedStyleType.HEADING_5]: 'Heading5',
  }
  return value === undefined ? undefined : names[value]
}

function descendants(element: XmlElement, localName: string): XmlElement[] {
  const namespaced = element.getElementsByTagNameNS(WORD_NS, localName)
  if (namespaced.length > 0)
    return Array.from({ length: namespaced.length }, (_, index) => namespaced.item(index)).filter((item): item is XmlElement => Boolean(item))
  const all = element.getElementsByTagName('*')
  return Array.from({ length: all.length }, (_, index) => all.item(index)).filter((item): item is XmlElement => item?.localName === localName)
}

function childElements(element: XmlElement): XmlElement[] {
  const children: XmlElement[] = []
  for (let node = element.firstChild; node; node = node.nextSibling) {
    if (node.nodeType === 1)
      children.push(node as XmlElement)
  }
  return children
}

function firstElement(element: XmlElement | null | undefined, localName: string): XmlElement | undefined {
  if (!element)
    return undefined
  return descendants(element, localName)[0]
}

function attribute(element: XmlElement | null | undefined, localName: string): string | undefined {
  if (!element)
    return undefined
  return element.getAttributeNS(WORD_NS, localName) ?? element.getAttribute(`w:${localName}`) ?? element.getAttribute(localName) ?? undefined
}

function numericAttribute(element: XmlElement | null | undefined, localName: string): number | undefined {
  const value = attribute(element, localName)
  if (value === undefined)
    return undefined
  const number = Number(value)
  return Number.isFinite(number) ? number : undefined
}

function enabled(element: XmlElement | undefined): boolean {
  if (!element)
    return false
  const value = attribute(element, 'val')
  return value === undefined || !['0', 'false', 'off', 'none'].includes(value.toLowerCase())
}

function normalizeColor(value?: string): string | undefined {
  if (!value || value === 'auto' || !/^[0-9a-f]{6}$/i.test(value))
    return undefined
  return `#${value.toUpperCase()}`
}

function twipsToPixels(value?: number): number | undefined {
  return value === undefined ? undefined : value / 15
}

function pointValue(value?: { v: number }): number | undefined {
  return value?.v
}

function escapeXml(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
}

function escapeAttribute(value: string): string {
  return escapeXml(value).replaceAll('"', '&quot;').replaceAll('\'', '&apos;')
}
