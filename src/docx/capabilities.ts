import JSZip from 'jszip'

export interface DocxInspection {
  readOnly: boolean
  reasons: string[]
  paragraphCount: number
}

const unsupportedParts: Array<[RegExp, string]> = [
  [/^word\/media\//i, 'images'],
  [/^word\/charts\//i, 'charts'],
  [/^word\/diagrams\//i, 'diagrams'],
  [/^word\/embeddings\//i, 'embedded objects'],
  [/^word\/(comments|commentsExtended|people)\.xml$/i, 'comments'],
  [/^word\/(footnotes|endnotes)\.xml$/i, 'footnotes or endnotes'],
  [/^word\/(header|footer)\d+\.xml$/i, 'headers or footers'],
  [/^word\/vbaProject\.bin$/i, 'VBA macros'],
  [/^word\/(glossary|webSettings)\//i, 'advanced Word resources'],
]

const unsupportedDocumentMarkup: Array<[RegExp, string]> = [
  [/<w:tbl\b/i, 'tables'],
  [/<w:(?:drawing|pict|object)\b/i, 'drawings or text boxes'],
  [/<w:(?:ins|del|moveFrom|moveTo)\b/i, 'tracked changes'],
  [/<w:(?:commentRangeStart|commentReference)\b/i, 'comments'],
  [/<w:(?:footnoteReference|endnoteReference)\b/i, 'footnotes or endnotes'],
  [/<w:(?:fldChar|instrText|fldSimple)\b/i, 'fields'],
  [/<w:sdt\b/i, 'content controls'],
  [/<w:altChunk\b/i, 'embedded HTML'],
  [/<w:numPr\b/i, 'numbered or bulleted lists'],
  [/<w:hyperlink\b/i, 'hyperlinks'],
  [/<w:(?:bookmarkStart|bookmarkEnd)\b/i, 'bookmarks'],
  [/<w:(?:sym|cr)\b/i, 'special Word characters'],
  [/<w:(?:tabs|pBdr|pageBreakBefore|textDirection|bidi)\b/i, 'advanced paragraph formatting'],
  [/<w:(?:vanish|webHidden|outline|shadow|emboss|imprint|smallCaps|caps)\b/i, 'advanced text effects'],
  [/<w:(?:dstrike|bCs|iCs|rtl|cs|em|fitText|kern|position|spacing|w|eastAsianLayout)\b/i, 'advanced run formatting'],
  [/<w:(?:shd|lang|rPrChange|pPrChange)\b/i, 'formatting that cannot be written safely'],
  [/<w:spacing[^>]*w:lineRule\s*=\s*["'](?!(?:auto|atLeast)["'])/i, 'exact paragraph line spacing'],
  [/<w:sectPr(?:\s[^>]*)?>[\s\S]*?<w:(?:cols|pgBorders|lnNumType)\b/i, 'advanced section layout'],
]

export async function inspectDocx(buffer: ArrayBuffer): Promise<DocxInspection> {
  await validateDocxBuffer(buffer)
  const zip = await JSZip.loadAsync(buffer)
  const fileNames = Object.keys(zip.files)
  const reasons = new Set<string>()

  for (const fileName of fileNames) {
    for (const [pattern, reason] of unsupportedParts) {
      if (pattern.test(fileName))
        reasons.add(reason)
    }
  }

  const documentXml = await zip.file('word/document.xml')?.async('text') ?? ''
  for (const [pattern, reason] of unsupportedDocumentMarkup) {
    if (pattern.test(documentXml))
      reasons.add(reason)
  }

  inspectStyleReferences(documentXml, reasons)
  inspectBreaks(documentXml, reasons)

  if ((documentXml.match(/<w:sectPr\b/gi)?.length ?? 0) > 1)
    reasons.add('multiple document sections')

  const settingsXml = await zip.file('word/settings.xml')?.async('text')
  if (settingsXml && /<w:documentProtection\b/i.test(settingsXml))
    reasons.add('document protection')

  return {
    readOnly: reasons.size > 0,
    reasons: [...reasons].sort(),
    paragraphCount: documentXml.match(/<w:p(?:\s|>)/gi)?.length ?? 0,
  }
}

function inspectBreaks(documentXml: string, reasons: Set<string>): void {
  const breakTags = documentXml.match(/<w:br[^>]*>/gi) ?? []
  if (breakTags.some(tag => !/w:type\s*=\s*["']page["']/i.test(tag)))
    reasons.add('manual line or column breaks')
}

function inspectStyleReferences(documentXml: string, reasons: Set<string>): void {
  const supportedParagraphStyles = new Set(['normal', 'title', 'subtitle', 'heading1', 'heading2', 'heading3', 'heading4', 'heading5'])
  const paragraphStyleTags = documentXml.match(/<w:pStyle(?:\s[^>]*)?>/gi) ?? []
  for (const tag of paragraphStyleTags) {
    const value = /w:val\s*=\s*["']([^"']+)["']/i.exec(tag)?.[1]
    const style = value?.replace(/[\s_-]/g, '').toLowerCase()
    if (style && !supportedParagraphStyles.has(style))
      reasons.add('custom paragraph styles')
  }
  if (/<w:rStyle\b/i.test(documentXml))
    reasons.add('character styles')
}

export async function validateDocxBuffer(buffer: ArrayBuffer): Promise<void> {
  if (buffer.byteLength === 0)
    throw new Error('The DOCX file is empty')

  const zip = await JSZip.loadAsync(buffer)
  if (!zip.file('[Content_Types].xml') || !zip.file('_rels/.rels') || !zip.file('word/document.xml'))
    throw new Error('The file is missing required DOCX metadata')

  const documentXml = await zip.file('word/document.xml')?.async('text')
  if (!documentXml || !/<w:document\b/i.test(documentXml) || !/<w:body\b/i.test(documentXml))
    throw new Error('The DOCX document body is invalid')
}
