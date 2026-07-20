import { BooleanNumber, HorizontalAlign } from '@univerjs/core'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { createEmptyDocx, exportDocx, importDocx } from './converter'

describe('local docx converter', () => {
  it('creates an editable empty DOCX package', async () => {
    const buffer = await createEmptyDocx('Notes')
    const zip = await JSZip.loadAsync(buffer)
    expect(zip.file('word/document.xml')).toBeTruthy()
    expect(zip.file('word/styles.xml')).toBeTruthy()
    expect(zip.file('docProps/core.xml')).toBeTruthy()
  })

  it('imports paragraphs and direct text formatting', async () => {
    const zip = await JSZip.loadAsync(await createEmptyDocx('Import'))
    zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
        <w:p><w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="C00000"/><w:sz w:val="28"/></w:rPr><w:t>Hello</w:t></w:r></w:p>
        <w:p><w:r><w:t xml:space="preserve">中文 文档</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
      </w:body></w:document>`)
    const data = await importDocx(await zip.generateAsync({ type: 'arraybuffer' }), 'Import')

    expect(data.body?.dataStream).toBe('Hello\r中文 文档\r\n')
    expect(data.body?.textRuns?.[0]).toMatchObject({ st: 0, ed: 5, ts: { bl: BooleanNumber.TRUE, fs: 14, cl: { rgb: '#C00000' } } })
    expect(data.body?.paragraphs?.[0].paragraphStyle?.horizontalAlign).toBe(HorizontalAlign.CENTER)
  })

  it('round-trips edited text through DOCX', async () => {
    const original = await createEmptyDocx('Round trip')
    const data = await importDocx(original, 'Round trip')
    data.body = {
      ...data.body,
      dataStream: 'First paragraph\rSecond paragraph\r\n',
      paragraphs: [{ startIndex: 15 }, { startIndex: 32 }],
      textRuns: [{ st: 0, ed: 5, ts: { bl: BooleanNumber.TRUE, ff: 'Microsoft YaHei' } }],
    }
    const output = await exportDocx(data, original)
    const imported = await importDocx(output, 'Round trip')
    expect(imported.body?.dataStream).toBe(data.body.dataStream)
    expect(imported.body?.textRuns?.[0]).toMatchObject({ st: 0, ed: 5, ts: { bl: BooleanNumber.TRUE, ff: 'Microsoft YaHei' } })
  })
})
