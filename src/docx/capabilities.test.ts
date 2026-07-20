import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { inspectDocx, validateDocxBuffer } from './capabilities'
import { createEmptyDocx } from './converter'

describe('docx capability inspection', () => {
  it('allows a basic document', async () => {
    const buffer = await createEmptyDocx('Basic')
    await expect(validateDocxBuffer(buffer)).resolves.toBeUndefined()
    await expect(inspectDocx(buffer)).resolves.toEqual({ readOnly: false, reasons: [], paragraphCount: 1 })
  })

  it('protects document resources that cannot be written safely', async () => {
    const zip = await JSZip.loadAsync(await createEmptyDocx('Protected'))
    zip.file('word/media/image1.png', new Uint8Array([1, 2, 3]))
    zip.file('word/comments.xml', '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
    const inspection = await inspectDocx(await zip.generateAsync({ type: 'arraybuffer' }))
    expect(inspection.readOnly).toBe(true)
    expect(inspection.reasons).toEqual(['comments', 'images'])
  })

  it('protects advanced document markup', async () => {
    const zip = await JSZip.loadAsync(await createEmptyDocx('Protected'))
    zip.file('word/document.xml', `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl/><w:p><w:hyperlink><w:r><w:t>Link</w:t></w:r></w:hyperlink></w:p><w:sectPr/></w:body></w:document>`)
    const inspection = await inspectDocx(await zip.generateAsync({ type: 'arraybuffer' }))
    expect(inspection.reasons).toEqual(['hyperlinks', 'tables'])
  })

  it('protects custom styles and manual line breaks', async () => {
    const zip = await JSZip.loadAsync(await createEmptyDocx('Styles'))
    zip.file('word/document.xml', `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="CompanyQuote"/></w:pPr><w:r><w:t>First</w:t><w:br/><w:t>Second</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`)
    expect((await inspectDocx(await zip.generateAsync({ type: 'arraybuffer' }))).reasons).toEqual([
      'custom paragraph styles',
      'manual line or column breaks',
    ])
  })

  it('protects direct formatting that the DOCX exporter cannot reproduce', async () => {
    const zip = await JSZip.loadAsync(await createEmptyDocx('Formatting'))
    zip.file('word/document.xml', `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:shd w:fill="FFFF00"/></w:pPr><w:r><w:rPr><w:lang w:val="zh-CN"/></w:rPr><w:t>Styled</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`)

    expect((await inspectDocx(await zip.generateAsync({ type: 'arraybuffer' }))).reasons).toEqual([
      'formatting that cannot be written safely',
    ])
  })

  it('rejects malformed packages', async () => {
    const zip = new JSZip()
    zip.file('readme.txt', 'not a document')
    await expect(validateDocxBuffer(await zip.generateAsync({ type: 'arraybuffer' }))).rejects.toThrow('required DOCX metadata')
  })
})
