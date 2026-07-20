import type { IDocumentData } from '@univerjs/core'
import { describe, expect, it } from 'vitest'
import { assertSafeDocumentTransition, documentTextLength } from './safety'

function document(dataStream: string): IDocumentData {
  return { id: 'doc', documentStyle: {}, body: { dataStream } }
}

describe('docx document safety', () => {
  it('ignores document control characters when measuring content', () => {
    expect(documentTextLength(document('\r\n\f\v'))).toBe(0)
    expect(documentTextLength(document(' 内容 \r\n'))).toBe(2)
  })

  it('refuses an unexpected transition to an empty document', () => {
    expect(() => assertSafeDocumentTransition(document('Text\r\n'), document('\r\n'))).toThrow('all document text disappeared')
  })

  it('allows intentional edits that retain content', () => {
    expect(() => assertSafeDocumentTransition(document('Before\r\n'), document('After\r\n'))).not.toThrow()
  })
})
