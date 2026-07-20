import type { IDocumentData } from '@univerjs/core'

export function documentTextLength(document: IDocumentData): number {
  return (document.body?.dataStream ?? '')
    .replace(/[\r\n\v\f\0]/g, '')
    .trim()
    .length
}

export function assertSafeDocumentTransition(previous: IDocumentData, next: IDocumentData): void {
  if (documentTextLength(previous) > 0 && documentTextLength(next) === 0)
    throw new Error('Save refused because all document text disappeared')
}
