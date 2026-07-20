import type { Options } from 'docx-preview'

export const DOCX_PREVIEW_CLASS = 'univer-plus-docx'

export const DOCX_PREVIEW_OPTIONS: Partial<Options> = {
  className: DOCX_PREVIEW_CLASS,
  inWrapper: true,
  hideWrapperOnPrint: true,
  ignoreWidth: false,
  ignoreHeight: false,
  ignoreFonts: false,
  breakPages: true,
  ignoreLastRenderedPageBreak: false,
  experimental: true,
  trimXmlDeclaration: true,
  useBase64URL: true,
  renderChanges: false,
  renderHeaders: true,
  renderFooters: true,
  renderFootnotes: true,
  renderEndnotes: true,
  renderComments: false,
  renderAltChunks: false,
  debug: false,
}

export async function renderDocxPreview(buffer: ArrayBuffer, container: HTMLElement): Promise<void> {
  const { renderAsync } = await import('docx-preview')
  await renderAsync(buffer, container, container, DOCX_PREVIEW_OPTIONS)
}
