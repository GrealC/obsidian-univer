import type { App } from 'obsidian'
import { normalizePath, Notice } from 'obsidian'
import { createEmptyDocx } from '@/docx/converter'
import { createEmptyXlsx } from '@/excel/converter'
import { Type as DocxType } from '@/views/docx'
import { Type as DocType } from '@/views/udoc'
import { Type as SheetType } from '@/views/usheet'
import { Type as XlsxType } from '@/views/xlsx'

export async function createNewFile(app: App, suffix: string, folderPath?: string): Promise<void> {
  try {
    if (folderPath && !app.vault.getAbstractFileByPath(folderPath))
      await app.vault.createFolder(folderPath)

    const filePath = await nextAvailablePath(app, suffix, folderPath)
    if (suffix === 'xlsx')
      await app.vault.createBinary(filePath, await createEmptyXlsx())
    else if (suffix === 'docx')
      await app.vault.createBinary(filePath, await createEmptyDocx(filePath.replace(/^.*\//, '').replace(/\.docx$/i, '')))
    else
      await app.vault.create(filePath, '')

    await app.workspace.getLeaf(true).setViewState({
      type: getFileType(suffix),
      active: true,
      state: { file: filePath },
    })
    new Notice(`Created ${filePath}`)
  }
  catch (error) {
    new Notice(`Cannot create ${suffix} file: ${errorMessage(error)}`, 8000)
  }
}

async function nextAvailablePath(app: App, suffix: string, folderPath?: string): Promise<string> {
  for (let index = 0; index < 10000; index++) {
    const name = `Untitled${index === 0 ? '' : `-${index}`}.${suffix}`
    const path = normalizePath(folderPath ? `${folderPath}/${name}` : name)
    if (!await app.vault.adapter.exists(path))
      return path
  }
  throw new Error('Too many Untitled files already exist')
}

function getFileType(suffix: string): string {
  if (suffix === 'udoc')
    return DocType
  if (suffix === 'xlsx')
    return XlsxType
  if (suffix === 'docx')
    return DocxType
  return SheetType
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}
