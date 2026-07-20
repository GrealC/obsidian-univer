import type { App, TFile } from 'obsidian'
import { normalizePath } from 'obsidian'

const BACKUP_DIR = 'plugins/univer-plus/backups'

export async function createBinaryBackup(app: App, file: TFile, data: ArrayBuffer): Promise<string> {
  const adapter = app.vault.adapter
  const backupDir = normalizePath(`${app.vault.configDir}/${BACKUP_DIR}`)
  if (!await adapter.exists(backupDir))
    await adapter.mkdir(backupDir)

  const timestamp = new Date().toISOString().replaceAll(':', '-').replaceAll('.', '-')
  const safeName = file.path.replace(/[\\/:*?"<>|]/g, '_')
  const backupPath = normalizePath(`${backupDir}/${timestamp}-${safeName}`)
  await adapter.writeBinary(backupPath, data)
  return backupPath
}

export const createXlsxBackup = createBinaryBackup
