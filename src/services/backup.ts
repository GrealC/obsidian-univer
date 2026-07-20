import type { App, TFile } from 'obsidian'
import { normalizePath } from 'obsidian'
import { selectObsoleteBackups } from './backupRetention'

const BACKUP_DIR = 'plugins/univer-plus/backups'
const BACKUPS_PER_FILE = 3

export async function createBinaryBackup(app: App, file: TFile, data: ArrayBuffer): Promise<string> {
  const adapter = app.vault.adapter
  const backupDir = normalizePath(`${app.vault.configDir}/${BACKUP_DIR}`)
  if (!await adapter.exists(backupDir))
    await adapter.mkdir(backupDir)

  const timestamp = new Date().toISOString().replaceAll(':', '-').replaceAll('.', '-')
  const safeName = safeBackupName(file.path)
  const backupPath = normalizePath(`${backupDir}/${timestamp}-${safeName}`)
  await adapter.writeBinary(backupPath, data)
  return backupPath
}

export const createXlsxBackup = createBinaryBackup

export async function pruneBinaryBackups(app: App, file: TFile, keep = BACKUPS_PER_FILE): Promise<void> {
  const adapter = app.vault.adapter
  const backupDir = normalizePath(`${app.vault.configDir}/${BACKUP_DIR}`)
  if (!await adapter.exists(backupDir))
    return

  const listing = await adapter.list(backupDir)
  const safeName = safeBackupName(file.path)
  const obsolete = selectObsoleteBackups(listing.files, safeName, keep)
  await Promise.all(obsolete.map(path => adapter.remove(path)))
}

function safeBackupName(filePath: string): string {
  return filePath.replace(/[\\/:*?"<>|]/g, '_')
}
