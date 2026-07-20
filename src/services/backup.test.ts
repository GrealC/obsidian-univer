import { describe, expect, it } from 'vitest'
import { selectObsoleteBackups } from './backupRetention'

describe('backup retention', () => {
  it('keeps only the newest backups for the same source file', () => {
    const files = [
      'plugins/univer-plus/backups/2026-07-18T08-00-00-000Z-notes_file.docx',
      'plugins/univer-plus/backups/2026-07-19T08-00-00-000Z-notes_file.docx',
      'plugins/univer-plus/backups/2026-07-20T08-00-00-000Z-notes_file.docx',
      'plugins/univer-plus/backups/2026-07-21T08-00-00-000Z-notes_file.docx',
      'plugins/univer-plus/backups/2026-07-17T08-00-00-000Z-other.docx',
    ]

    expect(selectObsoleteBackups(files, 'notes_file.docx', 3)).toEqual([
      'plugins/univer-plus/backups/2026-07-18T08-00-00-000Z-notes_file.docx',
    ])
  })

  it('can remove every matching backup when retention is zero', () => {
    expect(selectObsoleteBackups(['backups/2-file.xlsx', 'backups/1-file.xlsx'], 'file.xlsx', 0)).toEqual([
      'backups/2-file.xlsx',
      'backups/1-file.xlsx',
    ])
  })
})
