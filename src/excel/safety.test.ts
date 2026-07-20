import type { IWorkbookData } from '@univerjs/core'
import { LocaleType } from '@univerjs/core'
import { describe, expect, it } from 'vitest'
import { assertSafeWorkbookTransition, workbookMetrics } from './safety'

function workbook(cellValue?: string): IWorkbookData {
  return {
    id: 'workbook',
    name: 'Workbook',
    appVersion: '0.25.1',
    locale: LocaleType.EN_US,
    styles: {},
    sheetOrder: ['sheet'],
    sheets: {
      sheet: {
        id: 'sheet',
        name: 'Sheet1',
        cellData: cellValue === undefined ? {} : { 0: { 0: { v: cellValue } } },
      },
    },
  }
}

describe('catastrophic data loss guard', () => {
  it('counts populated cells', () => {
    expect(workbookMetrics(workbook('value'))).toEqual({ sheetCount: 1, populatedCellCount: 1, preservedTableCount: 0 })
  })

  it('rejects a transition that unexpectedly clears the workbook', () => {
    expect(() => assertSafeWorkbookTransition(workbook('value'), workbook())).toThrow('every populated cell disappeared')
  })

  it('allows normal edits', () => {
    expect(() => assertSafeWorkbookTransition(workbook('before'), workbook('after'))).not.toThrow()
  })
})
