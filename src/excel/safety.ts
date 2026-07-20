import type { ICellData, IWorkbookData } from '@univerjs/core'

export interface WorkbookMetrics {
  sheetCount: number
  populatedCellCount: number
  preservedTableCount: number
}

export function workbookMetrics(workbook: IWorkbookData): WorkbookMetrics {
  let populatedCellCount = 0
  let preservedTableCount = 0
  for (const sheetId of workbook.sheetOrder) {
    const sheet = workbook.sheets[sheetId]
    const metadata = sheet?.custom?.excel as { tables?: unknown[] } | undefined
    preservedTableCount += metadata?.tables?.length ?? 0
    for (const row of Object.values(sheet?.cellData ?? {})) {
      const cells = Object.values(row ?? {}) as Array<ICellData | null | undefined>
      for (const cell of cells) {
        if (cell && (cell.v !== undefined || cell.f || cell.p))
          populatedCellCount++
      }
    }
  }
  return { sheetCount: workbook.sheetOrder.length, populatedCellCount, preservedTableCount }
}

export function assertSafeWorkbookTransition(previous: IWorkbookData, next: IWorkbookData): void {
  const before = workbookMetrics(previous)
  const after = workbookMetrics(next)
  if (before.sheetCount > 0 && after.sheetCount === 0)
    throw new Error('Save refused because every worksheet disappeared from the workbook')
  if (before.populatedCellCount > 0 && after.populatedCellCount === 0)
    throw new Error('Save refused because every populated cell disappeared from the workbook')
  if (after.preservedTableCount < before.preservedTableCount)
    throw new Error('Save refused because a preserved Excel table disappeared from the workbook')
}
