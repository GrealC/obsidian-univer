import type { IDisposable, IWorkbookData } from '@univerjs/core'
import type { WorkspaceLeaf } from 'obsidian'
import type { UniverPluginSettings } from '@/types/setting'
import type { UniverRuntime } from '@/univer/create'
import { Notice, TextFileView } from 'obsidian'
import { sheetInit } from '@/univer/sheets'
import { observeTheme } from '@/univer/theme'
import { getLanguage } from '@/utils/common'

export const Type = 'univer-sheet'

export class USheetView extends TextFileView {
  private runtime?: UniverRuntime
  private commandDisposable?: IDisposable
  private themeObserver?: MutationObserver
  private sourceData = ''

  constructor(leaf: WorkspaceLeaf, private readonly settings: UniverPluginSettings) {
    super(leaf)
  }

  getViewData(): string {
    const workbook = this.runtime?.univerAPI.getActiveWorkbook()
    return workbook ? JSON.stringify(workbook.save()) : this.sourceData
  }

  setViewData(data: string): void {
    this.disposeEditor()
    this.sourceData = data
    this.contentEl.empty()
    this.contentEl.addClass('univer-view')
    const container = this.contentEl.createDiv({ cls: 'univer-editor-container' })

    let workbookData: Partial<IWorkbookData> = { name: this.file?.basename ?? 'Univer Sheet' }
    if (data) {
      try {
        workbookData = JSON.parse(data) as IWorkbookData
      }
      catch (error) {
        container.createDiv({ cls: 'univer-error', text: 'This Univer Sheet is not valid JSON. The original file has not been changed.' })
        new Notice(`Cannot open ${this.file?.name ?? 'Univer Sheet'}: ${errorMessage(error)}`)
        return
      }
    }

    this.runtime = sheetInit(container, this.settings)
    this.runtime.univerAPI.createWorkbook(workbookData)
    this.themeObserver = observeTheme(this.runtime.univerAPI)
    this.commandDisposable = this.runtime.univerAPI.addEvent(this.runtime.univerAPI.Event.CommandExecuted, () => {
      this.requestSave()
    })
  }

  getViewType(): string {
    return Type
  }

  setLanguage(): void {
    this.runtime?.univerAPI.setLocale(getLanguage(this.settings))
  }

  clear(): void {
    this.disposeEditor()
    this.contentEl.empty()
  }

  async onClose(): Promise<void> {
    if (this.runtime && this.file)
      await this.save()
    this.disposeEditor()
  }

  private disposeEditor(): void {
    this.commandDisposable?.dispose()
    this.commandDisposable = undefined
    this.themeObserver?.disconnect()
    this.themeObserver = undefined
    this.runtime?.univer.dispose()
    this.runtime = undefined
  }
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}
