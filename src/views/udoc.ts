import type { IDisposable, IDocumentData } from '@univerjs/core'
import type { WorkspaceLeaf } from 'obsidian'
import type { UniverPluginSettings } from '@/types/setting'
import type { UniverRuntime } from '@/univer/create'
import { Notice, TextFileView } from 'obsidian'
import { docInit } from '@/univer/docs'
import { observeTheme } from '@/univer/theme'

export const Type = 'univer-doc'

export class UDocView extends TextFileView {
  private runtime?: UniverRuntime
  private commandDisposable?: IDisposable
  private themeObserver?: MutationObserver
  private sourceData = ''

  constructor(leaf: WorkspaceLeaf, private readonly settings: UniverPluginSettings) {
    super(leaf)
  }

  getViewData(): string {
    const document = this.runtime?.univerAPI.getActiveDocument()
    return document ? JSON.stringify(document.getSnapshot()) : this.sourceData
  }

  setViewData(data: string): void {
    this.disposeEditor()
    this.sourceData = data
    this.contentEl.empty()
    this.contentEl.addClass('univer-view')
    const container = this.contentEl.createDiv({ cls: 'univer-editor-container' })

    let documentData: Partial<IDocumentData> = {}
    if (data) {
      try {
        documentData = JSON.parse(data) as IDocumentData
      }
      catch (error) {
        container.createDiv({ cls: 'univer-error', text: 'This Univer Document is not valid JSON. The original file has not been changed.' })
        new Notice(`Cannot open ${this.file?.name ?? 'Univer Document'}: ${errorMessage(error)}`)
        return
      }
    }

    this.runtime = docInit(container, this.settings)
    this.runtime.univerAPI.createUniverDoc(documentData)
    this.themeObserver = observeTheme(this.runtime.univerAPI)
    this.commandDisposable = this.runtime.univerAPI.addEvent(this.runtime.univerAPI.Event.CommandExecuted, () => {
      this.requestSave()
    })
  }

  getViewType(): string {
    return Type
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
