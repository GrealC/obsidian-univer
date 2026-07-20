import type { IDisposable, IDocumentData } from '@univerjs/core'
import type { IconName, TFile, WorkspaceLeaf } from 'obsidian'
import type { UniverPluginSettings } from '@/types/setting'
import type { UniverRuntime } from '@/univer/create'
import { CommandType } from '@univerjs/core'
import { FileView, normalizePath, Notice, setIcon } from 'obsidian'
import { inspectDocx } from '@/docx/capabilities'
import { exportDocx, importDocx } from '@/docx/converter'
import { renderDocxPreview } from '@/docx/preview'
import { assertSafeDocumentTransition } from '@/docx/safety'
import { createBinaryBackup, pruneBinaryBackups } from '@/services/backup'
import { SaveCoordinator } from '@/services/saveCoordinator'
import { docInit } from '@/univer/docs'
import { observeTheme } from '@/univer/theme'

export const Type = 'univer-docx'
const AUTO_SAVE_DELAY = 2000
type DocumentStatus = 'loading' | 'dirty' | 'saving' | 'saved' | 'protected' | 'error'

const STATUS_ICONS: Record<DocumentStatus, IconName> = {
  loading: 'loader-circle',
  dirty: 'circle',
  saving: 'loader-circle',
  saved: 'check',
  protected: 'shield-alert',
  error: 'circle-alert',
}

export class DocxTypeView extends FileView {
  private readonly saveCoordinator = new SaveCoordinator()
  private runtime?: UniverRuntime
  private commandDisposable?: IDisposable
  private themeObserver?: MutationObserver
  private autosaveTimer?: number
  private statusActionEl?: HTMLElement
  private statusLiveEl?: HTMLElement
  private statusState?: DocumentStatus
  private baseline = ''
  private lastSavedSnapshot?: IDocumentData
  private pendingDestructiveSnapshot?: string
  private fileDeleted = false
  private readOnly = false
  private backupCreated = false
  private lastKnownBuffer?: ArrayBuffer
  private protectedReasons: string[] = []
  private lastReadOnlyNoticeAt = 0
  private creatingEditableCopy = false

  constructor(leaf: WorkspaceLeaf, private readonly settings: UniverPluginSettings) {
    super(leaf)
  }

  getViewType(): string {
    return Type
  }

  async onOpen(): Promise<void> {
    this.ensureStatusAction()
    this.registerEvent(this.app.vault.on('delete', (file) => {
      if (file === this.file)
        this.fileDeleted = true
    }))
    this.registerDomEvent(window, 'keydown', (event) => {
      const isSave = (event.ctrlKey || event.metaKey) && !event.altKey && event.key.toLowerCase() === 's'
      if (!isSave || this.app.workspace.getActiveViewOfType(DocxTypeView) !== this)
        return
      event.preventDefault()
      void this.saveNow('manual').catch(() => undefined)
    })
  }

  async onLoadFile(file: TFile): Promise<void> {
    await super.onLoadFile(file)
    this.resetState()
    this.contentEl.empty()
    this.contentEl.addClass('univer-view', 'univer-office-view', 'univer-docx-view')

    if (!this.settings.isSupportDocx) {
      this.contentEl.createDiv({ cls: 'univer-empty-state', text: 'DOCX support is disabled. Re-enable it in Univer settings and restart Obsidian.' })
      return
    }

    this.ensureStatusAction()
    this.statusLiveEl = this.contentEl.createDiv({ cls: 'univer-save-status-live' })
    this.statusLiveEl.setAttrs({ 'role': 'status', 'aria-live': 'polite', 'aria-atomic': 'true' })
    this.setStatus('loading')
    const container = this.contentEl.createDiv({ cls: 'univer-editor-container' })

    try {
      const raw = await this.app.vault.readBinary(file)
      this.lastKnownBuffer = raw
      const inspection = await inspectDocx(raw)
      this.readOnly = inspection.readOnly
      this.protectedReasons = inspection.reasons

      if (this.readOnly) {
        this.renderProtectedHeader(container)
        try {
          container.addClass('univer-docx-preview')
          await renderDocxPreview(raw, container)
        }
        catch (previewError) {
          container.empty()
          container.removeClass('univer-docx-preview')
          const documentData = await importDocx(raw, file.basename)
          this.createProtectedUniverFallback(container, documentData)
          new Notice(`High-fidelity Word preview failed; showing the text fallback: ${errorMessage(previewError)}`, 8000)
        }
        this.setStatus('protected')
        return
      }

      const documentData = await importDocx(raw, file.basename)
      this.runtime = docInit(container, this.settings)
      this.runtime.univerAPI.createUniverDoc(documentData)
      this.themeObserver = observeTheme(this.runtime.univerAPI)
      const initialSnapshot = this.runtime.univerAPI.getActiveDocument()?.getSnapshot()
      this.baseline = initialSnapshot ? JSON.stringify(initialSnapshot) : ''
      this.lastSavedSnapshot = this.baseline ? cloneSnapshot(this.baseline) : undefined
      this.commandDisposable = this.runtime.univerAPI.addEvent(this.runtime.univerAPI.Event.CommandExecuted, (event) => {
        if (event.type === CommandType.MUTATION)
          this.scheduleSave()
      })
      this.setStatus('saved')
    }
    catch (error) {
      this.disposeEditor()
      container.createDiv({ cls: 'univer-error', text: `Cannot open this DOCX document: ${errorMessage(error)}` })
      this.setStatus('error')
      new Notice(`Cannot open ${file.name}: ${errorMessage(error)}`)
    }
  }

  async onUnloadFile(file: TFile): Promise<void> {
    try {
      if (!this.fileDeleted)
        await this.saveNow('close')
    }
    catch {
      // saveNow reports the error and leaves the original file untouched.
    }
    finally {
      this.disposeEditor()
      await super.onUnloadFile(file)
    }
  }

  async onClose(): Promise<void> {
    if (this.autosaveTimer !== undefined)
      window.clearTimeout(this.autosaveTimer)
    await this.saveCoordinator.flush()
    this.disposeEditor()
    this.saveCoordinator.dispose()
  }

  async saveNow(reason: 'auto' | 'manual' | 'close'): Promise<boolean> {
    if (!this.runtime || !this.file || this.readOnly || this.fileDeleted) {
      if (reason === 'manual' && this.readOnly)
        this.showReadOnlyNotice()
      return false
    }

    return this.saveCoordinator.run(async () => {
      const file = this.file
      if (!file || this.fileDeleted)
        return false
      try {
        const snapshot = this.runtime?.univerAPI.getActiveDocument()?.getSnapshot()
        if (!snapshot)
          return false
        const serialized = JSON.stringify(snapshot)
        if (serialized === this.baseline) {
          this.pendingDestructiveSnapshot = undefined
          this.setStatus('saved')
          return false
        }
        this.assertSafeTransition(snapshot, serialized, reason)

        const current = await this.app.vault.readBinary(file)
        if (!this.lastKnownBuffer || !buffersEqual(current, this.lastKnownBuffer))
          throw new Error('The file changed outside this editor. Reopen it before saving to avoid overwriting external changes.')

        this.setStatus('saving')
        const output = await exportDocx(snapshot, current)
        const latest = await this.app.vault.readBinary(file)
        if (!buffersEqual(current, latest))
          throw new Error('The file changed while it was being exported. Reopen it before saving.')
        if (this.settings.createBackups && !this.backupCreated) {
          await createBinaryBackup(this.app, file, latest)
          this.backupCreated = true
        }
        await this.app.vault.modifyBinary(file, output)
        await pruneBinaryBackups(this.app, file).catch(() => undefined)
        this.lastKnownBuffer = output
        this.baseline = serialized
        this.lastSavedSnapshot = cloneSnapshot(serialized)
        this.pendingDestructiveSnapshot = undefined
        this.setStatus('saved')
        if (reason === 'manual')
          new Notice(`${file.name} saved`)
        return true
      }
      catch (error) {
        this.setStatus('error')
        new Notice(`Cannot save ${file.name}: ${errorMessage(error)}`, 8000)
        throw error
      }
    })
  }

  private scheduleSave(): void {
    if (this.autosaveTimer !== undefined)
      window.clearTimeout(this.autosaveTimer)
    this.setStatus('dirty')
    this.autosaveTimer = window.setTimeout(() => {
      this.autosaveTimer = undefined
      void this.saveNow('auto').catch(() => undefined)
    }, AUTO_SAVE_DELAY)
  }

  private assertSafeTransition(snapshot: IDocumentData, serialized: string, reason: 'auto' | 'manual' | 'close'): void {
    if (!this.lastSavedSnapshot) {
      this.pendingDestructiveSnapshot = undefined
      return
    }
    try {
      assertSafeDocumentTransition(this.lastSavedSnapshot, snapshot)
      this.pendingDestructiveSnapshot = undefined
    }
    catch (error) {
      if (reason === 'manual' && this.pendingDestructiveSnapshot === serialized) {
        this.pendingDestructiveSnapshot = undefined
        return
      }
      this.pendingDestructiveSnapshot = serialized
      throw new Error(`${errorMessage(error)}. Press Ctrl/Cmd+S again to confirm that you intended to clear the document.`)
    }
  }

  private ensureStatusAction(): void {
    if (this.statusActionEl)
      return
    this.statusActionEl = this.addAction('loader-circle', this.statusLabel('loading'), () => this.handleStatusAction())
    this.statusActionEl.addClass('univer-save-status')
  }

  private setStatus(status: DocumentStatus): void {
    this.ensureStatusAction()
    const label = this.statusLabel(status)
    const changed = this.statusState !== status
    this.statusState = status
    if (this.statusActionEl) {
      setIcon(this.statusActionEl, STATUS_ICONS[status])
      this.statusActionEl.dataset.state = status
      this.statusActionEl.setAttrs({ 'title': label, 'aria-label': label })
      const actionable = status === 'dirty' || status === 'protected' || (status === 'error' && Boolean(this.runtime) && !this.readOnly)
      this.statusActionEl.setAttribute('aria-disabled', String(!actionable))
      this.statusActionEl.tabIndex = actionable ? 0 : -1
    }
    if (changed || !this.statusLiveEl?.textContent)
      this.statusLiveEl?.setText(label)
  }

  private statusLabel(status: DocumentStatus): string {
    const zh = this.settings.language === 'ZH' || this.settings.language === 'TW'
    const labels = zh
      ? {
          loading: '正在打开 Word 文档',
          dirty: '有未保存的更改，点击立即保存',
          saving: '正在保存 Word 文档',
          saved: 'Word 文档已保存',
          protected: '保护视图，点击查看原因',
          error: 'Word 文档操作失败',
        }
      : {
          loading: 'Opening Word document',
          dirty: 'Unsaved changes; click to save now',
          saving: 'Saving Word document',
          saved: 'Word document saved',
          protected: 'Protected view; click for details',
          error: 'Word document operation failed',
        }
    return labels[status]
  }

  private handleStatusAction(): void {
    if (this.statusState === 'protected') {
      this.showReadOnlyNotice()
      return
    }
    if (this.statusState === 'dirty' || (this.statusState === 'error' && this.runtime && !this.readOnly))
      void this.saveNow('manual').catch(() => undefined)
  }

  private showReadOnlyNotice(): void {
    const now = Date.now()
    if (now - this.lastReadOnlyNoticeAt < 2000)
      return
    this.lastReadOnlyNoticeAt = now
    const reasons = this.protectedReasons.length > 0 ? `: ${this.protectedReasons.join(', ')}` : ''
    new Notice(`This DOCX file is protected because it contains Word features that cannot be saved safely${reasons}.`, 8000)
  }

  private renderProtectedHeader(container: HTMLElement): void {
    const zh = this.settings.language === 'ZH' || this.settings.language === 'TW'
    const bar = this.contentEl.createDiv({ cls: 'univer-protected-bar' })
    const icon = bar.createSpan({ cls: 'univer-protected-bar__icon' })
    setIcon(icon, 'shield-check')
    const copy = bar.createDiv({ cls: 'univer-protected-bar__copy' })
    copy.createEl('strong', { text: zh ? '高保真只读预览' : 'High-fidelity read-only preview' })
    copy.createSpan({ text: zh ? '复杂 Word 版式已保留，原文件不会被修改' : 'Complex Word layout is preserved; the original file will not be modified' })
    const editCopy = bar.createEl('button', {
      cls: 'univer-protected-bar__edit',
      attr: {
        type: 'button',
        title: zh ? '创建可编辑 DOCX 副本' : 'Create an editable DOCX copy',
      },
    })
    const editIcon = editCopy.createSpan()
    setIcon(editIcon, 'file-pen-line')
    editCopy.createSpan({ text: zh ? '编辑副本' : 'Edit copy' })
    editCopy.addEventListener('click', () => {
      void this.createEditableCopy(editCopy).catch(() => undefined)
    })
    const details = bar.createEl('button', {
      cls: 'clickable-icon univer-protected-bar__details',
      attr: {
        'aria-label': zh ? '查看保护视图原因' : 'View protected-view details',
        'title': zh ? '查看保护视图原因' : 'View protected-view details',
        'type': 'button',
      },
    })
    setIcon(details, 'info')
    details.addEventListener('click', () => this.showReadOnlyNotice())
    container.before(bar)
  }

  private async createEditableCopy(button: HTMLButtonElement): Promise<void> {
    if (this.creatingEditableCopy || !this.file)
      return
    this.creatingEditableCopy = true
    button.disabled = true
    try {
      const sourceFile = this.file
      const source = await this.app.vault.readBinary(sourceFile)
      const documentData = await importDocx(source, `${sourceFile.basename}-editable`)
      const output = await exportDocx(documentData)
      const path = await this.nextEditableCopyPath(sourceFile)
      await this.app.vault.createBinary(path, output)
      await this.app.workspace.getLeaf(true).setViewState({
        type: Type,
        active: true,
        state: { file: path },
      })
      new Notice(`Created editable copy: ${path}`)
    }
    catch (error) {
      new Notice(`Cannot create an editable DOCX copy: ${errorMessage(error)}`, 8000)
      throw error
    }
    finally {
      this.creatingEditableCopy = false
      button.disabled = false
    }
  }

  private async nextEditableCopyPath(sourceFile: TFile): Promise<string> {
    const directory = sourceFile.parent?.path
    for (let index = 0; index < 10000; index++) {
      const suffix = index === 0 ? '' : `-${index}`
      const name = `${sourceFile.basename}-editable${suffix}.docx`
      const path = normalizePath(directory ? `${directory}/${name}` : name)
      if (!await this.app.vault.adapter.exists(path))
        return path
    }
    throw new Error('Too many editable copies already exist')
  }

  private createProtectedUniverFallback(container: HTMLElement, documentData: IDocumentData): void {
    container.addClass('univer-editor-readonly')
    container.setAttr('aria-readonly', 'true')
    this.runtime = docInit(container, this.settings)
    this.runtime.univerAPI.createUniverDoc(documentData)
    this.themeObserver = observeTheme(this.runtime.univerAPI)
    this.commandDisposable = this.runtime.univerAPI.addEvent(this.runtime.univerAPI.Event.BeforeCommandExecute, (event) => {
      if (event.type !== CommandType.MUTATION)
        return
      event.cancel = true
      this.showReadOnlyNotice()
    })
  }

  private resetState(): void {
    this.disposeEditor()
    this.statusLiveEl = undefined
    this.statusState = undefined
    this.baseline = ''
    this.lastSavedSnapshot = undefined
    this.pendingDestructiveSnapshot = undefined
    this.fileDeleted = false
    this.readOnly = false
    this.backupCreated = false
    this.lastKnownBuffer = undefined
    this.protectedReasons = []
    this.lastReadOnlyNoticeAt = 0
    this.creatingEditableCopy = false
  }

  private disposeEditor(): void {
    if (this.autosaveTimer !== undefined) {
      window.clearTimeout(this.autosaveTimer)
      this.autosaveTimer = undefined
    }
    this.commandDisposable?.dispose()
    this.commandDisposable = undefined
    this.themeObserver?.disconnect()
    this.themeObserver = undefined
    this.runtime?.univer.dispose()
    this.runtime = undefined
  }
}

function buffersEqual(left: ArrayBuffer, right: ArrayBuffer): boolean {
  if (left.byteLength !== right.byteLength)
    return false
  const leftBytes = new Uint8Array(left)
  const rightBytes = new Uint8Array(right)
  return leftBytes.every((value, index) => value === rightBytes[index])
}

function cloneSnapshot(serialized: string): IDocumentData {
  return JSON.parse(serialized) as IDocumentData
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}
