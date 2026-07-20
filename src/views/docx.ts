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
import { uiText } from '@/i18n'
import { createBinaryBackup, pruneBinaryBackups } from '@/services/backup'
import { hasUnsavedSnapshot, isOfficeSaveShortcut } from '@/services/officeSave'
import { SaveCoordinator } from '@/services/saveCoordinator'
import { docInit } from '@/univer/docs'
import { observeTheme } from '@/univer/theme'
import { getLanguage } from '@/utils/common'

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
  private loadedFile?: TFile
  private fileDeleted = false
  private readOnly = false
  private readonly knownBuffers = new Map<string, ArrayBuffer>()
  private readonly backedUpFiles = new Set<string>()
  private lastKnownBuffer?: ArrayBuffer
  private protectedReasons: string[] = []
  private lastReadOnlyNoticeAt = 0
  private creatingEditableCopy = false
  private protectedTitleEl?: HTMLElement
  private protectedMessageEl?: HTMLElement
  private editCopyEl?: HTMLButtonElement
  private editCopyLabelEl?: HTMLElement
  private protectedDetailsEl?: HTMLButtonElement

  constructor(leaf: WorkspaceLeaf, private readonly settings: UniverPluginSettings) {
    super(leaf)
  }

  getViewType(): string {
    return Type
  }

  setLanguage(): void {
    this.runtime?.univerAPI.setLocale(getLanguage(this.settings))
    if (this.statusState) {
      const status = this.statusState
      this.statusState = undefined
      this.setStatus(status)
    }
    this.refreshProtectedHeader()
  }

  async onOpen(): Promise<void> {
    this.ensureStatusAction()
    this.registerEvent(this.app.vault.on('delete', (file) => {
      if (file.path === this.loadedFile?.path)
        this.fileDeleted = true
    }))
    this.registerDomEvent(window, 'keydown', (event) => {
      if (!isOfficeSaveShortcut(event) || this.app.workspace.activeLeaf?.view !== this)
        return
      event.preventDefault()
      void this.saveNow('manual').catch(() => undefined)
    }, { capture: true })
  }

  async onLoadFile(file: TFile): Promise<void> {
    const previousFile = this.loadedFile
    if (previousFile && previousFile.path !== file.path) {
      if (this.runtime && !this.fileDeleted)
        await this.saveNow('close', previousFile).catch(() => undefined)
      this.knownBuffers.delete(previousFile.path)
      this.backedUpFiles.delete(previousFile.path)
    }

    await super.onLoadFile(file)
    this.resetState()
    this.loadedFile = file
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
      this.knownBuffers.set(file.path, raw)
      this.backedUpFiles.delete(file.path)
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
      this.commandDisposable = this.runtime.univerAPI.addEvent(this.runtime.univerAPI.Event.CommandExecuted, () => this.scheduleSave())
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
    const unloadedRuntime = this.runtime
    try {
      if (!this.fileDeleted)
        await this.saveNow('close', file)
    }
    catch {
      // saveNow reports the error and leaves the original file untouched.
    }
    finally {
      if (this.loadedFile?.path === file.path && this.runtime === unloadedRuntime) {
        this.disposeEditor()
        this.loadedFile = undefined
        this.knownBuffers.delete(file.path)
        this.backedUpFiles.delete(file.path)
      }
      await super.onUnloadFile(file)
    }
  }

  async onClose(): Promise<void> {
    if (this.autosaveTimer !== undefined)
      window.clearTimeout(this.autosaveTimer)
    try {
      if (!this.fileDeleted)
        await this.saveNow('close')
    }
    catch {
      // saveNow already reports the error and leaves the original file untouched.
    }
    finally {
      await this.saveCoordinator.flush()
      this.disposeEditor()
      this.saveCoordinator.dispose()
    }
  }

  async saveNow(reason: 'auto' | 'manual' | 'close', targetFile = this.loadedFile): Promise<boolean> {
    const runtime = this.runtime
    if (!runtime || !targetFile || targetFile.path !== this.loadedFile?.path || this.readOnly || this.fileDeleted) {
      if (reason === 'manual' && this.readOnly)
        this.showReadOnlyNotice()
      return false
    }

    const snapshot = runtime.univerAPI.getActiveDocument()?.getSnapshot()
    if (!snapshot)
      return false
    const serialized = JSON.stringify(snapshot)
    if (serialized === this.baseline) {
      this.pendingDestructiveSnapshot = undefined
      this.setStatus('saved')
      return false
    }
    this.assertSafeTransition(snapshot, serialized, reason)

    const sessionBuffer = this.lastKnownBuffer

    return this.saveCoordinator.runDeduplicated(targetFile.path, serialized, async () => {
      const file = targetFile
      const isCurrentSession = () => this.runtime === runtime && this.loadedFile?.path === file.path
      try {
        if (isCurrentSession() && serialized === this.baseline) {
          this.pendingDestructiveSnapshot = undefined
          this.setStatus('saved')
          return false
        }

        const current = await this.app.vault.readBinary(file)
        const expectedBuffer = this.knownBuffers.get(file.path) ?? sessionBuffer
        if (!expectedBuffer || !buffersEqual(current, expectedBuffer))
          throw new Error('The file changed outside this editor. Reopen it before saving to avoid overwriting external changes.')

        if (isCurrentSession())
          this.setStatus('saving')
        const output = await exportDocx(snapshot, current)
        const latest = await this.app.vault.readBinary(file)
        if (!buffersEqual(current, latest))
          throw new Error('The file changed while it was being exported. Reopen it before saving.')
        if (this.settings.createBackups && !this.backedUpFiles.has(file.path)) {
          await createBinaryBackup(this.app, file, latest)
          this.backedUpFiles.add(file.path)
        }
        await this.app.vault.modifyBinary(file, output)
        await pruneBinaryBackups(this.app, file).catch(() => undefined)
        this.knownBuffers.set(file.path, output)
        if (isCurrentSession()) {
          this.lastKnownBuffer = output
          this.baseline = serialized
          this.lastSavedSnapshot = cloneSnapshot(serialized)
          this.pendingDestructiveSnapshot = undefined
          this.setStatus('saved')
        }
        if (reason === 'manual')
          new Notice(`${file.name} saved`)
        return true
      }
      catch (error) {
        if (isCurrentSession())
          this.setStatus('error')
        new Notice(`Cannot save ${file.name}: ${errorMessage(error)}`, 8000)
        throw error
      }
    })
  }

  private scheduleSave(): void {
    if (!hasUnsavedSnapshot(this.getSnapshot(), this.baseline))
      return
    if (this.autosaveTimer !== undefined)
      window.clearTimeout(this.autosaveTimer)
    this.setStatus('dirty')
    this.autosaveTimer = window.setTimeout(() => {
      this.autosaveTimer = undefined
      void this.saveNow('auto').catch(() => undefined)
    }, AUTO_SAVE_DELAY)
  }

  private getSnapshot(): IDocumentData | undefined {
    return this.runtime?.univerAPI.getActiveDocument()?.getSnapshot()
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
    const text = uiText(this.settings.language)
    const labels = {
      loading: text.openingDocument,
      dirty: text.documentDirty,
      saving: text.savingDocument,
      saved: text.documentSaved,
      protected: text.documentProtected,
      error: text.documentError,
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
    const bar = this.contentEl.createDiv({ cls: 'univer-protected-bar' })
    const icon = bar.createSpan({ cls: 'univer-protected-bar__icon' })
    setIcon(icon, 'shield-check')
    const copy = bar.createDiv({ cls: 'univer-protected-bar__copy' })
    this.protectedTitleEl = copy.createEl('strong')
    this.protectedMessageEl = copy.createSpan()
    const editCopy = bar.createEl('button', {
      cls: 'univer-protected-bar__edit',
      attr: {
        type: 'button',
      },
    })
    this.editCopyEl = editCopy
    const editIcon = editCopy.createSpan()
    setIcon(editIcon, 'file-pen-line')
    this.editCopyLabelEl = editCopy.createSpan()
    editCopy.addEventListener('click', () => {
      void this.createEditableCopy(editCopy).catch(() => undefined)
    })
    const details = bar.createEl('button', {
      cls: 'clickable-icon univer-protected-bar__details',
      attr: { type: 'button' },
    })
    this.protectedDetailsEl = details
    setIcon(details, 'info')
    details.addEventListener('click', () => this.showReadOnlyNotice())
    this.refreshProtectedHeader()
    container.before(bar)
  }

  private refreshProtectedHeader(): void {
    const text = uiText(this.settings.language)
    this.protectedTitleEl?.setText(text.previewTitle)
    this.protectedMessageEl?.setText(text.previewMessage)
    this.editCopyLabelEl?.setText(text.editCopy)
    this.editCopyEl?.setAttrs({ 'aria-label': text.editCopyTitle, 'title': text.editCopyTitle })
    this.protectedDetailsEl?.setAttrs({ 'aria-label': text.protectedDetails, 'title': text.protectedDetails })
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
    this.loadedFile = undefined
    this.fileDeleted = false
    this.readOnly = false
    this.lastKnownBuffer = undefined
    this.protectedReasons = []
    this.lastReadOnlyNoticeAt = 0
    this.creatingEditableCopy = false
    this.protectedTitleEl = undefined
    this.protectedMessageEl = undefined
    this.editCopyEl = undefined
    this.editCopyLabelEl = undefined
    this.protectedDetailsEl = undefined
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
