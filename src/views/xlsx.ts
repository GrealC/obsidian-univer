import type { IDisposable, IWorkbookData } from '@univerjs/core'
import type { IconName, TFile, WorkspaceLeaf } from 'obsidian'
import type { UniverPluginSettings } from '@/types/setting'
import type { UniverRuntime } from '@/univer/create'
import { CommandType } from '@univerjs/core'
import { FileView, Notice, setIcon } from 'obsidian'
import { inspectXlsx } from '@/excel/capabilities'
import { exportXlsx, importXlsx } from '@/excel/converter'
import { assertSafeWorkbookTransition } from '@/excel/safety'
import { uiText } from '@/i18n'
import { createXlsxBackup, pruneBinaryBackups } from '@/services/backup'
import { SaveCoordinator } from '@/services/saveCoordinator'
import { sheetInit } from '@/univer/sheets'
import { observeTheme } from '@/univer/theme'
import { getLanguage } from '@/utils/common'

export const Type = 'univer-xlsx'
const AUTO_SAVE_DELAY = 2000
type WorkbookStatus = 'loading' | 'dirty' | 'saving' | 'saved' | 'protected' | 'error'

const STATUS_ICONS: Record<WorkbookStatus, IconName> = {
  loading: 'loader-circle',
  dirty: 'circle',
  saving: 'loader-circle',
  saved: 'check',
  protected: 'shield-alert',
  error: 'circle-alert',
}

export class XlsxTypeView extends FileView {
  private readonly saveCoordinator = new SaveCoordinator()
  private runtime?: UniverRuntime
  private commandDisposable?: IDisposable
  private themeObserver?: MutationObserver
  private autosaveTimer?: number
  private statusActionEl?: HTMLElement
  private statusLiveEl?: HTMLElement
  private statusState?: WorkbookStatus
  private baseline = ''
  private lastSavedSnapshot?: IWorkbookData
  private pendingDestructiveSnapshot?: string
  private fileDeleted = false
  private readOnly = false
  private backupCreated = false
  private lastKnownBuffer?: ArrayBuffer
  private lastReadOnlyNoticeAt = 0

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
  }

  async onOpen(): Promise<void> {
    this.ensureStatusAction()
    this.registerEvent(this.app.vault.on('delete', (file) => {
      if (file === this.file)
        this.fileDeleted = true
    }))

    this.registerDomEvent(window, 'keydown', (event) => {
      const isSave = (event.ctrlKey || event.metaKey) && !event.altKey && event.key.toLowerCase() === 's'
      if (!isSave || this.app.workspace.getActiveViewOfType(XlsxTypeView) !== this)
        return

      event.preventDefault()
      void this.saveNow('manual').catch(() => undefined)
    })
  }

  async onLoadFile(file: TFile): Promise<void> {
    await super.onLoadFile(file)
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
    this.lastReadOnlyNoticeAt = 0
    this.contentEl.empty()
    this.contentEl.addClass('univer-view', 'univer-xlsx-view')

    if (!this.settings.isSupportXlsx) {
      this.contentEl.createDiv({ cls: 'univer-empty-state', text: 'Excel support is disabled. Re-enable it in Univer settings and restart Obsidian.' })
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
      const inspection = await inspectXlsx(raw)
      const workbookData = await importXlsx(raw, file.basename, getLanguage(this.settings))
      this.runtime = sheetInit(container, this.settings)
      this.runtime.univerAPI.createWorkbook(workbookData)
      this.themeObserver = observeTheme(this.runtime.univerAPI)
      const initialSnapshot = this.runtime.univerAPI.getActiveWorkbook()?.save()
      this.baseline = initialSnapshot ? JSON.stringify(initialSnapshot) : ''
      this.lastSavedSnapshot = this.baseline ? cloneSnapshot(this.baseline) : undefined
      this.readOnly = inspection.readOnly

      if (this.readOnly) {
        container.classList.add('univer-editor-readonly')
        container.setAttr('aria-readonly', 'true')
        const warning = this.contentEl.createDiv({
          cls: 'univer-readonly-warning',
          text: `Protected view: ${inspection.reasons.join(', ')} cannot be preserved safely. The original file will not be overwritten.`,
        })
        container.before(warning)
        this.commandDisposable = this.runtime.univerAPI.addEvent(this.runtime.univerAPI.Event.BeforeCommandExecute, (event) => {
          if (event.type !== CommandType.MUTATION)
            return
          event.cancel = true
          this.showReadOnlyNotice()
        })
        this.setStatus('protected')
      }
      else {
        this.commandDisposable = this.runtime.univerAPI.addEvent(this.runtime.univerAPI.Event.CommandExecuted, (event) => {
          if (event.type === CommandType.MUTATION)
            this.scheduleSave()
        })
        this.setStatus('saved')
      }
    }
    catch (error) {
      this.disposeEditor()
      container.createDiv({ cls: 'univer-error', text: `Cannot open this workbook: ${errorMessage(error)}` })
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
      // saveNow already reports the error and leaves the original file untouched.
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
        new Notice('This workbook is read-only because it contains unsupported Excel features.')
      return false
    }

    return this.saveCoordinator.run(async () => {
      const file = this.file
      if (!file || this.fileDeleted)
        return false

      try {
        const snapshot = this.runtime?.univerAPI.getActiveWorkbook()?.save()
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
        const output = await exportXlsx(snapshot as IWorkbookData)
        const latest = await this.app.vault.readBinary(file)
        if (!buffersEqual(current, latest))
          throw new Error('The file changed while it was being exported. Reopen it before saving.')
        if (this.settings.createBackups && !this.backupCreated) {
          await createXlsxBackup(this.app, file, latest)
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

  private assertSafeTransition(snapshot: IWorkbookData, serialized: string, reason: 'auto' | 'manual' | 'close'): void {
    if (!this.lastSavedSnapshot) {
      this.pendingDestructiveSnapshot = undefined
      return
    }

    try {
      assertSafeWorkbookTransition(this.lastSavedSnapshot, snapshot)
      this.pendingDestructiveSnapshot = undefined
    }
    catch (error) {
      if (reason === 'manual' && this.pendingDestructiveSnapshot === serialized) {
        this.pendingDestructiveSnapshot = undefined
        return
      }

      this.pendingDestructiveSnapshot = serialized
      throw new Error(`${errorMessage(error)}. Press Ctrl/Cmd+S again to confirm that you intended to clear the workbook.`)
    }
  }

  private ensureStatusAction(): void {
    if (this.statusActionEl)
      return
    this.statusActionEl = this.addAction('loader-circle', this.statusLabel('loading'), () => this.handleStatusAction())
    this.statusActionEl.addClass('univer-save-status')
  }

  private setStatus(status: WorkbookStatus): void {
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

  private statusLabel(status: WorkbookStatus): string {
    const text = uiText(this.settings.language)
    const labels = {
      loading: text.openingWorkbook,
      dirty: text.workbookDirty,
      saving: text.savingWorkbook,
      saved: text.workbookSaved,
      protected: text.workbookProtected,
      error: text.workbookError,
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
    new Notice('This workbook is protected because it contains Excel features that cannot be saved safely.')
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

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

function cloneSnapshot(serialized: string): IWorkbookData {
  return JSON.parse(serialized) as IWorkbookData
}
