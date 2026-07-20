import type { Command, TAbstractFile } from 'obsidian'
import type { UniverLanguage, UniverPluginSettings } from '@/types/setting'
import { defu } from 'defu'
import { addIcon, Plugin, TFolder } from 'obsidian'
import { uiText } from '@/i18n'
import { createNewFile } from '@/utils/file'
import { ChooseTypeModal } from './modals/chooseType'
import { SettingTab } from './modals/settingTab'
import { univerIconSvg } from './utils/common'
import { Type as DocxType, DocxTypeView } from './views/docx'
import { Type as UDocType, UDocView } from './views/udoc'
import { Type as USheetType, USheetView } from './views/usheet'
import { Type as XlsxType, XlsxTypeView } from './views/xlsx'
import './style/univer.css'

export type ViewType = typeof USheetType | typeof UDocType | typeof XlsxType | typeof DocxType
type CommandTextKey = 'commandSheet' | 'commandDoc' | 'commandExcel' | 'commandWord'

export default class UniverPlugin extends Plugin {
  settings: UniverPluginSettings
  private ribbonEl?: HTMLElement
  private readonly localizedCommands: Array<{ command: Command, key: CommandTextKey }> = []

  async onload() {
    await this.loadSettings()
    const text = uiText(this.settings.language)

    addIcon('univer', univerIconSvg)

    // ribbon icon & the class
    this.ribbonEl = this.addRibbonIcon('univer', text.ribbonTitle, () => {
      const modal = new ChooseTypeModal(this.app, this.settings)
      modal.open()
    })

    this.localizedCommands.push({ command: this.addCommand({
      id: 'univer-sheet',
      name: text.commandSheet,
      callback: () => {
        void createNewFile(this.app, 'usheet', this.settings.language)
      },
    }), key: 'commandSheet' })

    this.localizedCommands.push({ command: this.addCommand({
      id: 'univer-doc',
      name: text.commandDoc,
      callback: () => {
        void createNewFile(this.app, 'udoc', this.settings.language)
      },
    }), key: 'commandDoc' })

    this.localizedCommands.push({ command: this.addCommand({
      id: 'univer-docx',
      name: text.commandWord,
      checkCallback: (checking) => {
        if (!this.settings.isSupportDocx)
          return false
        if (!checking)
          void createNewFile(this.app, 'docx', this.settings.language)
        return true
      },
    }), key: 'commandWord' })

    this.localizedCommands.push({ command: this.addCommand({
      id: 'univer-xlsx',
      name: text.commandExcel,
      checkCallback: (checking) => {
        if (!this.settings.isSupportXlsx)
          return false
        if (!checking)
          void createNewFile(this.app, 'xlsx', this.settings.language)
        return true
      },
    }), key: 'commandExcel' })

    this.registerEvent(this.app.workspace.on('file-menu', (menu, file) => {
      if (!this.settings.isSupportXlsx && !this.settings.isSupportDocx)
        return

      const folderPath = creationFolderPath(file)
      const menuText = uiText(this.settings.language)
      menu.addSeparator()
      if (this.settings.isSupportXlsx) {
        menu.addItem(item => item
          .setTitle(menuText.newExcel)
          .setIcon('file-spreadsheet')
          .onClick(() => { void createNewFile(this.app, 'xlsx', this.settings.language, folderPath) }))
      }
      if (this.settings.isSupportDocx) {
        menu.addItem(item => item
          .setTitle(menuText.newWord)
          .setIcon('file-text')
          .onClick(() => { void createNewFile(this.app, 'docx', this.settings.language, folderPath) }))
      }
    }))

    // add the setting tab
    this.addSettingTab(new SettingTab(this.app, this))
    // register view
    this.registerView(USheetType, leaf => new USheetView(leaf, this.settings))
    this.registerExtensions(['usheet'], USheetType)

    this.registerView(UDocType, leaf => new UDocView(leaf, this.settings))
    this.registerExtensions(['udoc'], UDocType)

    if (this.settings.isSupportDocx) {
      this.registerView(DocxType, leaf => new DocxTypeView(leaf, this.settings))
      this.registerExtensions(['docx'], DocxType)
    }

    if (this.settings.isSupportXlsx) {
      this.registerView(XlsxType, leaf => new XlsxTypeView(leaf, this.settings))
      this.registerExtensions(['xlsx'], XlsxType)
    }
  }

  async loadSettings() {
    const loadedSettings = await this.loadData()
    this.settings = defu(loadedSettings, {
      language: 'EN',
      isSupportXlsx: true,
      isSupportDocx: true,
      createBackups: true,
    })
  }

  async saveSettings() {
    await this.saveData(this.settings)
  }

  async updateLanguage(language: UniverLanguage): Promise<void> {
    this.settings.language = language
    await this.saveSettings()
    this.refreshLocalizedChrome()

    for (const viewType of [USheetType, UDocType, XlsxType, DocxType]) {
      for (const leaf of this.app.workspace.getLeavesOfType(viewType)) {
        const view = leaf.view
        if (view instanceof USheetView || view instanceof UDocView || view instanceof XlsxTypeView || view instanceof DocxTypeView)
          view.setLanguage()
      }
    }
  }

  private refreshLocalizedChrome(): void {
    const text = uiText(this.settings.language)
    this.ribbonEl?.setAttrs({ 'aria-label': text.ribbonTitle, 'title': text.ribbonTitle })
    for (const { command, key } of this.localizedCommands)
      command.name = text[key]
  }
}

function creationFolderPath(file: TAbstractFile): string | undefined {
  const folder = file instanceof TFolder ? file : file.parent
  return !folder || folder.isRoot() ? undefined : folder.path
}
