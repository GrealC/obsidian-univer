import type { UniverPluginSettings } from '@/types/setting'
import { defu } from 'defu'
import { addIcon, Plugin } from 'obsidian'
import { createNewFile } from '@/utils/file'
import { ChooseTypeModal } from './modals/chooseType'
import { SettingTab } from './modals/settingTab'
import { univerIconSvg } from './utils/common'
import { Type as UDocType, UDocView } from './views/udoc'
import { Type as USheetType, USheetView } from './views/usheet'
import { Type as XlsxType, XlsxTypeView } from './views/xlsx'
import './style/univer.css'

export type ViewType = typeof USheetType | typeof UDocType | typeof XlsxType
export default class UniverPlugin extends Plugin {
  settings: UniverPluginSettings
  async onload() {
    await this.loadSettings()

    addIcon('univer', univerIconSvg)

    // ribbon icon & the class
    this.addRibbonIcon('univer', 'univer-plus', () => {
      const modal = new ChooseTypeModal(this.app, this.settings)
      modal.open()
    })

    this.addCommand({
      id: 'univer-sheet',
      name: 'Create Univer Sheet',
      callback: () => {
        createNewFile(this.app, 'usheet')
      },
    })

    this.addCommand({
      id: 'univer-doc',
      name: 'Create Univer Doc',
      callback: () => {
        createNewFile(this.app, 'udoc')
      },
    })

    this.addCommand({
      id: 'univer-xlsx',
      name: 'Create Univer Xlsx',
      checkCallback: (checking) => {
        if (!this.settings.isSupportXlsx)
          return false
        if (!checking)
          void createNewFile(this.app, 'xlsx')
        return true
      },
    })

    // add the setting tab
    this.addSettingTab(new SettingTab(this.app, this))
    // register view
    this.registerView(USheetType, leaf => new USheetView(leaf, this.settings))
    this.registerExtensions(['usheet'], USheetType)

    this.registerView(UDocType, leaf => new UDocView(leaf, this.settings))
    this.registerExtensions(['udoc'], UDocType)

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
      createBackups: true,
    })
  }

  async saveSettings() {
    await this.saveData(this.settings)
  }
}
