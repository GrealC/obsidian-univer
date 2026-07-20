import type { App } from 'obsidian'
import type UniverPlugin from '../main'
import type { UniverLanguage } from '@/types/setting'
import { Notice, PluginSettingTab, Setting } from 'obsidian'
import { uiText } from '@/i18n'

export class SettingTab extends PluginSettingTab {
  plugin: UniverPlugin

  constructor(app: App, plugin: UniverPlugin) {
    super(app, plugin)
    this.plugin = plugin
  }

  display(): void {
    const { containerEl } = this
    const text = uiText(this.plugin.settings.language)
    containerEl.empty()
    new Setting(containerEl).setName('univer-plus').setHeading()

    new Setting(containerEl)
      .setName(text.settingsLanguageName)
      .setDesc(text.settingsLanguageDesc)
      .addDropdown((drop) => {
        drop
          .addOptions({
            EN: 'English',
            ZH: '简体中文',
            RU: 'Русский',
            VN: 'Tiếng Việt',
            TW: '繁體中文',
          })
          .setValue(this.plugin.settings.language)
          .onChange(async (value: UniverLanguage) => {
            await this.plugin.updateLanguage(value)
            this.display()
          })
      })
    new Setting(containerEl)
      .setName(text.settingsExcelName)
      .setDesc(text.settingsExcelDesc)
      .addToggle((toggle) => {
        toggle
          .setValue(this.plugin.settings.isSupportXlsx)
          .onChange(async (value: boolean) => {
            this.plugin.settings.isSupportXlsx = value
            await this.plugin.saveSettings()
            new Notice(text.restartExcel)
          })
      })
    new Setting(containerEl)
      .setName(text.settingsWordName)
      .setDesc(text.settingsWordDesc)
      .addToggle((toggle) => {
        toggle
          .setValue(this.plugin.settings.isSupportDocx)
          .onChange(async (value: boolean) => {
            this.plugin.settings.isSupportDocx = value
            await this.plugin.saveSettings()
            new Notice(text.restartWord)
          })
      })
    new Setting(containerEl)
      .setName(text.settingsBackupName)
      .setDesc(text.settingsBackupDesc)
      .addToggle((toggle) => {
        toggle
          .setValue(this.plugin.settings.createBackups)
          .onChange(async (value: boolean) => {
            this.plugin.settings.createBackups = value
            await this.plugin.saveSettings()
          })
      })
  }
}
