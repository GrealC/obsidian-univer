import type { App } from 'obsidian'
import type UniverPlugin from '../main'
import { Notice, PluginSettingTab, Setting } from 'obsidian'

export class SettingTab extends PluginSettingTab {
  plugin: UniverPlugin

  constructor(app: App, plugin: UniverPlugin) {
    super(app, plugin)
    this.plugin = plugin
  }

  display(): void {
    const { containerEl } = this
    containerEl.empty()
    new Setting(containerEl).setName('univer-plus').setHeading()

    new Setting(containerEl)
      .setName('Language')
      .setDesc('Language used by new Univer editor views')
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
          .onChange(async (value: 'ZH' | 'EN' | 'RU' | 'TW' | 'VN') => {
            this.plugin.settings.language = value
            await this.plugin.saveSettings()
          })
      })
    new Setting(containerEl)
      .setName('Open Excel workbooks')
      .setDesc('Register the .xlsx editor. Restart Obsidian after changing this setting. Legacy .xls and macro-enabled .xlsm files are not modified.')
      .addToggle((toggle) => {
        toggle
          .setValue(this.plugin.settings.isSupportXlsx)
          .onChange(async (value: boolean) => {
            this.plugin.settings.isSupportXlsx = value
            await this.plugin.saveSettings()
            new Notice('Restart Obsidian to apply the Excel file handler change.')
          })
      })
    new Setting(containerEl)
      .setName('Back up Excel files')
      .setDesc('Create one backup before the first save in each editor session')
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
