export type UniverLanguage = 'ZH' | 'EN' | 'RU' | 'VN' | 'TW'

export interface UniverPluginSettings {
  language: UniverLanguage
  isSupportXlsx: boolean
  isSupportDocx: boolean
  createBackups: boolean
}
