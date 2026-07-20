import type { UniverLanguage } from './types/setting'
import { describe, expect, it } from 'vitest'
import { uiText, univerLocaleText } from './i18n'

describe('plugin translations', () => {
  it('provides the same complete key set for every supported language', () => {
    const languages: UniverLanguage[] = ['EN', 'ZH', 'TW', 'RU', 'VN']
    const englishKeys = Object.keys(uiText('EN')).sort()

    for (const language of languages) {
      expect(Object.keys(uiText(language)).sort()).toEqual(englishKeys)
      expect(univerLocaleText(language).univerPlus.searchFonts).not.toBe('')
    }
  })
})
