import { describe, expect, it } from 'vitest'
import { createLocalFontConfigs, matchesFontSearch, parseInstalledFontNames } from './fonts'

describe('local font discovery', () => {
  it('keeps installed families, trims names, and removes duplicates or unsafe values', () => {
    expect(createLocalFontConfigs([' Zed Sans ', 'Arial', 'arial', 'Courier Mono', 'Bad\nName', 'x'.repeat(129)])).toEqual([
      { value: 'Arial', label: 'Arial', category: 'sans-serif' },
      { value: 'Courier Mono', label: 'Courier Mono', category: 'monospace' },
      { value: 'Zed Sans', label: 'Zed Sans', category: 'sans-serif' },
    ])
  })

  it('parses the PowerShell JSON response in either supported shape', () => {
    expect(parseInstalledFontNames('["Calibri","Microsoft YaHei"]')).toEqual(['Calibri', 'Microsoft YaHei'])
    expect(parseInstalledFontNames('"Consolas"')).toEqual(['Consolas'])
    expect(parseInstalledFontNames('[{"family":"Microsoft YaHei","label":"微软雅黑"}]')).toEqual(['Microsoft YaHei'])
    expect(parseInstalledFontNames('not json')).toEqual([])
  })

  it('uses localized font labels for Chinese search while preserving the CSS family name', () => {
    const [font] = createLocalFontConfigs([{ family: 'Microsoft YaHei', label: '微软雅黑' }])

    expect(font).toEqual({ value: 'Microsoft YaHei', label: '微软雅黑', category: 'sans-serif' })
    expect(matchesFontSearch(font, '微软雅黑')).toBe(true)
    expect(matchesFontSearch(font, 'yahei')).toBe(true)
  })

  it('prefers the localized Windows label when the browser reports the same family', () => {
    const [font] = createLocalFontConfigs([
      { family: 'Microsoft YaHei' },
      { family: 'Microsoft YaHei', label: '微软雅黑' },
    ])

    expect(font.label).toBe('微软雅黑')
  })

  it('adds Chinese search labels to locally discovered CJK fonts without adding new fonts', () => {
    const fonts = createLocalFontConfigs(['KaiTi', 'STKaiti', 'LXGW WenKai'])

    expect(fonts.map(font => [font.value, font.label])).toEqual([
      ['KaiTi', '楷体'],
      ['LXGW WenKai', '霞鹜文楷'],
      ['STKaiti', '华文楷体'],
    ])
    expect(fonts.filter(font => matchesFontSearch(font, '楷')).map(font => font.value)).toEqual([
      'KaiTi',
      'LXGW WenKai',
      'STKaiti',
    ])
  })
})
