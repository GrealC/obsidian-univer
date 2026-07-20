import type { UniverRuntime } from './create'
import type { UniverPluginSettings } from '@/types/setting'
import { LocaleType, mergeLocales } from '@univerjs/core'
import { UniverDocsCorePreset } from '@univerjs/preset-docs-core'
import enUS from '@univerjs/preset-docs-core/locales/en-US'
import ruRU from '@univerjs/preset-docs-core/locales/ru-RU'
import viVN from '@univerjs/preset-docs-core/locales/vi-VN'
import zhCN from '@univerjs/preset-docs-core/locales/zh-CN'
import zhTW from '@univerjs/preset-docs-core/locales/zh-TW'
import { defaultTheme } from '@univerjs/themes'
import { univerLocaleText } from '@/i18n'
import { getLanguage } from '@/utils/common'
import { createUniver } from './create'
import { registerLocalFonts } from './fonts'
import { installLocalFontPicker } from './localFontPicker'

import '@univerjs/preset-docs-core/lib/index.css'

const localePacks = { EN: enUS, ZH: zhCN, RU: ruRU, VN: viVN, TW: zhTW }
const locales = {
  [LocaleType.EN_US]: mergeLocales(localePacks.EN, univerLocaleText('EN')),
  [LocaleType.ZH_CN]: mergeLocales(localePacks.ZH, univerLocaleText('ZH')),
  [LocaleType.RU_RU]: mergeLocales(localePacks.RU, univerLocaleText('RU')),
  [LocaleType.VI_VN]: mergeLocales(localePacks.VN, univerLocaleText('VN')),
  [LocaleType.ZH_TW]: mergeLocales(localePacks.TW, univerLocaleText('TW')),
}

export function docInit(container: HTMLElement, settings: UniverPluginSettings): UniverRuntime {
  const locale = getLanguage(settings)
  const runtime = createUniver({
    locale,
    locales,
    theme: defaultTheme,
    presets: [
      UniverDocsCorePreset({
        container,
        header: true,
        disableAutoFocus: true,
      }),
    ],
  })

  installLocalFontPicker(runtime)
  void registerLocalFonts(runtime)

  return runtime
}
