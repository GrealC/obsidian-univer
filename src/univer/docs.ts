import type { UniverRuntime } from './create'
import type { UniverPluginSettings } from '@/types/setting'
import { UniverDocsCorePreset } from '@univerjs/preset-docs-core'
import enUS from '@univerjs/preset-docs-core/locales/en-US'
import ruRU from '@univerjs/preset-docs-core/locales/ru-RU'
import viVN from '@univerjs/preset-docs-core/locales/vi-VN'
import zhCN from '@univerjs/preset-docs-core/locales/zh-CN'
import zhTW from '@univerjs/preset-docs-core/locales/zh-TW'
import { defaultTheme } from '@univerjs/themes'
import { getLanguage } from '@/utils/common'
import { createUniver } from './create'
import { registerLocalFonts } from './fonts'
import { installLocalFontPicker } from './localFontPicker'

import '@univerjs/preset-docs-core/lib/index.css'

const localePacks = { EN: enUS, ZH: zhCN, RU: ruRU, VN: viVN, TW: zhTW }

export function docInit(container: HTMLElement, settings: UniverPluginSettings): UniverRuntime {
  const locale = getLanguage(settings)
  const runtime = createUniver({
    locale,
    locales: {
      [locale]: localePacks[settings.language],
    },
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
