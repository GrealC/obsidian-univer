import type { UniverRuntime } from './create'
import type { UniverPluginSettings } from '@/types/setting'
import { mergeLocales } from '@univerjs/core'
import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting'
import conditionalEnUS from '@univerjs/preset-sheets-conditional-formatting/locales/en-US'
import conditionalRuRU from '@univerjs/preset-sheets-conditional-formatting/locales/ru-RU'
import conditionalViVN from '@univerjs/preset-sheets-conditional-formatting/locales/vi-VN'
import conditionalZhCN from '@univerjs/preset-sheets-conditional-formatting/locales/zh-CN'
import conditionalZhTW from '@univerjs/preset-sheets-conditional-formatting/locales/zh-TW'
import {
  CancelFrozenCommand,
  InsertColBeforeCommand,
  InsertRowBeforeCommand,
  RibbonInsertGroup,
  RibbonPosition,
  RibbonViewGroup,
  SetFirstColumnFrozenCommand,
  SetFirstRowFrozenCommand,
  ToggleGridlinesCommand,
  UniverSheetsCorePreset,
} from '@univerjs/preset-sheets-core'
import coreEnUS from '@univerjs/preset-sheets-core/locales/en-US'
import coreRuRU from '@univerjs/preset-sheets-core/locales/ru-RU'
import coreViVN from '@univerjs/preset-sheets-core/locales/vi-VN'
import coreZhCN from '@univerjs/preset-sheets-core/locales/zh-CN'
import coreZhTW from '@univerjs/preset-sheets-core/locales/zh-TW'
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation'
import validationEnUS from '@univerjs/preset-sheets-data-validation/locales/en-US'
import validationRuRU from '@univerjs/preset-sheets-data-validation/locales/ru-RU'
import validationViVN from '@univerjs/preset-sheets-data-validation/locales/vi-VN'
import validationZhCN from '@univerjs/preset-sheets-data-validation/locales/zh-CN'
import validationZhTW from '@univerjs/preset-sheets-data-validation/locales/zh-TW'
import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter'
import filterEnUS from '@univerjs/preset-sheets-filter/locales/en-US'
import filterRuRU from '@univerjs/preset-sheets-filter/locales/ru-RU'
import filterViVN from '@univerjs/preset-sheets-filter/locales/vi-VN'
import filterZhCN from '@univerjs/preset-sheets-filter/locales/zh-CN'
import filterZhTW from '@univerjs/preset-sheets-filter/locales/zh-TW'
import { UniverSheetsFindReplacePreset } from '@univerjs/preset-sheets-find-replace'
import findEnUS from '@univerjs/preset-sheets-find-replace/locales/en-US'
import findRuRU from '@univerjs/preset-sheets-find-replace/locales/ru-RU'
import findViVN from '@univerjs/preset-sheets-find-replace/locales/vi-VN'
import findZhCN from '@univerjs/preset-sheets-find-replace/locales/zh-CN'
import findZhTW from '@univerjs/preset-sheets-find-replace/locales/zh-TW'
import { UniverSheetsHyperLinkPreset } from '@univerjs/preset-sheets-hyper-link'
import hyperlinkEnUS from '@univerjs/preset-sheets-hyper-link/locales/en-US'
import hyperlinkRuRU from '@univerjs/preset-sheets-hyper-link/locales/ru-RU'
import hyperlinkViVN from '@univerjs/preset-sheets-hyper-link/locales/vi-VN'
import hyperlinkZhCN from '@univerjs/preset-sheets-hyper-link/locales/zh-CN'
import hyperlinkZhTW from '@univerjs/preset-sheets-hyper-link/locales/zh-TW'
import { UniverSheetsSortPreset } from '@univerjs/preset-sheets-sort'
import sortEnUS from '@univerjs/preset-sheets-sort/locales/en-US'
import sortRuRU from '@univerjs/preset-sheets-sort/locales/ru-RU'
import sortViVN from '@univerjs/preset-sheets-sort/locales/vi-VN'
import sortZhCN from '@univerjs/preset-sheets-sort/locales/zh-CN'
import sortZhTW from '@univerjs/preset-sheets-sort/locales/zh-TW'
import { defaultTheme } from '@univerjs/themes'
import { getLanguage } from '@/utils/common'
import { createUniver } from './create'
import { registerLocalFonts } from './fonts'
import { installLocalFontPicker } from './localFontPicker'

import '@univerjs/preset-sheets-core/lib/index.css'
import '@univerjs/preset-sheets-conditional-formatting/lib/index.css'
import '@univerjs/preset-sheets-data-validation/lib/index.css'
import '@univerjs/preset-sheets-filter/lib/index.css'
import '@univerjs/preset-sheets-find-replace/lib/index.css'
import '@univerjs/preset-sheets-hyper-link/lib/index.css'
import '@univerjs/preset-sheets-sort/lib/index.css'

const localePacks = {
  EN: mergeLocales(coreEnUS, conditionalEnUS, validationEnUS, filterEnUS, findEnUS, hyperlinkEnUS, sortEnUS),
  ZH: mergeLocales(coreZhCN, conditionalZhCN, validationZhCN, filterZhCN, findZhCN, hyperlinkZhCN, sortZhCN),
  RU: mergeLocales(coreRuRU, conditionalRuRU, validationRuRU, filterRuRU, findRuRU, hyperlinkRuRU, sortRuRU),
  VN: mergeLocales(coreViVN, conditionalViVN, validationViVN, filterViVN, findViVN, hyperlinkViVN, sortViVN),
  TW: mergeLocales(coreZhTW, conditionalZhTW, validationZhTW, filterZhTW, findZhTW, hyperlinkZhTW, sortZhTW),
}

export function sheetInit(container: HTMLElement, settings: UniverPluginSettings): UniverRuntime {
  const locale = getLanguage(settings)
  const runtime = createUniver({
    locale,
    locales: {
      [locale]: localePacks[settings.language],
    },
    theme: defaultTheme,
    presets: [
      UniverSheetsCorePreset({
        container,
        header: true,
        toolbar: true,
        formulaBar: true,
        contextMenu: true,
        ribbonType: 'classic',
        footer: {
          sheetBar: true,
          statisticBar: true,
          menus: true,
          zoomSlider: true,
        },
        customFontFamily: {
          override: true,
          list: [],
        },
        disableAutoFocus: true,
      }),
      UniverSheetsConditionalFormattingPreset(),
      UniverSheetsDataValidationPreset(),
      UniverSheetsFilterPreset(),
      UniverSheetsFindReplacePreset(),
      UniverSheetsHyperLinkPreset(),
      UniverSheetsSortPreset(),
    ],
  })

  installBasicMenus(runtime, settings.language)
  installLocalFontPicker(runtime)
  void registerLocalFonts(runtime)

  return runtime
}

function installBasicMenus(runtime: UniverRuntime, language: UniverPluginSettings['language']): void {
  const zh = language === 'ZH' || language === 'TW'
  const labels = zh
    ? {
        insertRow: '在上方插入一行',
        insertColumn: '在左侧插入一列',
        freezeRow: '冻结首行',
        freezeColumn: '冻结首列',
        unfreeze: '取消冻结窗格',
        gridlines: '显示或隐藏网格线',
      }
    : {
        insertRow: 'Insert row above',
        insertColumn: 'Insert column left',
        freezeRow: 'Freeze top row',
        freezeColumn: 'Freeze first column',
        unfreeze: 'Unfreeze panes',
        gridlines: 'Toggle gridlines',
      }

  runtime.univerAPI.createMenu({
    id: 'univer-plus.insert-row-before',
    title: labels.insertRow,
    action: () => { void runtime.univerAPI.executeCommand(InsertRowBeforeCommand.id, { value: 1 }) },
  }).appendTo([RibbonPosition.INSERT, RibbonInsertGroup.EDIT])
  runtime.univerAPI.createMenu({
    id: 'univer-plus.insert-column-before',
    title: labels.insertColumn,
    action: () => { void runtime.univerAPI.executeCommand(InsertColBeforeCommand.id, { value: 1 }) },
  }).appendTo([RibbonPosition.INSERT, RibbonInsertGroup.EDIT])

  for (const [id, title, command] of [
    ['freeze-row', labels.freezeRow, SetFirstRowFrozenCommand.id],
    ['freeze-column', labels.freezeColumn, SetFirstColumnFrozenCommand.id],
    ['unfreeze', labels.unfreeze, CancelFrozenCommand.id],
    ['gridlines', labels.gridlines, ToggleGridlinesCommand.id],
  ] as const) {
    runtime.univerAPI.createMenu({
      id: `univer-plus.${id}`,
      title,
      action: () => { void runtime.univerAPI.executeCommand(command) },
    }).appendTo([RibbonPosition.VIEW, RibbonViewGroup.DISPLAY])
  }
}
