import type { IFontConfig } from '@univerjs/preset-sheets-core'
import type { KeyboardEvent } from 'react'
import type { UniverRuntime } from './create'
import { ICommandService, LocaleService } from '@univerjs/core'
import {
  FONT_FAMILY_ITEM_COMPONENT,
  IFontService,
  ILayoutService,
  useDependency,
  useObservable,
} from '@univerjs/preset-sheets-core'
import { useDeferredValue, useMemo, useState } from 'react'
import { matchesFontSearch } from './fonts'

const MAX_VISIBLE_FONTS = 160

export function installLocalFontPicker(runtime: UniverRuntime): void {
  runtime.univerAPI.registerComponent(FONT_FAMILY_ITEM_COMPONENT, LocalFontFamilyItem)
}

export function filterLocalFonts(fonts: readonly IFontConfig[], search: string, limit = MAX_VISIBLE_FONTS): IFontConfig[] {
  const matchingFonts = search.trim()
    ? fonts.filter(font => matchesFontSearch(font, search))
    : fonts

  return matchingFonts.slice(0, limit)
}

function LocalFontFamilyItem({ id, value }: { id: string, value: string }) {
  const commandService = useDependency(ICommandService)
  const fontService = useDependency(IFontService)
  const layoutService = useDependency(ILayoutService)
  const localeService = useDependency(LocaleService)
  const fonts = useObservable<IFontConfig[]>(fontService.fonts$, [], true)
  const [search, setSearch] = useState('')
  const deferredSearch = useDeferredValue(search)
  const visibleFonts = useMemo(() => filterLocalFonts(fonts, deferredSearch), [deferredSearch, fonts])

  function selectFont(font: IFontConfig): void {
    layoutService.focus()
    void commandService.executeCommand(id, { value: font.value })
  }

  function handleSearchKeyDown(event: KeyboardEvent<HTMLInputElement>): void {
    if (event.key === 'Escape')
      return

    event.stopPropagation()
    if (event.key !== 'Enter')
      return

    event.preventDefault()
    const firstMatch = visibleFonts[0]
    if (firstMatch)
      selectFont(firstMatch)
  }

  return (
    <div className="univer-plus-font-picker">
      <div className="univer-plus-font-picker__search-wrap">
        <input
          autoFocus
          aria-label="Search local fonts"
          className="univer-plus-font-picker__search"
          onChange={event => setSearch(event.target.value)}
          onKeyDown={handleSearchKeyDown}
          placeholder="Search fonts"
          type="search"
          value={search}
        />
      </div>
      <ul aria-label="Local fonts" className="univer-plus-font-picker__list" role="listbox">
        {visibleFonts.map(font => (
          <li key={font.value}>
            <button
              aria-selected={font.value === value}
              className="univer-plus-font-picker__option"
              data-current={String(font.value === value)}
              onClick={() => selectFont(font)}
              role="option"
              style={{ fontFamily: quoteFontFamily(font.value) }}
              title={localeService.t(font.label)}
              type="button"
            >
              {localeService.t(font.label)}
            </button>
          </li>
        ))}
        {visibleFonts.length === 0 && (
          <li className="univer-plus-font-picker__empty" role="presentation">
            {search.trim() ? 'No matching fonts' : 'No local fonts found'}
          </li>
        )}
      </ul>
    </div>
  )
}

function quoteFontFamily(value: string): string {
  return `"${value.replace(/["\\]/g, '\\$&')}"`
}
