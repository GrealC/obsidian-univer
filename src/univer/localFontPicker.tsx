import type { IFontConfig } from '@univerjs/preset-sheets-core'
import type { KeyboardEvent, PointerEvent as ReactPointerEvent } from 'react'
import type { UniverRuntime } from './create'
import { ICommandService, LocaleService } from '@univerjs/core'
import {
  FONT_FAMILY_ITEM_COMPONENT,
  IFontService,
  ILayoutService,
  useDependency,
  useObservable,
} from '@univerjs/preset-sheets-core'
import { useDeferredValue, useLayoutEffect, useMemo, useRef, useState } from 'react'
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

export function keepFontSearchOpen(event: { type: string, stopPropagation: () => void }): void {
  // Radix needs pointerdown to reach the menu item or pointerup synthesizes an outer click.
  if (event.type !== 'pointerdown')
    event.stopPropagation()
}

export function preventFontMenuSelection(event: { preventDefault: () => void }, searchInteraction: boolean): void {
  if (searchInteraction)
    event.preventDefault()
}

export function bindFontMenuSelectionGuard(menuItem: EventTarget, isSearchInteraction: () => boolean): () => void {
  const handleMenuItemSelect = (event: Event) => {
    preventFontMenuSelection(event, isSearchInteraction())
  }
  menuItem.addEventListener('menu.itemSelect', handleMenuItemSelect)
  return () => menuItem.removeEventListener('menu.itemSelect', handleMenuItemSelect)
}

function LocalFontFamilyItem({ id, value }: { id: string, value: string }) {
  const commandService = useDependency(ICommandService)
  const fontService = useDependency(IFontService)
  const layoutService = useDependency(ILayoutService)
  const localeService = useDependency(LocaleService)
  const fonts = useObservable<IFontConfig[]>(fontService.fonts$, [], true)
  const [search, setSearch] = useState('')
  const pickerRef = useRef<HTMLDivElement>(null)
  const searchInteractionRef = useRef(false)
  const deferredSearch = useDeferredValue(search)
  const visibleFonts = useMemo(() => filterLocalFonts(fonts, deferredSearch), [deferredSearch, fonts])

  useLayoutEffect(() => {
    const menuItem = pickerRef.current?.closest('[role="menuitem"]')
    if (!menuItem)
      return

    return bindFontMenuSelectionGuard(menuItem, () => searchInteractionRef.current)
  }, [])

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

  function handlePickerPointerDownCapture(event: ReactPointerEvent<HTMLDivElement>): void {
    const target = event.target as HTMLElement
    searchInteractionRef.current = Boolean(target.closest?.('.univer-plus-font-picker__search-wrap'))
  }

  return (
    <div
      ref={pickerRef}
      className="univer-plus-font-picker"
      onPointerDownCapture={handlePickerPointerDownCapture}
    >
      <div className="univer-plus-font-picker__search-wrap">
        <input
          autoFocus
          aria-label={localeService.t('univerPlus.searchFonts')}
          className="univer-plus-font-picker__search"
          onChange={event => setSearch(event.target.value)}
          onClick={keepFontSearchOpen}
          onKeyDown={handleSearchKeyDown}
          onMouseDown={keepFontSearchOpen}
          onMouseUp={keepFontSearchOpen}
          onPointerDown={keepFontSearchOpen}
          onPointerMove={keepFontSearchOpen}
          onPointerUp={keepFontSearchOpen}
          placeholder={localeService.t('univerPlus.searchFonts')}
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
            {localeService.t(search.trim() ? 'univerPlus.noMatchingFonts' : 'univerPlus.noLocalFonts')}
          </li>
        )}
      </ul>
    </div>
  )
}

function quoteFontFamily(value: string): string {
  return `"${value.replace(/["\\]/g, '\\$&')}"`
}
