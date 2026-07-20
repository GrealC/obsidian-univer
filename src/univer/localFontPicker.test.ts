import { describe, expect, it, vi } from 'vitest'
import { bindFontMenuSelectionGuard, keepFontSearchOpen, preventFontMenuSelection } from './localFontPicker'

describe('local font picker interaction', () => {
  it('lets Radix observe pointerdown but blocks pointerup from synthesizing an outer click', () => {
    const stopPropagation = vi.fn()

    keepFontSearchOpen({ type: 'pointerdown', stopPropagation })
    expect(stopPropagation).not.toHaveBeenCalled()

    keepFontSearchOpen({ type: 'pointerup', stopPropagation })
    expect(stopPropagation).toHaveBeenCalledOnce()
  })

  it('blocks the final click from selecting the outer menu item', () => {
    const stopPropagation = vi.fn()

    keepFontSearchOpen({ type: 'click', stopPropagation })

    expect(stopPropagation).toHaveBeenCalledOnce()
  })

  it('cancels the Radix menu selection only while the search field is active', () => {
    const preventDefault = vi.fn()

    preventFontMenuSelection({ preventDefault }, true)
    expect(preventDefault).toHaveBeenCalledOnce()

    preventDefault.mockClear()
    preventFontMenuSelection({ preventDefault }, false)
    expect(preventDefault).not.toHaveBeenCalled()
  })

  it('keeps a cancelable Radix menu event open without blocking font option selection', () => {
    const menuItem = new EventTarget()
    let searchInteraction = true
    const dispose = bindFontMenuSelectionGuard(menuItem, () => searchInteraction)

    const searchSelect = new Event('menu.itemSelect', { cancelable: true })
    expect(menuItem.dispatchEvent(searchSelect)).toBe(false)
    expect(searchSelect.defaultPrevented).toBe(true)

    searchInteraction = false
    const fontSelect = new Event('menu.itemSelect', { cancelable: true })
    expect(menuItem.dispatchEvent(fontSelect)).toBe(true)
    expect(fontSelect.defaultPrevented).toBe(false)

    dispose()
  })
})
