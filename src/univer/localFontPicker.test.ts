import { describe, expect, it, vi } from 'vitest'
import { keepFontSearchOpen } from './localFontPicker'

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
})
