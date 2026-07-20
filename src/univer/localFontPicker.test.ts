import { describe, expect, it, vi } from 'vitest'
import { keepFontSearchOpen } from './localFontPicker'

describe('local font picker interaction', () => {
  it('keeps pointer and click events inside the font search field', () => {
    const stopPropagation = vi.fn()

    keepFontSearchOpen({ stopPropagation })

    expect(stopPropagation).toHaveBeenCalledOnce()
  })
})
