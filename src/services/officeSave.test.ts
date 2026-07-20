import { describe, expect, it } from 'vitest'
import { hasUnsavedSnapshot, isOfficeSaveShortcut } from './officeSave'

describe('office save helpers', () => {
  it('recognizes Ctrl/Cmd+S without treating Alt+S as a save shortcut', () => {
    expect(isOfficeSaveShortcut({ ctrlKey: true, metaKey: false, altKey: false, key: 's' })).toBe(true)
    expect(isOfficeSaveShortcut({ ctrlKey: false, metaKey: true, altKey: false, key: 'S' })).toBe(true)
    expect(isOfficeSaveShortcut({ ctrlKey: true, metaKey: false, altKey: true, key: 's' })).toBe(false)
  })

  it('uses the current document snapshot instead of the command type as the dirty signal', () => {
    expect(hasUnsavedSnapshot({ value: 'after' }, JSON.stringify({ value: 'before' }))).toBe(true)
    expect(hasUnsavedSnapshot({ value: 'before' }, JSON.stringify({ value: 'before' }))).toBe(false)
    expect(hasUnsavedSnapshot(undefined, '')).toBe(false)
  })
})
