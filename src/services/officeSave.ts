export type OfficeSaveReason = 'auto' | 'manual' | 'close'

export interface SaveShortcutEvent {
  altKey: boolean
  ctrlKey: boolean
  key: string
  metaKey: boolean
}

/** Returns true for the platform save shortcut, including shifted variants. */
export function isOfficeSaveShortcut(event: SaveShortcutEvent): boolean {
  return (event.ctrlKey || event.metaKey)
    && !event.altKey
    && event.key.toLowerCase() === 's'
}

/** Avoid relying on a particular Univer command type to determine whether data changed. */
export function hasUnsavedSnapshot(snapshot: unknown, baseline: string): boolean {
  return snapshot !== undefined && JSON.stringify(snapshot) !== baseline
}
