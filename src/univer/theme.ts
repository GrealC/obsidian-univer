import type { FUniver } from '@univerjs/core/facade'

export function observeTheme(univerAPI: FUniver): MutationObserver {
  const syncTheme = () => univerAPI.toggleDarkMode(document.body.classList.contains('theme-dark'))
  syncTheme()
  const observer = new MutationObserver(syncTheme)
  observer.observe(document.body, { attributes: true, attributeFilter: ['class'] })
  return observer
}
