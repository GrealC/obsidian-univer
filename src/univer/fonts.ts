import type { IFontConfig } from '@univerjs/preset-sheets-core'
import type { UniverRuntime } from './create'
import { execFile } from 'node:child_process'
import process from 'node:process'
import { promisify } from 'node:util'

const LOCAL_FONT_LIMIT = 2048
const LOCALIZED_FONT_LABELS: Readonly<Record<string, string>> = {
  'alibaba puhuiti': '阿里巴巴普惠体',
  'dengxian': '等线',
  'dfkai-sb': '标楷体',
  'fangsong': '仿宋',
  'fzshuti': '方正舒体',
  'fzyaoti': '方正姚体',
  'kaiti': '楷体',
  'lisu': '隶书',
  'lxgw wenkai': '霞鹜文楷',
  'lxgw wenkai mono': '霞鹜文楷等宽',
  'microsoft jhenghei': '微软正黑体',
  'microsoft jhenghei ui': '微软正黑体 UI',
  'microsoft yahei': '微软雅黑',
  'microsoft yahei light': '微软雅黑 Light',
  'microsoft yahei ui': '微软雅黑 UI',
  'microsoft yahei ui light': '微软雅黑 UI Light',
  'mingliu': '细明体',
  'mingliu_hkscs': '细明体_HKSCS',
  'nsimsun': '新宋体',
  'pmingliu': '新细明体',
  'simhei': '黑体',
  'simsun': '宋体',
  'source han sans sc': '思源黑体',
  'source han serif sc': '思源宋体',
  'stfangsong': '华文仿宋',
  'stheiti': '华文黑体',
  'stkaiti': '华文楷体',
  'stliti': '华文隶书',
  'stsong': '华文宋体',
  'stxingkai': '华文行楷',
  'stxinwei': '华文新魏',
  'youyuan': '幼圆',
}
const WINDOWS_FONT_COMMAND = [
  'Add-Type -AssemblyName System.Drawing',
  '$zhCn = [System.Globalization.CultureInfo]::GetCultureInfo("zh-CN").LCID',
  '[System.Drawing.Text.InstalledFontCollection]::new().Families | ForEach-Object { [PSCustomObject]@{ family = $_.Name; label = $_.GetName($zhCn) } } | ConvertTo-Json -Compress',
].join('; ')
const execFileAsync = promisify(execFile)
const registeredFontRuntimes = new WeakSet<UniverRuntime>()

let localFontsPromise: Promise<IFontConfig[]> | undefined

interface LocalFontData {
  family?: unknown
}

export interface LocalFontFamily {
  family: string
  label?: string
}

interface LocalFontWindow extends Window {
  queryLocalFonts?: () => Promise<ReadonlyArray<LocalFontData>>
}

/**
 * Returns the actual font families available to the current desktop client.
 * The Electron API is preferred; Windows falls back to the OS font collection.
 */
export function discoverLocalFonts(): Promise<IFontConfig[]> {
  localFontsPromise ??= loadLocalFonts()
  return localFontsPromise
}

/** Adds the cached local-font list to one Univer runtime exactly once. */
export async function registerLocalFonts(runtime: UniverRuntime): Promise<void> {
  if (registeredFontRuntimes.has(runtime))
    return

  registeredFontRuntimes.add(runtime)
  const fonts = await discoverLocalFonts()
  if (fonts.length === 0)
    return

  try {
    runtime.univerAPI.addFonts(fonts)
  }
  catch {
    // The view may have been disposed while the asynchronous discovery ran.
    registeredFontRuntimes.delete(runtime)
  }
}

export function createLocalFontConfigs(fontFamilies: Iterable<unknown>, limit = LOCAL_FONT_LIMIT): IFontConfig[] {
  const configs = new Map<string, { value: string, label: string }>()
  const fontKeys = new Map<string, string>()

  for (const fontFamily of fontFamilies) {
    const font = toLocalFontFamily(fontFamily)
    if (!font)
      continue

    const value = font.family.trim()
    const label = getSafeLabel(font.label, value)
    if (!isSafeFontName(value) || !isSafeFontName(label))
      continue

    const valueKey = normalizeFontName(value)
    const labelKey = normalizeFontName(label)
    const existingKey = fontKeys.get(valueKey) ?? fontKeys.get(labelKey)
    if (existingKey) {
      const existing = configs.get(existingKey)
      if (existing && existing.label === existing.value && label !== value)
        existing.label = label
      fontKeys.set(valueKey, existingKey)
      fontKeys.set(labelKey, existingKey)
      continue
    }

    configs.set(valueKey, { value, label })
    fontKeys.set(valueKey, valueKey)
    fontKeys.set(labelKey, valueKey)
  }

  const maxFonts = Number.isFinite(limit) ? Math.max(0, Math.floor(limit)) : LOCAL_FONT_LIMIT
  return [...configs.values()]
    .sort((left, right) => left.value.localeCompare(right.value, undefined, { sensitivity: 'base' }))
    .slice(0, maxFonts)
    .map(font => ({ ...font, category: inferCategory(font.value) }))
}

export function parseInstalledFontNames(output: string): string[] {
  return parseInstalledFontFamilies(output).map(font => font.family)
}

export function matchesFontSearch(font: IFontConfig, search: string): boolean {
  const normalizedSearch = normalizeFontName(search.trim())
  return !normalizedSearch
    || normalizeFontName(font.value).includes(normalizedSearch)
    || normalizeFontName(font.label).includes(normalizedSearch)
}

export function parseInstalledFontFamilies(output: string): LocalFontFamily[] {
  const serialized = output.trim()
  if (!serialized)
    return []

  try {
    const result: unknown = JSON.parse(serialized)
    const values = Array.isArray(result) ? result : [result]
    return values.flatMap((value) => {
      const font = toLocalFontFamily(value)
      return font ? [font] : []
    })
  }
  catch {
    // The fallback is optional. An invalid response simply leaves the picker empty.
  }

  return []
}

async function loadLocalFonts(): Promise<IFontConfig[]> {
  const [browserFonts, windowsFonts] = await Promise.all([
    queryBrowserFontFamilies(),
    queryWindowsFontFamilies(),
  ])
  // Windows reports localized Chinese labels while retaining the canonical CSS family name.
  return createLocalFontConfigs([...windowsFonts, ...browserFonts])
}

async function queryBrowserFontFamilies(): Promise<LocalFontFamily[]> {
  if (typeof window === 'undefined')
    return []

  const queryLocalFonts = (window as LocalFontWindow).queryLocalFonts
  if (typeof queryLocalFonts !== 'function')
    return []

  try {
    const fonts = await queryLocalFonts.call(window) as ReadonlyArray<LocalFontData>
    return fonts.flatMap((font) => {
      const family = typeof font.family === 'string' ? font.family : ''
      return family ? [{ family }] : []
    })
  }
  catch {
    return []
  }
}

async function queryWindowsFontFamilies(): Promise<LocalFontFamily[]> {
  if (typeof process === 'undefined' || process.platform !== 'win32')
    return []

  try {
    const { stdout } = await execFileAsync(getPowerShellPath(), [
      '-NoProfile',
      '-NonInteractive',
      '-Command',
      WINDOWS_FONT_COMMAND,
    ], {
      encoding: 'utf8',
      maxBuffer: 1024 * 1024,
      windowsHide: true,
    })
    return parseInstalledFontFamilies(String(stdout))
  }
  catch {
    return []
  }
}

function toLocalFontFamily(value: unknown): LocalFontFamily | undefined {
  if (typeof value === 'string')
    return { family: value }
  if (!value || typeof value !== 'object' || Array.isArray(value))
    return undefined

  const { family, label } = value as LocalFontData & { label?: unknown }
  if (typeof family !== 'string')
    return undefined
  return {
    family,
    label: typeof label === 'string' ? label : undefined,
  }
}

function getSafeLabel(value: string | undefined, fallback: string): string {
  const label = value?.trim()
  if (label && label !== fallback && isSafeFontName(label))
    return label
  return LOCALIZED_FONT_LABELS[normalizeFontName(fallback)] ?? fallback
}

function getPowerShellPath(): string {
  const systemRoot = process.env.SystemRoot ?? process.env.WINDIR
  return systemRoot
    ? `${systemRoot}\\System32\\WindowsPowerShell\\v1.0\\powershell.exe`
    : 'powershell.exe'
}

function isSafeFontName(value: string): boolean {
  return value.length > 0 && value.length <= 128 && !hasControlCharacter(value)
}

function hasControlCharacter(value: string): boolean {
  return [...value].some((character) => {
    const code = character.charCodeAt(0)
    return code < 32 || code === 127
  })
}

function normalizeFontName(value: string): string {
  return value.toLocaleLowerCase()
}

function inferCategory(value: string): IFontConfig['category'] {
  if (/mono|courier|consolas|code/i.test(value))
    return 'monospace'
  if (/serif|times|cambria|georgia|song|ming/i.test(value))
    return 'serif'
  return 'sans-serif'
}
