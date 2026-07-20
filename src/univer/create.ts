import type { IUniverConfig, Plugin, PluginCtor } from '@univerjs/core'
import type { FUniver } from '@univerjs/core/facade'
import { LogLevel, Univer } from '@univerjs/core'
import { FUniver as FUniverClass } from '@univerjs/core/facade'

type PluginEntry = PluginCtor<Plugin> | [PluginCtor<Plugin>, unknown]

export interface UniverPreset {
  plugins: PluginEntry[]
}

export interface UniverRuntime {
  univer: Univer
  univerAPI: FUniver
}

interface CreateUniverOptions extends Partial<IUniverConfig> {
  presets: UniverPreset[]
  plugins?: PluginEntry[]
}

export function createUniver(options: CreateUniverOptions): UniverRuntime {
  const { presets, plugins = [], ...config } = options
  const univer = new Univer({
    logLevel: LogLevel.WARN,
    ...config,
  })
  const registered = new Map<string, { plugin: PluginCtor<Plugin>, options?: unknown }>()

  for (const preset of presets) {
    for (const entry of preset.plugins) {
      const [plugin, pluginOptions] = Array.isArray(entry) ? entry : [entry, undefined]
      registered.set(plugin.pluginName, { plugin, options: pluginOptions })
    }
  }

  for (const entry of plugins) {
    const [plugin, pluginOptions] = Array.isArray(entry) ? entry : [entry, undefined]
    if (registered.has(plugin.pluginName))
      throw new Error(`Univer plugin ${plugin.pluginName} is registered more than once`)

    registered.set(plugin.pluginName, { plugin, options: pluginOptions })
  }

  for (const { plugin, options: pluginOptions } of registered.values())
    univer.registerPlugin(plugin, pluginOptions)

  return {
    univer,
    univerAPI: FUniverClass.newAPI(univer),
  }
}
