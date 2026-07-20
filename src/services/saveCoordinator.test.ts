import { describe, expect, it } from 'vitest'
import { SaveCoordinator } from './saveCoordinator'

describe('saveCoordinator', () => {
  it('serializes concurrent save operations', async () => {
    const coordinator = new SaveCoordinator()
    const events: string[] = []
    let releaseFirst: () => void = () => undefined
    const firstGate = new Promise<void>((resolve) => {
      releaseFirst = resolve
    })

    const first = coordinator.run(async () => {
      events.push('first:start')
      await firstGate
      events.push('first:end')
    })
    const second = coordinator.run(async () => {
      events.push('second:start')
      events.push('second:end')
    })

    await Promise.resolve()
    expect(events).toEqual(['first:start'])
    releaseFirst()
    await Promise.all([first, second])
    expect(events).toEqual(['first:start', 'first:end', 'second:start', 'second:end'])
  })

  it('continues after a failed save', async () => {
    const coordinator = new SaveCoordinator()
    await expect(coordinator.run(async () => {
      throw new Error('write failed')
    })).rejects.toThrow('write failed')
    await expect(coordinator.run(async () => 'saved')).resolves.toBe('saved')
  })
})
