export class SaveCoordinator {
  private tail: Promise<void> = Promise.resolve()
  private readonly inFlight = new Map<unknown, Map<string, Promise<unknown>>>()
  private disposed = false

  run<T>(task: () => Promise<T>): Promise<T> {
    if (this.disposed)
      return Promise.reject(new Error('The save coordinator has been disposed'))

    const result = this.tail.then(task)
    this.tail = result.then(() => undefined, () => undefined)
    return result
  }

  runDeduplicated<T>(key: unknown, fingerprint: string, task: () => Promise<T>): Promise<T> {
    const existing = this.inFlight.get(key)?.get(fingerprint)
    if (existing)
      return existing as Promise<T>

    const promise = this.run(task)
    const keyedSaves = this.inFlight.get(key) ?? new Map<string, Promise<unknown>>()
    keyedSaves.set(fingerprint, promise)
    this.inFlight.set(key, keyedSaves)
    void promise.finally(() => {
      const currentSaves = this.inFlight.get(key)
      if (currentSaves?.get(fingerprint) !== promise)
        return
      currentSaves.delete(fingerprint)
      if (currentSaves.size === 0)
        this.inFlight.delete(key)
    }).catch(() => undefined)
    return promise
  }

  async flush(): Promise<void> {
    await this.tail
  }

  dispose(): void {
    this.disposed = true
    this.inFlight.clear()
  }
}
