export class SaveCoordinator {
  private tail: Promise<void> = Promise.resolve()
  private disposed = false

  run<T>(task: () => Promise<T>): Promise<T> {
    if (this.disposed)
      return Promise.reject(new Error('The save coordinator has been disposed'))

    const result = this.tail.then(task)
    this.tail = result.then(() => undefined, () => undefined)
    return result
  }

  async flush(): Promise<void> {
    await this.tail
  }

  dispose(): void {
    this.disposed = true
  }
}
