export function selectObsoleteBackups(files: readonly string[], safeName: string, keep: number): string[] {
  const suffix = `-${safeName}`
  return files
    .filter(path => path.endsWith(suffix))
    .sort((left, right) => right.localeCompare(left))
    .slice(Math.max(0, keep))
}
