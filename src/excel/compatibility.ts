import JSZip from 'jszip'

const SPREADSHEETML_NAMESPACE = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

export async function normalizeXlsxForExcelJs(buffer: ArrayBuffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(buffer)
  let changed = false

  for (const entry of Object.values(zip.files)) {
    if (entry.dir || (!entry.name.toLowerCase().endsWith('.xml') && !entry.name.toLowerCase().endsWith('.rels')))
      continue

    const source = await entry.async('text')
    const normalized = entry.name.toLowerCase().endsWith('.rels')
      ? normalizeRelationshipTargets(source, entry.name)
      : normalizeSpreadsheetMlElements(source)

    if (normalized !== source) {
      zip.file(entry.name, normalized)
      changed = true
    }
  }

  return changed ? zip.generateAsync({ type: 'arraybuffer' }) : buffer
}

function normalizeSpreadsheetMlElements(source: string): string {
  const namespacePattern = new RegExp(`\\sxmlns:([A-Za-z_][\\w.-]*)=(["'])${escapeRegExp(SPREADSHEETML_NAMESPACE)}\\2`)
  const defaultNamespacePattern = new RegExp(`\\sxmlns=(["'])${escapeRegExp(SPREADSHEETML_NAMESPACE)}\\1`)
  let xml = source
  let match = namespacePattern.exec(xml)

  while (match) {
    const prefix = escapeRegExp(match[1])
    xml = xml.replace(new RegExp(`(<\\/?)${prefix}:([A-Za-z_][\\w.-]*)`, 'g'), '$1$2')
    xml = xml.replace(match[0], defaultNamespacePattern.test(xml) ? '' : ` xmlns=${match[2]}${SPREADSHEETML_NAMESPACE}${match[2]}`)
    match = namespacePattern.exec(xml)
  }

  return xml
}

function normalizeRelationshipTargets(source: string, relationshipPath: string): string {
  const sourceDirectory = relationshipSourceDirectory(relationshipPath)
  return source.replace(/<(?:[A-Za-z_][\w.-]*:)?Relationship\b[^>]*>/g, (relationship) => {
    if (/\bTargetMode\s*=\s*(["'])External\1/i.test(relationship))
      return relationship

    return relationship.replace(/\bTarget\s*=\s*(["'])(\/(?!\/)[^"']*)\1/i, (_, quote: string, target: string) => {
      return `Target=${quote}${relativePackagePath(sourceDirectory, target)}${quote}`
    })
  })
}

function relationshipSourceDirectory(path: string): string {
  const marker = '/_rels/'
  const index = path.lastIndexOf(marker)
  return index < 0 ? '' : path.slice(0, index)
}

function relativePackagePath(fromDirectory: string, absoluteTarget: string): string {
  const from = fromDirectory.split('/').filter(Boolean)
  const target = absoluteTarget.split('/').filter(Boolean)

  while (from.length > 0 && target.length > 0 && from[0] === target[0]) {
    from.shift()
    target.shift()
  }

  return [...from.map(() => '..'), ...target].join('/')
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}
