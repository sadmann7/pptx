import JSZip from 'jszip'
import { parseXml, get, toArray, attr } from './xml.js'

export type PptxZip = JSZip

/**
 * Load a PPTX (which is a ZIP) from any supported input.
 */
export async function loadZip(input: ArrayBuffer | Uint8Array | Blob): Promise<PptxZip> {
  return JSZip.loadAsync(input)
}

/**
 * Read a file from the ZIP as a UTF-8 string.
 * Returns empty string if the file doesn't exist.
 */
export async function readString(zip: PptxZip, path: string): Promise<string> {
  const normalized = normalizePath(path)
  const file = zip.file(normalized)
  if (!file) return ''
  return file.async('string')
}

/**
 * Read a file from the ZIP as a Uint8Array.
 * Returns null if the file doesn't exist.
 */
export async function readBytes(zip: PptxZip, path: string): Promise<Uint8Array | null> {
  const normalized = normalizePath(path)
  const file = zip.file(normalized)
  if (!file) return null
  return file.async('uint8array')
}

/**
 * Read a file and parse it as XML, returning the parsed object.
 * Returns an empty object if the file doesn't exist.
 */
export async function readXml(zip: PptxZip, path: string): Promise<Record<string, unknown>> {
  const str = await readString(zip, path)
  if (!str) return {}
  return parseXml(str)
}

/**
 * Read a media file and return a blob URL (browser) or base64 data URI (Node).
 */
export async function readMediaAsUrl(zip: PptxZip, path: string): Promise<{ src: string; mimeType: string }> {
  const normalized = normalizePath(path)
  const file = zip.file(normalized)
  if (!file) return { src: '', mimeType: '' }

  const ext = normalized.split('.').pop()?.toLowerCase() ?? ''
  const mimeType = MIME_TYPES[ext] ?? 'application/octet-stream'

  const arrayBuffer = await file.async('arraybuffer')

  // Use Blob URL in browser environments, fall back to base64 in Node
  if (typeof Blob !== 'undefined' && typeof URL !== 'undefined') {
    const blob = new Blob([arrayBuffer], { type: mimeType })
    return { src: URL.createObjectURL(blob), mimeType }
  }

  // Node fallback: base64 data URI
  const base64 = Buffer.from(arrayBuffer).toString('base64')
  return { src: `data:${mimeType};base64,${base64}`, mimeType }
}

// ─── Relationship (.rels) resolution ─────────────────────────────────────────

export interface Relationship {
  id: string
  type: string
  target: string
}

const RELS_CACHE = new WeakMap<PptxZip, Map<string, Map<string, Relationship>>>()

/**
 * Parse the .rels file associated with a given XML file path.
 *
 * e.g. for 'ppt/slides/slide1.xml', loads 'ppt/slides/_rels/slide1.xml.rels'
 */
export async function loadRels(zip: PptxZip, filePath: string): Promise<Map<string, Relationship>> {
  let zipCache = RELS_CACHE.get(zip)
  if (!zipCache) {
    zipCache = new Map()
    RELS_CACHE.set(zip, zipCache)
  }

  if (zipCache.has(filePath)) {
    return zipCache.get(filePath)!
  }

  const relsPath = toRelsPath(filePath)
  const xml = await readXml(zip, relsPath)

  const map = new Map<string, Relationship>()

  const relationships = toArray(
    get(xml, 'Relationships', 'Relationship') as unknown[]
  )

  for (const rel of relationships) {
    const id = attr(rel, 'Id') ?? ''
    const type = attr(rel, 'Type') ?? ''
    const rawTarget = attr(rel, 'Target') ?? ''

    const targetMode = attr(rel, 'TargetMode')
    const target = targetMode === 'External'
      ? rawTarget
      : resolveRelTarget(filePath, rawTarget)

    if (id) {
      map.set(id, { id, type, target })
    }
  }

  zipCache.set(filePath, map)
  return map
}

/** Resolve a relationship target path relative to a source file */
function resolveRelTarget(sourcePath: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)

  const sourceDir = sourcePath.includes('/')
    ? sourcePath.substring(0, sourcePath.lastIndexOf('/'))
    : ''

  return joinPath(sourceDir, target)
}

function joinPath(dir: string, file: string): string {
  const parts = `${dir}/${file}`.split('/')
  const result: string[] = []
  for (const part of parts) {
    if (part === '..') {
      result.pop()
    } else if (part !== '.' && part !== '') {
      result.push(part)
    }
  }
  return result.join('/')
}

function toRelsPath(filePath: string): string {
  const lastSlash = filePath.lastIndexOf('/')
  const dir = lastSlash >= 0 ? filePath.substring(0, lastSlash) : ''
  const filename = lastSlash >= 0 ? filePath.substring(lastSlash + 1) : filePath
  return dir ? `${dir}/_rels/${filename}.rels` : `_rels/${filename}.rels`
}

/** Normalize Windows-style backslashes and strip leading slash */
function normalizePath(path: string): string {
  return path.replace(/\\/g, '/').replace(/^\//, '')
}

// ─── MIME type map ────────────────────────────────────────────────────────────

const MIME_TYPES: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  tiff: 'image/tiff',
  tif: 'image/tiff',
  webp: 'image/webp',
  svg: 'image/svg+xml',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
  mp4: 'video/mp4',
  mov: 'video/quicktime',
  avi: 'video/x-msvideo',
}

// ─── Content types ────────────────────────────────────────────────────────────

export interface ContentTypeEntry {
  partName: string
  contentType: string
}

export async function loadContentTypes(zip: PptxZip): Promise<ContentTypeEntry[]> {
  const xml = await readXml(zip, '[Content_Types].xml')
  const overrides = toArray(get(xml, 'Types', 'Override') as unknown[])
  return overrides.map((o) => ({
    partName: (attr(o, 'PartName') ?? '').replace(/^\//, ''),
    contentType: attr(o, 'ContentType') ?? '',
  }))
}
