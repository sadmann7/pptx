import { XMLParser } from 'fast-xml-parser'

/**
 * Elements that must always be treated as arrays even when there's only
 * one instance in the XML.
 */
const ALWAYS_ARRAY = new Set([
  // Presentation
  'p:sldIdLst',
  'p:sldId',
  'p:sldMasterIdLst',
  'p:sldMasterId',
  'p:sldLayoutIdLst',
  'p:sldLayoutId',
  // Slide content
  'p:sp',
  'p:pic',
  'p:graphicFrame',
  'p:grpSp',
  'p:cxnSp',
  'p:spTree',
  // Text
  'a:p',
  'a:r',
  'a:br',
  'a:fld',
  // Table
  'a:tr',
  'a:tc',
  'a:gridCol',
  'a:tblStyleLst',
  // Relationships
  'Relationship',
  // Content types
  'Override',
  'Default',
  // Theme
  'a:fmtScheme',
  // Gradient stops
  'a:gs',
  // Paragraph spacing
  'a:tab',
  // List styles
  'a:lvl1pPr',
  'a:lvl2pPr',
  'a:lvl3pPr',
  'a:lvl4pPr',
  'a:lvl5pPr',
  'a:lvl6pPr',
  'a:lvl7pPr',
  'a:lvl8pPr',
  'a:lvl9pPr',
])

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  allowBooleanAttributes: true,
  parseAttributeValue: false,
  parseTagValue: true,
  trimValues: true,
  isArray: (name) => ALWAYS_ARRAY.has(name),
})

export function parseXml(xmlString: string): Record<string, unknown> {
  return parser.parse(xmlString) as Record<string, unknown>
}

/** Safe attribute getter — returns string or undefined */
export function attr(node: unknown, name: string): string | undefined {
  if (node === null || typeof node !== 'object') return undefined
  const val = (node as Record<string, unknown>)[`@_${name}`]
  if (val === undefined || val === null) return undefined
  return String(val)
}

/** Safe numeric attribute getter */
export function attrNum(node: unknown, name: string): number | undefined {
  const s = attr(node, name)
  if (s === undefined) return undefined
  const n = Number(s)
  return Number.isNaN(n) ? undefined : n
}

/** Safe boolean attribute getter — treats '1', 'true' as true */
export function attrBool(node: unknown, name: string): boolean | undefined {
  const s = attr(node, name)
  if (s === undefined) return undefined
  return s === '1' || s === 'true'
}

/**
 * Get a child node by dotted path, e.g. get(root, 'p:sld', 'p:cSld', 'p:spTree')
 * Returns undefined if any segment is missing.
 */
export function get(root: unknown, ...path: string[]): unknown {
  let cur: unknown = root
  for (const key of path) {
    if (cur === null || typeof cur !== 'object') return undefined
    cur = (cur as Record<string, unknown>)[key]
  }
  return cur
}

/**
 * Returns a node as an array. If it's already an array, return it.
 * If it's a non-null value, wrap it. If undefined/null, return [].
 */
export function toArray<T>(val: T | T[] | undefined | null): T[] {
  if (val === undefined || val === null) return []
  if (Array.isArray(val)) return val
  return [val]
}

/** Get text content of a node (handles both string leaf and object with #text) */
export function textContent(node: unknown): string {
  if (typeof node === 'string') return node
  if (typeof node === 'number') return String(node)
  if (node === null || node === undefined) return ''
  if (typeof node === 'object') {
    const t = (node as Record<string, unknown>)['#text']
    if (t !== undefined) return String(t)
  }
  return ''
}
