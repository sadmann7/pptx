import type { Background, Slide } from '../types.js'
import { parseBackground } from './fill.js'
import { parseSpTree } from './shape.js'
import type { PptxZip } from '../zip.js'
import { loadRels, readXml } from '../zip.js'
import { get, toArray } from '../xml.js'

/**
 * Parse a single slide XML file into a Slide AST node.
 */
export async function parseSlide(
  zip: PptxZip,
  slidePath: string,
  index: number,
  rId: string,
  skipImages: boolean,
  skipNotes: boolean,
): Promise<Slide> {
  const slideXml = await readXml(zip, slidePath)
  const rels = await loadRels(zip, slidePath)

  const sld = get(slideXml, 'p:sld') as Record<string, unknown> | undefined
  const cSld = get(sld, 'p:cSld') as Record<string, unknown> | undefined

  // Background
  let background: Background | undefined
  const bg = get(cSld, 'p:bg')
  if (bg) {
    const fill = parseBackground(bg)
    if (fill) background = { fill }
  }

  // Shape tree
  const spTree = get(cSld, 'p:spTree') as Record<string, unknown> | undefined
  const elements = spTree
    ? await parseSpTree(spTree, rels, zip, slidePath, skipImages)
    : []

  // Notes
  let notes: string | undefined
  if (!skipNotes) {
    const notesRel = [...rels.values()].find((r) =>
      r.type.includes('notesSlide')
    )
    if (notesRel) {
      notes = await parseNotes(zip, notesRel.target)
    }
  }

  return {
    index,
    rId,
    path: slidePath,
    elements,
    background,
    notes,
  }
}

async function parseNotes(zip: PptxZip, notesPath: string): Promise<string> {
  const xml = await readXml(zip, notesPath)
  const cSld = get(xml, 'p:notes', 'p:cSld') as Record<string, unknown> | undefined
  const spTree = get(cSld, 'p:spTree') as Record<string, unknown> | undefined
  if (!spTree) return ''

  const spNodes = toArray(spTree['p:sp'] as unknown[])
  const textParts: string[] = []

  for (const sp of spNodes) {
    const txBody = get(sp as Record<string, unknown>, 'p:txBody') as Record<string, unknown> | undefined
    if (!txBody) continue
    const pNodes = toArray(txBody['a:p'] as unknown[])
    for (const p of pNodes) {
      const pN = p as Record<string, unknown>
      const rNodes = toArray(pN['a:r'] as unknown[])
      const line = rNodes
        .map((r) => {
          const t = (r as Record<string, unknown>)['a:t']
          return typeof t === 'string' ? t : ''
        })
        .join('')
      if (line) textParts.push(line)
    }
  }

  return textParts.join('\n')
}
