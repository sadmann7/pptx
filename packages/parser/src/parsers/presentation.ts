import type { Presentation, SlideSize, Theme } from '../types.js'
import { parseTheme } from './theme.js'
import { parseSlide } from './slide.js'
import type { ParseOptions, PresentationInput } from '../types.js'
import type { PptxZip } from '../zip.js'
import { loadZip, loadRels, readXml } from '../zip.js'
import { attr, attrNum, get, toArray } from '../xml.js'
import { emuToPoints } from '../emu.js'
import { loadLayoutModel, loadMasterModel } from '../resolve/models.js'
import { resolveSlideInheritance } from '../resolve/inheritance.js'

const PRESENTATION_PATH = 'ppt/presentation.xml'

/**
 * Top-level presentation parser. Orchestrates ZIP loading, theme parsing,
 * master/layout pre-loading, per-slide parsing, and inheritance resolution.
 */
export async function parsePresentation(
  input: PresentationInput,
  options: ParseOptions = {},
): Promise<Presentation> {
  const { onProgress, skipNotes = false, skipImages = false } = options

  const zip = await loadZip(input as ArrayBuffer | Uint8Array | Blob)
  const presXml = await readXml(zip, PRESENTATION_PATH)
  const presRels = await loadRels(zip, PRESENTATION_PATH)

  // ── Slide size ────────────────────────────────────────────────────────────
  const sldSz = get(presXml, 'p:presentation', 'p:sldSz') as
    | Record<string, unknown>
    | undefined

  const slideSize: SlideSize = sldSz
    ? {
        width: emuToPoints(attrNum(sldSz, 'cx') ?? 9144000),
        height: emuToPoints(attrNum(sldSz, 'cy') ?? 6858000),
      }
    : { width: 720, height: 540 } // 10×7.5in fallback

  // ── Theme ─────────────────────────────────────────────────────────────────
  const themeRel = [...presRels.values()].find((r) => r.type.includes('theme'))
  const theme = await parseThemeFromRel(zip, themeRel?.target)

  // ── Slide list ────────────────────────────────────────────────────────────
  const sldIdLst = get(presXml, 'p:presentation', 'p:sldIdLst', 'p:sldId')
  const slideRefs = toArray(sldIdLst as unknown[])
  const total = slideRefs.length

  const slides = await Promise.all(
    slideRefs.map(async (ref, index) => {
      const rId = attr(ref, 'r:id') ?? attr(ref, 'id') ?? ''
      const rel = presRels.get(rId)
      const slidePath = rel?.target ?? `ppt/slides/slide${index + 1}.xml`

      // Parse raw slide
      const rawSlide = await parseSlide(zip, slidePath, index, rId, skipImages, skipNotes)

      // Load layout + master for this slide and resolve inheritance
      const slideRels = await loadRels(zip, slidePath)
      const layoutRel = [...slideRels.values()].find((r) => r.type.includes('slideLayout'))
      const layoutPath = layoutRel?.target

      if (layoutPath) {
        const layout = await loadLayoutModel(zip, layoutPath)
        const master = await loadMasterModel(zip, layout.masterPath)
        const resolved = resolveSlideInheritance(rawSlide, { layout, master })
        onProgress?.(index + 1, total)
        return resolved
      }

      onProgress?.(index + 1, total)
      return rawSlide
    }),
  )

  return { slideSize, theme, slides }
}

async function parseThemeFromRel(
  zip: PptxZip,
  themePath: string | undefined,
): Promise<Theme> {
  const path = themePath ?? 'ppt/theme/theme1.xml'
  const themeXml = await readXml(zip, path)
  return parseTheme(themeXml)
}
