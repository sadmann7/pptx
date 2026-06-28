/**
 * Placeholder inheritance resolution.
 *
 * OOXML inheritance chain (bottom wins):
 *   Theme → SlideMaster → SlideLayout → Slide
 *
 * For each element on a slide that has a `placeholder` field, this module
 * fills in any missing geometry, styles, and body properties from the
 * layout and master templates.
 *
 * Rule: slide value takes priority. If missing, use layout. If missing, use master.
 */

import type {
  BodyProperties,
  Fill,
  ParagraphStyle,
  Position,
  Size,
  Slide,
  SlideElement,
  TextShape,
  GeometricShape,
  ImageShape,
  ConnectorShape,
} from '../types'
import type { SlideLayoutModel, SlideMasterModel, PlaceholderTemplate } from './models'
import { placeholderKey } from './models'

export interface InheritanceContext {
  layout: SlideLayoutModel
  master: SlideMasterModel
}

/**
 * Resolve all placeholder elements on a slide against its layout and master.
 * Returns a new Slide with inherited geometry and styles applied.
 */
export function resolveSlideInheritance(slide: Slide, ctx: InheritanceContext): Slide {
  const elements = slide.elements.map((el) => resolveElement(el, ctx))

  // Background: slide bg → layout bg → master bg
  const background =
    slide.background ??
    ctx.layout.background ??
    ctx.master.background

  return { ...slide, elements, background }
}

// ─── Per-element resolution ───────────────────────────────────────────────────

function resolveElement(el: SlideElement, ctx: InheritanceContext): SlideElement {
  if (!el.placeholder) return el

  const key = placeholderKey(el.placeholder.type, el.placeholder.idx)
  const layoutTpl = ctx.layout.placeholders.get(key)
  const masterTpl = ctx.master.placeholders.get(key)

  if (!layoutTpl && !masterTpl) return el

  return mergeElement(el, layoutTpl, masterTpl, ctx)
}

function mergeElement(
  el: SlideElement,
  layout: PlaceholderTemplate | undefined,
  master: PlaceholderTemplate | undefined,
  ctx: InheritanceContext,
): SlideElement {
  // Geometry: prefer slide > layout > master
  const position = hasPosition(el.position) ? el.position : (layout?.position ?? master?.position ?? el.position)
  const size = hasSize(el.size) ? el.size : (layout?.size ?? master?.size ?? el.size)
  const transform = el.transform ?? layout?.transform ?? master?.transform

  // Fill: prefer slide > layout > master
  const fill = getElementFill(el) ?? layout?.fill ?? master?.fill

  const base = { ...el, position, size, transform }

  if (el.type === 'text') {
    return mergeTextShape(el, base as TextShape, layout, master, ctx, fill)
  }
  if (el.type === 'shape') {
    return mergeGeometricShape(el, base as GeometricShape, layout, master, fill)
  }
  if (el.type === 'image') {
    return { ...(base as ImageShape), ...(fill ? { fill } : {}) }
  }
  if (el.type === 'connector') {
    return { ...(base as ConnectorShape), ...(fill ? { fill } : {}) }
  }
  if (el.type === 'group') {
    // Group geometry is resolved, children are handled recursively by the slide resolver
    return base
  }

  return base
}

function mergeTextShape(
  el: TextShape,
  base: TextShape,
  layout: PlaceholderTemplate | undefined,
  master: PlaceholderTemplate | undefined,
  ctx: InheritanceContext,
  fill: Fill | undefined,
): TextShape {
  // Body properties: slide > layout > master
  const properties = mergeBodyProperties(el.properties, layout?.bodyProperties, master?.bodyProperties)

  // Merge paragraph-level styles into each paragraph's style
  const mergedLevelStyles = mergeLevelStyleMaps(
    ctx.master.bodyLevelStyles,
    master?.levelStyles ?? new Map(),
    layout?.levelStyles ?? new Map(),
  )

  const paragraphs = el.paragraphs.map((p) => {
    const level = p.style.level ?? 0
    const inheritedStyle = mergedLevelStyles.get(level) ?? mergedLevelStyles.get(0)
    if (!inheritedStyle) return p

    return {
      ...p,
      style: mergeParaStyle(inheritedStyle, p.style),
    }
  })

  return {
    ...base,
    properties,
    paragraphs,
    ...(fill ? { fill } : {}),
  }
}

function mergeGeometricShape(
  el: GeometricShape,
  base: GeometricShape,
  layout: PlaceholderTemplate | undefined,
  master: PlaceholderTemplate | undefined,
  fill: Fill | undefined,
): GeometricShape {
  if (!el.body) return { ...base, ...(fill ? { fill } : {}) }

  const bodyProperties = mergeBodyProperties(
    el.body.properties,
    layout?.bodyProperties,
    master?.bodyProperties,
  )

  return {
    ...base,
    ...(fill ? { fill } : {}),
    body: { ...el.body, properties: bodyProperties },
  }
}

// ─── Style merging helpers ────────────────────────────────────────────────────

/**
 * Merge body properties, with slide taking priority over layout over master.
 * Only fills in missing values from parent.
 */
function mergeBodyProperties(
  slide: BodyProperties | undefined,
  layout: BodyProperties | undefined,
  master: BodyProperties | undefined,
): BodyProperties {
  return {
    ...master,
    ...layout,
    ...slide,
  }
}

/**
 * Merge a paragraph style — `child` takes priority over `parent`.
 * Only unset properties are inherited.
 */
function mergeParaStyle(parent: ParagraphStyle, child: ParagraphStyle): ParagraphStyle {
  return {
    alignment: child.alignment ?? parent.alignment,
    level: child.level ?? parent.level,
    indent: child.indent ?? parent.indent,
    marginLeft: child.marginLeft ?? parent.marginLeft,
    marginRight: child.marginRight ?? parent.marginRight,
    spaceBefore: child.spaceBefore ?? parent.spaceBefore,
    spaceAfter: child.spaceAfter ?? parent.spaceAfter,
    lineSpacing: child.lineSpacing ?? parent.lineSpacing,
    bullet: child.bullet ?? parent.bullet,
    defaultRunStyle: child.defaultRunStyle ?? parent.defaultRunStyle,
  }
}

/**
 * Merge multiple level-style maps, where later maps take priority.
 * body (master txStyles) < master placeholder lstStyle < layout placeholder lstStyle
 */
function mergeLevelStyleMaps(
  ...maps: Array<Map<number, ParagraphStyle>>
): Map<number, ParagraphStyle> {
  const result = new Map<number, ParagraphStyle>()
  for (const map of maps) {
    for (const [level, style] of map) {
      const existing = result.get(level)
      result.set(level, existing ? mergeParaStyle(existing, style) : style)
    }
  }
  return result
}

// ─── Utility ─────────────────────────────────────────────────────────────────

function hasPosition(pos: Position): boolean {
  return pos.x !== 0 || pos.y !== 0
}

function hasSize(size: Size): boolean {
  return size.width !== 0 || size.height !== 0
}

function getElementFill(el: SlideElement): Fill | undefined {
  if (el.type === 'group' || el.type === 'table' || el.type === 'chart') return undefined
  return el.fill
}
