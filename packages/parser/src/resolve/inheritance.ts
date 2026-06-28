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
  Paragraph,
  ParagraphStyle,
  Position,
  RunStyle,
  Size,
  Slide,
  SlideElement,
  TextShape,
  GeometricShape,
  ImageShape,
  ConnectorShape,
} from "../types";
import type { SlideLayoutModel, SlideMasterModel, PlaceholderTemplate } from "./models";
import { placeholderKey } from "./models";

export interface InheritanceContext {
  layout: SlideLayoutModel;
  master: SlideMasterModel;
  /** Presentation-level default text styles (p:defaultTextStyle) */
  presentationDefaultTextStyle?: Map<number, ParagraphStyle>;
}

/**
 * Resolve all placeholder elements on a slide against its layout and master.
 * Returns a new Slide with inherited geometry, styles, AND master/layout
 * background shapes / missing placeholder slots applied.
 */
export function resolveSlideInheritance(slide: Slide, ctx: InheritanceContext): Slide {
  const resolvedSlideElements = slide.elements.map((el) => resolveElement(el, ctx));

  // ── Missing placeholder slots ──────────────────────────────────────────────
  // Collect the placeholder keys already present on this slide.
  const slidePhKeys = new Set(
    slide.elements
      .filter((el) => el.placeholder)
      .map((el) => placeholderKey(el.placeholder!.type, el.placeholder!.idx)),
  );

  // For each layout/master placeholder that isn't in the slide, synthesise a
  // text element so footer text, page numbers, dates, etc. appear.
  const inheritedElements: SlideElement[] = [];
  const seenKeys = new Set<string>();

  // Layout placeholders take priority over master ones.
  for (const [key, tpl] of ctx.layout.placeholders) {
    if (slidePhKeys.has(key) || seenKeys.has(key)) continue;
    const el = buildPlaceholderElement(tpl, slide.index);
    if (el) {
      inheritedElements.push(el);
      seenKeys.add(key);
    }
  }
  for (const [key, tpl] of ctx.master.placeholders) {
    if (slidePhKeys.has(key) || seenKeys.has(key)) continue;
    const el = buildPlaceholderElement(tpl, slide.index);
    if (el) {
      inheritedElements.push(el);
      seenKeys.add(key);
    }
  }

  // ── Background shapes ──────────────────────────────────────────────────────
  // Master shapes render first (lowest z-order), then layout, then slide content.
  // Filter out explicitly hidden shapes so PowerPoint "hidden" master items
  // don't leak through.
  // Namespace IDs to avoid React key collisions with slide element IDs.
  const masterBg = (ctx.master.backgroundShapes ?? [])
    .filter((el) => !el.hidden)
    .map((el) => ({ ...el, id: `__master_${el.id}` }));
  const layoutBg = (ctx.layout.backgroundShapes ?? [])
    .filter((el) => !el.hidden)
    .map((el) => ({ ...el, id: `__layout_${el.id}` }));

  const elements = [...masterBg, ...layoutBg, ...resolvedSlideElements, ...inheritedElements];

  // Background: slide bg → layout bg → master bg
  const background = slide.background ?? ctx.layout.background ?? ctx.master.background;

  // Theme colors: each master may define its own color scheme. Attach the
  // master's theme colors to the slide so renderers can use the correct palette.
  const themeColors = ctx.master.theme?.colors;

  // Color map from <p:clrMap>: defines how semantic aliases (bg1, tx1) resolve to
  // actual theme slots. Critical for dark themes where bg1="dk1" instead of "lt1".
  const colorMap = ctx.master.colorMap;

  // Theme fonts from the master's theme so the renderer can resolve fontTheme refs.
  const themeFonts = ctx.master.theme?.fonts;

  return {
    ...slide,
    elements,
    background,
    ...(themeColors ? { themeColors } : {}),
    ...(colorMap ? { colorMap } : {}),
    ...(themeFonts ? { themeFonts } : {}),
  };
}

// ─── Synthetic element builder ────────────────────────────────────────────────

/**
 * Create a TextShape from a placeholder template so that master/layout
 * content (footer, date, slide number) appears on slides that don't override it.
 */
function buildPlaceholderElement(tpl: PlaceholderTemplate, slideIndex: number): TextShape | null {
  if (!tpl.position || !tpl.size) return null;

  // Resolve slide-number field text to the actual 1-based slide index.
  let paragraphs: Paragraph[] | undefined = tpl.paragraphs;
  if (tpl.phType === "sldNum" && tpl.paragraphs) {
    paragraphs = tpl.paragraphs.map((p) => ({
      ...p,
      runs: p.runs.map((r) =>
        r.type === "field" && r.fieldType?.toLowerCase().includes("slide")
          ? { ...r, text: String(slideIndex + 1) }
          : r,
      ),
    }));
  }

  // Only emit if there's something to render.
  if (!paragraphs?.length && tpl.phType !== "sldNum" && tpl.phType !== "dt") return null;

  const syntheticId = `__ph_${tpl.phType}_${tpl.phIdx}`;
  return {
    type: "text",
    id: syntheticId,
    position: tpl.position,
    size: tpl.size,
    transform: tpl.transform,
    placeholder: { type: tpl.phType as import("../types").PlaceholderType, idx: tpl.phIdx },
    paragraphs: paragraphs ?? [],
    properties: tpl.bodyProperties ?? {},
    fill: tpl.fill,
  } satisfies TextShape;
}

// ─── Per-element resolution ───────────────────────────────────────────────────

function resolveElement(el: SlideElement, ctx: InheritanceContext): SlideElement {
  if (!el.placeholder) return el;

  const key = placeholderKey(el.placeholder.type, el.placeholder.idx);
  const layoutTpl = ctx.layout.placeholders.get(key);
  const masterTpl = ctx.master.placeholders.get(key);

  if (!layoutTpl && !masterTpl) return el;

  return mergeElement(el, layoutTpl, masterTpl, ctx);
}

function mergeElement(
  el: SlideElement,
  layout: PlaceholderTemplate | undefined,
  master: PlaceholderTemplate | undefined,
  ctx: InheritanceContext,
): SlideElement {
  // Geometry: prefer slide > layout > master
  const position = hasPosition(el.position)
    ? el.position
    : (layout?.position ?? master?.position ?? el.position);
  const size = hasSize(el.size) ? el.size : (layout?.size ?? master?.size ?? el.size);
  const transform = el.transform ?? layout?.transform ?? master?.transform;

  // Fill: prefer slide > layout > master
  const fill = getElementFill(el) ?? layout?.fill ?? master?.fill;

  const base = { ...el, position, size, transform };

  if (el.type === "text") {
    return mergeTextShape(el, base as TextShape, layout, master, ctx, fill);
  }
  if (el.type === "shape") {
    return mergeGeometricShape(el, base as GeometricShape, layout, master, fill);
  }
  if (el.type === "image") {
    return { ...(base as ImageShape), ...(fill ? { fill } : {}) };
  }
  if (el.type === "connector") {
    return { ...(base as ConnectorShape), ...(fill ? { fill } : {}) };
  }
  if (el.type === "group") {
    // Group geometry is resolved, children are handled recursively by the slide resolver
    return base;
  }

  return base;
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
  const properties = mergeBodyProperties(
    el.properties,
    layout?.bodyProperties,
    master?.bodyProperties,
  );

  // Select the master txStyles map that matches this placeholder type per OOXML spec.
  // title/ctrTitle → p:titleStyle
  // ftr/dt/sldNum/hdr and non-placeholder text → p:otherStyle
  // body/obj/subTitle/pic/tbl/chart/dgm/media → p:bodyStyle
  const phType = el.placeholder?.type ?? "";
  const masterBaseLevelStyles =
    phType === "title" || phType === "ctrTitle"
      ? ctx.master.titleLevelStyles
      : phType === "body" ||
          phType === "obj" ||
          phType === "subTitle" ||
          phType === "pic" ||
          phType === "tbl" ||
          phType === "chart" ||
          phType === "dgm" ||
          phType === "media"
        ? ctx.master.bodyLevelStyles
        : ctx.master.otherLevelStyles;

  // Merge paragraph-level styles with full 7-level inheritance:
  // 1. presentation.defaultTextStyle
  // 2. master.defaultTextStyle
  // 3. master txStyles (by category)
  // 4. master placeholder lstStyle
  // 5. layout placeholder lstStyle
  const mergedLevelStyles = mergeLevelStyleMaps(
    ctx.presentationDefaultTextStyle ?? new Map(),
    ctx.master.defaultTextStyle ?? new Map(),
    masterBaseLevelStyles,
    master?.levelStyles ?? new Map(),
    layout?.levelStyles ?? new Map(),
  );

  const paragraphs = el.paragraphs.map((p) => {
    const level = p.style.level ?? 0;
    // In OOXML, lstStyle uses 1-based numbering: lvl1pPr (key 1) applies to
    // paragraphs with pPr.lvl=0 (default level), lvl2pPr (key 2) to lvl=1, etc.
    // defPPr (key 0) is the catch-all default for all levels.
    const inheritedStyle = mergedLevelStyles.get(level + 1) ?? mergedLevelStyles.get(0);
    if (!inheritedStyle) return p;

    return {
      ...p,
      style: mergeParaStyle(inheritedStyle, p.style),
    };
  });

  return {
    ...base,
    properties,
    paragraphs,
    ...(fill ? { fill } : {}),
  };
}

function mergeGeometricShape(
  el: GeometricShape,
  base: GeometricShape,
  layout: PlaceholderTemplate | undefined,
  master: PlaceholderTemplate | undefined,
  fill: Fill | undefined,
): GeometricShape {
  if (!el.body) return { ...base, ...(fill ? { fill } : {}) };

  const bodyProperties = mergeBodyProperties(
    el.body.properties,
    layout?.bodyProperties,
    master?.bodyProperties,
  );

  return {
    ...base,
    ...(fill ? { fill } : {}),
    body: { ...el.body, properties: bodyProperties },
  };
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
  };
}

/**
 * Merge run styles per-property so that a child with only `bold: true`
 * still inherits the parent's fontSize, color, etc.
 */
function mergeRunStyle(
  parent: RunStyle | undefined,
  child: RunStyle | undefined,
): RunStyle | undefined {
  if (!parent && !child) return undefined;
  if (!parent) return child;
  if (!child) return parent;
  return {
    bold: child.bold ?? parent.bold,
    italic: child.italic ?? parent.italic,
    underline: child.underline ?? parent.underline,
    strikethrough: child.strikethrough ?? parent.strikethrough,
    fontSize: child.fontSize ?? parent.fontSize,
    color: child.color ?? parent.color,
    fontFamily: child.fontFamily ?? parent.fontFamily,
    fontEa: child.fontEa ?? parent.fontEa,
    fontCs: child.fontCs ?? parent.fontCs,
    fontTheme: child.fontTheme ?? parent.fontTheme,
    letterSpacing: child.letterSpacing ?? parent.letterSpacing,
    kern: child.kern ?? parent.kern,
    cap: child.cap ?? parent.cap,
    highlight: child.highlight ?? parent.highlight,
    baseline: child.baseline ?? parent.baseline,
    link: child.link ?? parent.link,
    language: child.language ?? parent.language,
  };
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
    defaultRunStyle: mergeRunStyle(parent.defaultRunStyle, child.defaultRunStyle),
  };
}

/**
 * Merge multiple level-style maps, where later maps take priority.
 * body (master txStyles) < master placeholder lstStyle < layout placeholder lstStyle
 */
function mergeLevelStyleMaps(
  ...maps: Array<Map<number, ParagraphStyle>>
): Map<number, ParagraphStyle> {
  const result = new Map<number, ParagraphStyle>();
  for (const map of maps) {
    for (const [level, style] of map) {
      const existing = result.get(level);
      result.set(level, existing ? mergeParaStyle(existing, style) : style);
    }
  }
  return result;
}

// ─── Utility ─────────────────────────────────────────────────────────────────

function hasPosition(pos: Position): boolean {
  return pos.x !== 0 || pos.y !== 0;
}

function hasSize(size: Size): boolean {
  return size.width !== 0 || size.height !== 0;
}

function getElementFill(el: SlideElement): Fill | undefined {
  if (el.type === "group" || el.type === "table" || el.type === "chart") return undefined;
  return el.fill;
}
