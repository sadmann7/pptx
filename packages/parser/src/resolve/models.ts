/**
 * Slide Master and Slide Layout intermediate models.
 *
 * These capture everything a slide needs to inherit:
 *   - Placeholder geometry (position + size)
 *   - Paragraph-level text styles (lstStyle, per indent level 0-8)
 *   - Body properties
 *   - Background fill
 *   - Background (non-placeholder) shapes from the master/layout
 */

import type {
  Background,
  BodyProperties,
  ColorMap,
  Fill,
  Paragraph,
  ParagraphStyle,
  Position,
  Size,
  SlideElement,
  Theme,
  ThemeColors,
  Transform,
} from "../types";
import { parseFill, parseBackground } from "../parsers/fill";
import { parseBodyProperties, parseParagraphStyle, parseTextBody } from "../parsers/text";
import { parseSpTree } from "../parsers/shape";
import { parseTheme } from "../parsers/theme";
import type { PptxZip } from "../zip";
import { loadRels, readString } from "../zip";
import { parseXml } from "../xml";
import { attr, attrNum, extractSpTreeChildOrder, get, toArray } from "../xml";
import { angleToDegs, emuToPoints } from "../emu";

// ─── Types ───────────────────────────────────────────────────────────────────

export interface PlaceholderTemplate {
  /** 'title' | 'body' | 'obj' | 'dt' | 'ftr' | 'sldNum' | etc. */
  phType: string;
  phIdx: number;
  position?: Position;
  size?: Size;
  transform?: Transform;
  fill?: Fill;
  bodyProperties?: BodyProperties;
  /** Paragraph styles keyed by indent level (0 = default, 1-8 = list levels) */
  levelStyles: Map<number, ParagraphStyle>;
  /** Pre-parsed paragraphs from the master/layout text body (e.g. footer text) */
  paragraphs?: Paragraph[];
}

export interface SlideLayoutModel {
  path: string;
  /** Path to the parent slide master */
  masterPath: string;
  background?: Background;
  /** Key: `${phType}/${phIdx}` */
  placeholders: Map<string, PlaceholderTemplate>;
  /**
   * Non-placeholder shapes from the layout that appear on every slide
   * (decorative lines, logos, watermarks, etc.)
   */
  backgroundShapes: SlideElement[];
}

export interface SlideMasterModel {
  path: string;
  background?: Background;
  /** Key: `${phType}/${phIdx}` */
  placeholders: Map<string, PlaceholderTemplate>;
  /** txStyles body paragraph styles — used for body/obj/pic/tbl/chart/dgm placeholders */
  bodyLevelStyles: Map<number, ParagraphStyle>;
  /** txStyles title paragraph styles — used for title/ctrTitle placeholders */
  titleLevelStyles: Map<number, ParagraphStyle>;
  /** txStyles other paragraph styles — used for ftr/dt/sldNum/hdr and non-placeholder text */
  otherLevelStyles: Map<number, ParagraphStyle>;
  /**
   * Non-placeholder shapes from the master that appear on every slide
   * (decorative lines, logos, watermarks, etc.)
   */
  backgroundShapes: SlideElement[];
  /**
   * Theme parsed from this master's own rels.
   * Each slide master in a multi-master PPTX can have a distinct theme
   * (colors, fonts, effects). Slides must resolve colors against their
   * master's theme, not the presentation-level theme.
   */
  theme?: Theme;
  /**
   * Color map from <p:clrMap> — maps semantic aliases (bg1, tx1) to actual theme slots.
   * Critical for dark-theme presentations where bg1="dk1" instead of the default "lt1".
   */
  colorMap?: ColorMap;
  /** Master-level default text styles (p:sldMaster > p:defaultTextStyle) */
  defaultTextStyle?: Map<number, ParagraphStyle>;
}

// ─── Cache ───────────────────────────────────────────────────────────────────

const layoutCache = new WeakMap<PptxZip, Map<string, SlideLayoutModel>>();
const masterCache = new WeakMap<PptxZip, Map<string, SlideMasterModel>>();

// ─── Public loaders ──────────────────────────────────────────────────────────

export async function loadLayoutModel(zip: PptxZip, layoutPath: string): Promise<SlideLayoutModel> {
  let cache = layoutCache.get(zip);
  if (!cache) {
    cache = new Map();
    layoutCache.set(zip, cache);
  }
  if (cache.has(layoutPath)) return cache.get(layoutPath)!;

  const model = await parseLayoutModel(zip, layoutPath);
  cache.set(layoutPath, model);
  return model;
}

export async function loadMasterModel(zip: PptxZip, masterPath: string): Promise<SlideMasterModel> {
  let cache = masterCache.get(zip);
  if (!cache) {
    cache = new Map();
    masterCache.set(zip, cache);
  }
  if (cache.has(masterPath)) return cache.get(masterPath)!;

  const model = await parseMasterModel(zip, masterPath);
  cache.set(masterPath, model);
  return model;
}

// ─── Parsing ─────────────────────────────────────────────────────────────────

async function parseLayoutModel(zip: PptxZip, layoutPath: string): Promise<SlideLayoutModel> {
  const rawXml = await readString(zip, layoutPath);
  const xml = rawXml ? parseXml(rawXml) : {};
  const rels = await loadRels(zip, layoutPath);

  // Find parent master path
  const masterRel = [...rels.values()].find((r) => r.type.includes("slideMaster"));
  const masterPath = masterRel?.target ?? "ppt/slideMasters/slideMaster1.xml";

  const cSld = get(xml, "p:sldLayout", "p:cSld") as Record<string, unknown> | undefined;

  const bg = get(cSld, "p:bg");
  const bgFill = bg ? await parseBackground(bg, zip, rels) : undefined;
  const background = bgFill ? { fill: bgFill } : undefined;

  const spTree = get(cSld, "p:spTree") as Record<string, unknown> | undefined;
  const placeholders = spTree ? extractPlaceholders(spTree) : new Map();

  const childOrder = rawXml
    ? extractSpTreeChildOrder(rawXml, ["p:sldLayout", "p:cSld", "p:spTree"])
    : undefined;
  const allShapes = spTree
    ? await parseSpTree(spTree, rels, zip, layoutPath, false, undefined, childOrder)
    : [];
  const backgroundShapes = allShapes.filter((el) => !el.placeholder);

  return { path: layoutPath, masterPath, background, placeholders, backgroundShapes };
}

async function parseMasterModel(zip: PptxZip, masterPath: string): Promise<SlideMasterModel> {
  const rawXml = await readString(zip, masterPath);
  const xml = rawXml ? parseXml(rawXml) : {};
  const rels = await loadRels(zip, masterPath);

  const cSld = get(xml, "p:sldMaster", "p:cSld") as Record<string, unknown> | undefined;

  const spTree = get(cSld, "p:spTree") as Record<string, unknown> | undefined;
  const placeholders = spTree ? extractPlaceholders(spTree) : new Map();

  const childOrder = rawXml
    ? extractSpTreeChildOrder(rawXml, ["p:sldMaster", "p:cSld", "p:spTree"])
    : undefined;
  const allShapes = spTree
    ? await parseSpTree(spTree, rels, zip, masterPath, false, undefined, childOrder)
    : [];
  const backgroundShapes = allShapes.filter((el) => !el.placeholder);

  // txStyles — master-level paragraph style defaults for title, body, and other text
  const txStyles = get(xml, "p:sldMaster", "p:txStyles") as Record<string, unknown> | undefined;
  const titleStyles = get(txStyles, "p:titleStyle") as Record<string, unknown> | undefined;
  const bodyStyles = get(txStyles, "p:bodyStyle") as Record<string, unknown> | undefined;
  const otherStyles = get(txStyles, "p:otherStyle") as Record<string, unknown> | undefined;
  const titleLevelStyles = parseLevelStyles(titleStyles);
  const bodyLevelStyles = parseLevelStyles(bodyStyles);
  const otherLevelStyles = parseLevelStyles(otherStyles);

  // Each master can have its own theme (colors, fonts, effects). Load it from
  // the master's own rels so per-master color schemes resolve correctly.
  const themeRel = [...rels.values()].find((r) => r.type.includes("theme"));
  let theme: Theme | undefined;
  if (themeRel) {
    try {
      const themeStr = await readString(zip, themeRel.target);
      const themeXml = themeStr ? parseXml(themeStr) : {};
      theme = parseTheme(themeXml);
    } catch {
      // theme parsing is optional — fall back to presentation-level theme
    }
  }

  // Parse background with theme available for p:bgRef resolution
  const bg = get(cSld, "p:bg");
  const bgFill = bg ? await parseBackground(bg, zip, rels, theme) : undefined;
  const background = bgFill ? { fill: bgFill } : undefined;

  // Parse <p:clrMap> which maps semantic color aliases (bg1, tx1) to actual theme slots.
  // This is critical for dark themes: bg1="dk1" means backgrounds use the dk1 slot (dark),
  // while a standard light theme has bg1="lt1".
  const clrMapNode = get(xml, "p:sldMaster", "p:clrMap") as Record<string, unknown> | undefined;
  const colorMap = clrMapNode ? parseColorMap(clrMapNode) : undefined;

  // Master-level default text styles
  const defaultTextStyleNode = get(xml, "p:sldMaster", "p:defaultTextStyle") as
    | Record<string, unknown>
    | undefined;
  const defaultTextStyle = parseLevelStyles(defaultTextStyleNode);

  return {
    path: masterPath,
    background,
    placeholders,
    titleLevelStyles,
    bodyLevelStyles,
    otherLevelStyles,
    backgroundShapes,
    theme,
    colorMap,
    ...(defaultTextStyle.size > 0 ? { defaultTextStyle } : {}),
  };
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

/**
 * Valid theme color slot names. Used to validate clrMap attribute values.
 */
const VALID_THEME_SLOTS = new Set<keyof ThemeColors>([
  "dk1",
  "dk2",
  "lt1",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
]);

/**
 * Parse a <p:clrMap> element into a ColorMap.
 * Each attribute maps a semantic alias to a theme color slot.
 * Example: bg1="dk1" means "when bg1 is used, look up dk1 from theme colors".
 */
function parseColorMap(node: Record<string, unknown>): ColorMap {
  const colorMap: ColorMap = {};
  // All known semantic aliases
  const aliases = [
    "bg1",
    "tx1",
    "bg2",
    "tx2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hlink",
    "folHlink",
  ];
  for (const alias of aliases) {
    const val = attr(node, alias) ?? attr(node, `@_${alias}`);
    if (val && VALID_THEME_SLOTS.has(val as keyof ThemeColors)) {
      (colorMap as Record<string, unknown>)[alias] = val as keyof ThemeColors;
    }
  }
  return colorMap;
}

function placeholderKey(phType: string, phIdx: number): string {
  return `${phType}/${phIdx}`;
}

export { placeholderKey };

function extractPlaceholders(spTree: Record<string, unknown>): Map<string, PlaceholderTemplate> {
  const map = new Map<string, PlaceholderTemplate>();

  const spNodes = toArray(spTree["p:sp"] as unknown[]);
  for (const sp of spNodes) {
    const spN = sp as Record<string, unknown>;
    const nvSpPr = get(spN, "p:nvSpPr") as Record<string, unknown> | undefined;
    const nvPr = get(nvSpPr, "p:nvPr");
    const ph = get(nvPr, "p:ph");
    if (!ph) continue; // not a placeholder

    const phType = attr(ph, "type") ?? "obj";
    const phIdx = attrNum(ph, "idx") ?? 0;

    // Geometry
    const spPr = get(spN, "p:spPr") as Record<string, unknown> | undefined;
    const xfrm = get(spPr, "a:xfrm");
    const { position, size, transform } = extractXfrm(xfrm);

    const fill = spPr ? parseFill(spPr) : undefined;

    // Body properties
    const txBody = get(spN, "p:txBody") as Record<string, unknown> | undefined;
    const bodyPrNode = get(txBody, "a:bodyPr");
    const bodyProperties = bodyPrNode ? parseBodyProperties(bodyPrNode) : undefined;

    // Per-level paragraph styles from lstStyle
    const lstStyle = get(txBody, "a:lstStyle") as Record<string, unknown> | undefined;
    const levelStyles = parseLevelStyles(lstStyle);

    // Text content from the master/layout placeholder (e.g. footer text)
    let paragraphs: Paragraph[] | undefined;
    if (txBody) {
      const parsed = parseTextBody(txBody);
      if (parsed.paragraphs.some((p) => p.runs.some((r) => r.type === "run" && r.text))) {
        paragraphs = parsed.paragraphs;
      }
    }

    const template: PlaceholderTemplate = {
      phType,
      phIdx,
      fill,
      bodyProperties,
      levelStyles,
      ...(paragraphs ? { paragraphs } : {}),
    };
    if (position) template.position = position;
    if (size) template.size = size;
    if (transform) template.transform = transform;

    map.set(placeholderKey(phType, phIdx), template);
  }

  return map;
}

function extractXfrm(xfrmNode: unknown): {
  position?: Position;
  size?: Size;
  transform?: Transform;
} {
  if (!xfrmNode || typeof xfrmNode !== "object") return {};

  const n = xfrmNode as Record<string, unknown>;
  let position: Position | undefined;
  let size: Size | undefined;
  let transform: Transform | undefined;

  const off = get(n, "a:off");
  if (off) {
    position = {
      x: emuToPoints(attrNum(off, "x") ?? 0),
      y: emuToPoints(attrNum(off, "y") ?? 0),
    };
  }

  const ext = get(n, "a:ext");
  if (ext) {
    size = {
      width: emuToPoints(attrNum(ext, "cx") ?? 0),
      height: emuToPoints(attrNum(ext, "cy") ?? 0),
    };
  }

  const rot = attrNum(n, "rot");
  const flipH = attr(n, "flipH") === "1";
  const flipV = attr(n, "flipV") === "1";
  if (rot !== undefined || flipH || flipV) {
    transform = {
      rotation: rot !== undefined ? angleToDegs(rot) : undefined,
      flipH: flipH || undefined,
      flipV: flipV || undefined,
    };
  }

  return { position, size, transform };
}

/**
 * Parse per-level paragraph styles from a lstStyle or txStyles node.
 * Keys 0 = defPPr (default), 1-8 = lvl1pPr through lvl8pPr, 9 = lvl9pPr.
 */
export function parseLevelStyles(node: unknown): Map<number, ParagraphStyle> {
  const map = new Map<number, ParagraphStyle>();
  if (!node || typeof node !== "object") return map;

  const n = node as Record<string, unknown>;

  const defPPr = get(n, "a:defPPr");
  if (defPPr) map.set(0, parseParagraphStyle(defPPr));

  for (let lvl = 1; lvl <= 9; lvl++) {
    const key = `a:lvl${lvl}pPr`;
    const raw = n[key];
    if (!raw) continue;
    // Guard against accidental ALWAYS_ARRAY wrapping (lvlNpPr never repeats).
    const lvlNode = Array.isArray(raw) ? raw[0] : raw;
    if (lvlNode) map.set(lvl, parseParagraphStyle(lvlNode));
  }

  return map;
}
