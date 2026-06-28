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
  Fill,
  Paragraph,
  ParagraphStyle,
  Position,
  Size,
  SlideElement,
  Transform,
} from "../types";
import { parseFill } from "../parsers/fill";
import { parseBodyProperties, parseParagraphStyle, parseTextBody } from "../parsers/text";
import { parseSpTree } from "../parsers/shape";
import type { PptxZip } from "../zip";
import { loadRels, readXml } from "../zip";
import { attr, attrNum, get, toArray } from "../xml";
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
  /** txStyles per-level paragraph styles (master-level body defaults) */
  bodyLevelStyles: Map<number, ParagraphStyle>;
  /**
   * Non-placeholder shapes from the master that appear on every slide
   * (decorative lines, logos, watermarks, etc.)
   */
  backgroundShapes: SlideElement[];
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
  const xml = await readXml(zip, layoutPath);
  const rels = await loadRels(zip, layoutPath);

  // Find parent master path
  const masterRel = [...rels.values()].find((r) => r.type.includes("slideMaster"));
  const masterPath = masterRel?.target ?? "ppt/slideMasters/slideMaster1.xml";

  const cSld = get(xml, "p:sldLayout", "p:cSld") as Record<string, unknown> | undefined;

  const bg = get(cSld, "p:bg");
  const background = bg ? parseBackgroundFill(bg) : undefined;

  const spTree = get(cSld, "p:spTree") as Record<string, unknown> | undefined;
  const placeholders = spTree ? extractPlaceholders(spTree) : new Map();

  // Parse all shapes (skipImages: true — masters rarely embed images)
  const allShapes = spTree ? await parseSpTree(spTree, rels, zip, layoutPath, true) : [];
  // Non-placeholder shapes appear on every slide as background decoration
  const backgroundShapes = allShapes.filter((el) => !el.placeholder);

  return { path: layoutPath, masterPath, background, placeholders, backgroundShapes };
}

async function parseMasterModel(zip: PptxZip, masterPath: string): Promise<SlideMasterModel> {
  const xml = await readXml(zip, masterPath);
  const rels = await loadRels(zip, masterPath);

  const cSld = get(xml, "p:sldMaster", "p:cSld") as Record<string, unknown> | undefined;

  const bg = get(cSld, "p:bg");
  const background = bg ? parseBackgroundFill(bg) : undefined;

  const spTree = get(cSld, "p:spTree") as Record<string, unknown> | undefined;
  const placeholders = spTree ? extractPlaceholders(spTree) : new Map();

  const allShapes = spTree ? await parseSpTree(spTree, rels, zip, masterPath, true) : [];
  const backgroundShapes = allShapes.filter((el) => !el.placeholder);

  // txStyles — master-level paragraph style defaults for body text
  const txStyles = get(xml, "p:sldMaster", "p:txStyles") as Record<string, unknown> | undefined;
  const bodyStyles = get(txStyles, "p:bodyStyle") as Record<string, unknown> | undefined;
  const bodyLevelStyles = parseLevelStyles(bodyStyles);

  return { path: masterPath, background, placeholders, bodyLevelStyles, backgroundShapes };
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

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
function parseLevelStyles(node: unknown): Map<number, ParagraphStyle> {
  const map = new Map<number, ParagraphStyle>();
  if (!node || typeof node !== "object") return map;

  const n = node as Record<string, unknown>;

  const defPPr = get(n, "a:defPPr");
  if (defPPr) map.set(0, parseParagraphStyle(defPPr));

  for (let lvl = 1; lvl <= 9; lvl++) {
    const key = `a:lvl${lvl}pPr`;
    const lvlNode = n[key];
    if (lvlNode) map.set(lvl, parseParagraphStyle(lvlNode));
  }

  return map;
}

function parseBackgroundFill(bgNode: unknown): Background | undefined {
  if (!bgNode || typeof bgNode !== "object") return undefined;
  const n = bgNode as Record<string, unknown>;
  const bgPr = n["p:bgPr"] ?? n["bgPr"];
  if (!bgPr) return undefined;
  const fill = parseFill(bgPr);
  if (!fill) return undefined;
  return { fill };
}
