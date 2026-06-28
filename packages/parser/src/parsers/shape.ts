import type {
  ChartShape,
  ConnectorShape,
  Effect,
  GeometricShape,
  GroupShape,
  ImageShape,
  PlaceholderInfo,
  PlaceholderType,
  Position,
  Size,
  SlideElement,
  TextShape,
  Transform,
} from "../types";
import { parseFill, parseEffects, parseStroke } from "./fill";
import { parseTable } from "./table";
import { parseTextBody } from "./text";
import type { PptxZip, Relationship } from "../zip";
import { readMediaAsUrl, readString } from "../zip";
import { attr, attrBool, attrNum, get, toArray } from "../xml";
import { angleToDegs, emuToPoints } from "../emu";

// ─── Adjustments ─────────────────────────────────────────────────────────────

/**
 * Parse <a:avLst> adjustment values. Each <a:gd fmla="val N"/> contributes
 * one number, in document order (adj → index 0, adj2 → index 1, …).
 */
function parseAdjustments(avLstNode: unknown): number[] {
  if (!avLstNode || typeof avLstNode !== "object") return [];
  const guides = toArray((avLstNode as Record<string, unknown>)["a:gd"]);
  const result: number[] = [];
  for (const gd of guides) {
    if (!gd || typeof gd !== "object") continue;
    const fmla = attr(gd as Record<string, unknown>, "fmla") ?? "";
    const m = fmla.match(/^val\s+(-?\d+)/);
    if (m?.[1]) result.push(parseInt(m[1], 10));
  }
  return result;
}

// ─── Placeholder ─────────────────────────────────────────────────────────────

/**
 * Read the <p:ph> element from the shape's nvPr and return placeholder info.
 * Returns undefined if the shape is not a placeholder.
 */
function readPlaceholder(nvPrNode: unknown): PlaceholderInfo | undefined {
  const ph = get(nvPrNode, "p:ph");
  if (!ph) return undefined;
  const type = (attr(ph, "type") ?? "obj") as PlaceholderType;
  const idx = attrNum(ph, "idx") ?? 0;
  return { type, idx };
}

// ─── Transform ───────────────────────────────────────────────────────────────

function parseXfrm(xfrmNode: unknown): { position: Position; size: Size; transform?: Transform } {
  const position: Position = { x: 0, y: 0 };
  const size: Size = { width: 0, height: 0 };
  let transform: Transform | undefined;

  if (!xfrmNode || typeof xfrmNode !== "object") {
    return { position, size };
  }

  const n = xfrmNode as Record<string, unknown>;

  const off = get(n, "a:off");
  if (off) {
    position.x = emuToPoints(attrNum(off, "x") ?? 0);
    position.y = emuToPoints(attrNum(off, "y") ?? 0);
  }

  const ext = get(n, "a:ext");
  if (ext) {
    size.width = emuToPoints(attrNum(ext, "cx") ?? 0);
    size.height = emuToPoints(attrNum(ext, "cy") ?? 0);
  }

  const rot = attrNum(n, "rot");
  const flipH = attrBool(n, "flipH");
  const flipV = attrBool(n, "flipV");

  if (rot !== undefined || flipH !== undefined || flipV !== undefined) {
    transform = {
      rotation: rot !== undefined ? angleToDegs(rot) : undefined,
      flipH,
      flipV,
    };
  }

  return { position, size, transform };
}

// ─── Shape (sp) ──────────────────────────────────────────────────────────────

function parseSp(
  spNode: Record<string, unknown>,
  themeEffectStyles?: Effect[][],
): GeometricShape | TextShape {
  const nvSpPr = get(spNode, "p:nvSpPr") as Record<string, unknown> | undefined;
  const cNvPr = get(nvSpPr, "p:cNvPr") as Record<string, unknown> | undefined;
  const id = attr(cNvPr, "id") ?? "";
  const name = attr(cNvPr, "name") ?? "";
  const hidden = attrBool(cNvPr, "hidden") ?? false;
  const placeholder = readPlaceholder(get(nvSpPr, "p:nvPr"));

  const spPr = get(spNode, "p:spPr") as Record<string, unknown> | undefined;
  const xfrm = get(spPr, "a:xfrm");
  const { position, size, transform } = parseXfrm(xfrm);

  const fill = parseFill(spPr);
  const stroke = parseStroke(get(spPr, "a:ln"));

  // Effects: inline effectLst takes priority; fall back to theme effectRef.
  let rawEffects = parseEffects(get(spPr, "a:effectLst"));
  if (!rawEffects.length && themeEffectStyles) {
    const styleNode = get(spNode, "p:style") as Record<string, unknown> | undefined;
    const effectRef = get(styleNode, "a:effectRef");
    if (effectRef) {
      const idx = attrNum(effectRef, "idx");
      if (idx !== undefined && idx > 0 && idx <= themeEffectStyles.length) {
        rawEffects = themeEffectStyles[idx - 1] ?? [];
      }
    }
  }
  const effects = rawEffects.length ? rawEffects : undefined;

  const prstGeom = get(spPr, "a:prstGeom");
  const custGeom = get(spPr, "a:custGeom");

  const txBody = get(spNode, "p:txBody");
  const hasText = !!txBody;

  // Text box (no geometry, has txBody)
  if (!prstGeom && !custGeom && hasText) {
    const { paragraphs, properties } = parseTextBody(txBody);
    return {
      type: "text",
      id,
      name,
      position,
      size,
      transform,
      hidden,
      placeholder,
      paragraphs,
      properties,
      fill,
      stroke,
      effects,
    } satisfies TextShape;
  }

  const shapeType = prstGeom ? (attr(prstGeom, "prst") ?? "rect") : "custom";

  // Parse adjustment values from <a:avLst><a:gd name="adj" fmla="val N"/>…
  const adjustments = parseAdjustments(get(prstGeom, "a:avLst"));

  let body: GeometricShape["body"] | undefined;
  if (hasText) {
    const { paragraphs, properties } = parseTextBody(txBody);
    body = { paragraphs, properties };
  }

  return {
    type: "shape",
    id,
    name,
    position,
    size,
    transform,
    hidden,
    placeholder,
    shapeType,
    ...(adjustments.length ? { adjustments } : {}),
    fill,
    stroke,
    effects,
    body,
  } satisfies GeometricShape;
}

// ─── Picture (pic) ───────────────────────────────────────────────────────────

async function parsePic(
  picNode: Record<string, unknown>,
  rels: Map<string, Relationship>,
  zip: PptxZip,
  skipImages: boolean,
): Promise<ImageShape> {
  const nvPicPr = get(picNode, "p:nvPicPr") as Record<string, unknown> | undefined;
  const cNvPr = get(nvPicPr, "p:cNvPr") as Record<string, unknown> | undefined;
  const id = attr(cNvPr, "id") ?? "";
  const name = attr(cNvPr, "name") ?? "";
  const hidden = attrBool(cNvPr, "hidden") ?? false;
  const placeholder = readPlaceholder(get(nvPicPr, "p:nvPr"));

  const spPr = get(picNode, "p:spPr") as Record<string, unknown> | undefined;
  const xfrm = get(spPr, "a:xfrm");
  const { position, size, transform } = parseXfrm(xfrm);

  const fill = parseFill(spPr);
  const stroke = parseStroke(get(spPr, "a:ln"));

  // Relationship ID → media path
  const blipFill = get(picNode, "p:blipFill") as Record<string, unknown> | undefined;
  const blip = get(blipFill, "a:blip") as Record<string, unknown> | undefined;
  const rId = attr(blip, "r:embed") ?? attr(blip, "embed") ?? "";

  // Crop rectangle (srcRect)
  let cropRect: ImageShape["cropRect"];
  const srcRect = get(blipFill, "a:srcRect");
  if (srcRect) {
    const t = attrNum(srcRect, "t") ?? 0;
    const r = attrNum(srcRect, "r") ?? 0;
    const b = attrNum(srcRect, "b") ?? 0;
    const l = attrNum(srcRect, "l") ?? 0;
    if (t || r || b || l) {
      cropRect = {
        top: t / 100000,
        right: r / 100000,
        bottom: b / 100000,
        left: l / 100000,
      };
    }
  }

  let src = "";
  let mimeType = "";

  if (!skipImages && rId) {
    const rel = rels.get(rId);
    if (rel) {
      const media = await readMediaAsUrl(zip, rel.target);
      src = media.src;
      mimeType = media.mimeType;
    }
  }

  return {
    type: "image",
    id,
    name,
    position,
    size,
    transform,
    hidden,
    placeholder,
    src,
    mimeType,
    rId,
    cropRect,
    fill,
    stroke,
  };
}

// ─── GraphicFrame (tables, charts, diagrams) ──────────────────────────────────

async function parseGraphicFrame(
  gfNode: Record<string, unknown>,
  rels: Map<string, Relationship>,
  zip: PptxZip,
  _slidePath: string,
): Promise<SlideElement | null> {
  const nvGrFrPr = get(gfNode, "p:nvGraphicFramePr") as Record<string, unknown> | undefined;
  const cNvPr = get(nvGrFrPr, "p:cNvPr") as Record<string, unknown> | undefined;
  const id = attr(cNvPr, "id") ?? "";
  const name = attr(cNvPr, "name") ?? "";

  const xfrm = get(gfNode, "p:xfrm");
  const { position, size } = parseXfrm(xfrm);

  const graphic = get(gfNode, "a:graphic") as Record<string, unknown> | undefined;
  const graphicData = get(graphic, "a:graphicData") as Record<string, unknown> | undefined;
  const uri = attr(graphicData, "uri") ?? "";

  // Table
  if (uri.includes("table") || "a:tbl" in (graphicData ?? {})) {
    return parseTable(gfNode, id, name, position, size);
  }

  // Chart
  if (uri.includes("chart")) {
    const chartEl = graphicData
      ? Object.values(graphicData).find((v) => v && typeof v === "object" && attr(v, "r:id"))
      : undefined;
    const chartRId = chartEl ? (attr(chartEl, "r:id") ?? attr(chartEl, "id") ?? "") : "";
    const rel = rels.get(chartRId);

    let chartXml = "";
    if (rel) {
      chartXml = await readString(zip, rel.target);
    }

    return {
      type: "chart",
      id,
      name,
      position,
      size,
      rId: chartRId,
      chartXml,
    } satisfies ChartShape;
  }

  // Unknown graphic frame — skip
  return null;
}

// ─── Connector (cxnSp) ───────────────────────────────────────────────────────

function parseCxnSp(cxnNode: Record<string, unknown>): ConnectorShape {
  const nvCxnSpPr = get(cxnNode, "p:nvCxnSpPr") as Record<string, unknown> | undefined;
  const cNvPr = get(nvCxnSpPr, "p:cNvPr") as Record<string, unknown> | undefined;
  const id = attr(cNvPr, "id") ?? "";
  const name = attr(cNvPr, "name") ?? "";
  const hidden = attrBool(cNvPr, "hidden") ?? false;
  const placeholder = readPlaceholder(get(nvCxnSpPr, "p:nvPr"));

  const spPr = get(cxnNode, "p:spPr") as Record<string, unknown> | undefined;
  const xfrm = get(spPr, "a:xfrm");
  const { position, size, transform } = parseXfrm(xfrm);

  const prstGeom = get(spPr, "a:prstGeom");
  const shapeType = attr(prstGeom, "prst") ?? "line";

  const fill = parseFill(spPr);
  const stroke = parseStroke(get(spPr, "a:ln"));

  return {
    type: "connector",
    id,
    name,
    position,
    size,
    transform,
    hidden,
    placeholder,
    shapeType,
    fill,
    stroke,
  };
}

// ─── Group shape (grpSp) ──────────────────────────────────────────────────────

async function parseGrpSp(
  grpNode: Record<string, unknown>,
  rels: Map<string, Relationship>,
  zip: PptxZip,
  slidePath: string,
  skipImages: boolean,
): Promise<GroupShape> {
  const nvGrpSpPr = get(grpNode, "p:nvGrpSpPr") as Record<string, unknown> | undefined;
  const cNvPr = get(nvGrpSpPr, "p:cNvPr") as Record<string, unknown> | undefined;
  const id = attr(cNvPr, "id") ?? "";
  const name = attr(cNvPr, "name") ?? "";

  const grpSpPr = get(grpNode, "p:grpSpPr") as Record<string, unknown> | undefined;
  const xfrm = get(grpSpPr, "a:xfrm");
  const { position, size, transform } = parseXfrm(xfrm);

  const children = await parseSpTree(grpNode, rels, zip, slidePath, skipImages);

  return {
    type: "group",
    id,
    name,
    position,
    size,
    transform,
    children,
  };
}

// ─── SpTree (the main container on a slide) ───────────────────────────────────

export async function parseSpTree(
  treeNode: Record<string, unknown>,
  rels: Map<string, Relationship>,
  zip: PptxZip,
  slidePath: string,
  skipImages: boolean,
  themeEffectStyles?: Effect[][],
): Promise<SlideElement[]> {
  const elements: SlideElement[] = [];

  const spNodes = toArray(treeNode["p:sp"] as unknown[]);
  for (const sp of spNodes) {
    elements.push(parseSp(sp as Record<string, unknown>, themeEffectStyles));
  }

  const picNodes = toArray(treeNode["p:pic"] as unknown[]);
  for (const pic of picNodes) {
    elements.push(await parsePic(pic as Record<string, unknown>, rels, zip, skipImages));
  }

  const gfNodes = toArray(treeNode["p:graphicFrame"] as unknown[]);
  for (const gf of gfNodes) {
    const el = await parseGraphicFrame(gf as Record<string, unknown>, rels, zip, slidePath);
    if (el) elements.push(el);
  }

  const cxnNodes = toArray(treeNode["p:cxnSp"] as unknown[]);
  for (const cxn of cxnNodes) {
    elements.push(parseCxnSp(cxn as Record<string, unknown>));
  }

  const grpNodes = toArray(treeNode["p:grpSp"] as unknown[]);
  for (const grp of grpNodes) {
    elements.push(
      await parseGrpSp(grp as Record<string, unknown>, rels, zip, slidePath, skipImages),
    );
  }

  return elements;
}
