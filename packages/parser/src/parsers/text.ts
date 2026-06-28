import type {
  BodyProperties,
  Bullet,
  FieldRun,
  LineBreak,
  LineSpacing,
  Paragraph,
  ParagraphContent,
  ParagraphStyle,
  RunStyle,
  TextAlignment,
  TextRun,
  TextVerticalAlignment,
  UnderlineStyle,
} from "../types";
import { parseColor } from "../color";
import { angleToDegs, emuToPoints, hunPtToPoints, perMilleToPercent } from "../emu";
import { attr, attrBool, attrNum, get, textContent, toArray } from "../xml";

// ─── Body properties ──────────────────────────────────────────────────────────

export function parseBodyProperties(bodyPrNode: unknown): BodyProperties {
  const props: BodyProperties = {};
  if (!bodyPrNode || typeof bodyPrNode !== "object") return props;

  const n = bodyPrNode as Record<string, unknown>;

  const vert = attr(n, "vert");
  if (vert) props.direction = vert as BodyProperties["direction"];

  const anchor = attr(n, "anchor");
  if (anchor) props.verticalAlignment = anchor as TextVerticalAlignment;

  const wrap = attr(n, "wrap");
  if (wrap) props.wrap = wrap as BodyProperties["wrap"];

  const rot = attrNum(n, "rot");
  if (rot !== undefined) props.rotation = angleToDegs(rot);

  // Insets (in EMU)
  const inL = attrNum(n, "lIns");
  const inR = attrNum(n, "rIns");
  const inT = attrNum(n, "tIns");
  const inB = attrNum(n, "bIns");

  if (inL !== undefined) props.insetLeft = emuToPoints(inL);
  if (inR !== undefined) props.insetRight = emuToPoints(inR);
  if (inT !== undefined) props.insetTop = emuToPoints(inT);
  if (inB !== undefined) props.insetBottom = emuToPoints(inB);

  const numCol = attrNum(n, "numCol");
  if (numCol !== undefined) props.columns = numCol;

  const spcCol = attrNum(n, "spcCol");
  if (spcCol !== undefined) props.columnSpacing = emuToPoints(spcCol);

  // Autofit
  if ("a:spAutoFit" in n || "spAutoFit" in n) props.autofit = "spAutoFit";
  else if ("a:normAutoFit" in n || "normAutoFit" in n) props.autofit = "normAutoFit";
  else if ("a:noAutofit" in n || "noAutofit" in n) props.autofit = "none";

  return props;
}

// ─── Line spacing ─────────────────────────────────────────────────────────────

function parseLineSpacing(node: unknown): LineSpacing | undefined {
  if (!node || typeof node !== "object") return undefined;
  const n = node as Record<string, unknown>;

  const spcPct = get(n, "a:spcPct");
  if (spcPct) {
    const val = attrNum(spcPct, "val");
    if (val !== undefined) return { value: perMilleToPercent(val), unit: "pct" };
  }

  const spcPts = get(n, "a:spcPts");
  if (spcPts) {
    const val = attrNum(spcPts, "val");
    if (val !== undefined) return { value: hunPtToPoints(val), unit: "pt" };
  }

  return undefined;
}

// ─── Bullet ───────────────────────────────────────────────────────────────────

function parseBullet(pPrNode: Record<string, unknown>): Bullet | undefined {
  if ("a:buNone" in pPrNode) return { type: "none" };

  if ("a:buChar" in pPrNode) {
    const char = attr(get(pPrNode, "a:buChar"), "char") ?? "•";
    const color = parseColor(get(pPrNode, "a:buClr"));
    const sz = attrNum(get(pPrNode, "a:buSzPct"), "val");
    const size = sz !== undefined ? perMilleToPercent(sz) : undefined;
    const fontFamily = attr(get(pPrNode, "a:buFont"), "typeface");
    return { type: "char", char, color, size, fontFamily };
  }

  if ("a:buAutoNum" in pPrNode) {
    const autoNode = pPrNode["a:buAutoNum"] as Record<string, unknown>;
    const style = attr(autoNode, "type") ?? "arabicPeriod";
    const startAt = attrNum(autoNode, "startAt");
    const color = parseColor(get(pPrNode, "a:buClr"));
    const sz = attrNum(get(pPrNode, "a:buSzPct"), "val");
    const size = sz !== undefined ? perMilleToPercent(sz) : undefined;
    return { type: "numeric", style, startAt, color, size };
  }

  // Inherited / default bullet — return undefined and let the caller inherit
  return undefined;
}

// ─── Paragraph style ─────────────────────────────────────────────────────────

export function parseParagraphStyle(pPrNode: unknown): ParagraphStyle {
  const style: ParagraphStyle = {};
  if (!pPrNode || typeof pPrNode !== "object") return style;

  const n = pPrNode as Record<string, unknown>;

  const algn = attr(n, "algn");
  if (algn) style.alignment = algn as TextAlignment;

  const lvl = attrNum(n, "lvl");
  if (lvl !== undefined) style.level = lvl;

  const indent = attrNum(n, "indent");
  if (indent !== undefined) style.indent = emuToPoints(indent);

  const marL = attrNum(n, "marL");
  if (marL !== undefined) style.marginLeft = emuToPoints(marL);

  const marR = attrNum(n, "marR");
  if (marR !== undefined) style.marginRight = emuToPoints(marR);

  const spcBef = get(n, "a:spcBef");
  if (spcBef) style.spaceBefore = parseLineSpacing(spcBef);

  const spcAft = get(n, "a:spcAft");
  if (spcAft) style.spaceAfter = parseLineSpacing(spcAft);

  const lnSpc = get(n, "a:lnSpc");
  if (lnSpc) style.lineSpacing = parseLineSpacing(lnSpc);

  const bullet = parseBullet(n);
  if (bullet) style.bullet = bullet;

  // Default run properties inside pPr
  const defRPr = get(n, "a:defRPr");
  if (defRPr) style.defaultRunStyle = parseRunStyle(defRPr);

  return style;
}

// ─── Run style ────────────────────────────────────────────────────────────────

export function parseRunStyle(rPrNode: unknown): RunStyle {
  const style: RunStyle = {};
  if (!rPrNode || typeof rPrNode !== "object") return style;

  const n = rPrNode as Record<string, unknown>;

  const b = attrBool(n, "b");
  if (b !== undefined) style.bold = b;

  const i = attrBool(n, "i");
  if (i !== undefined) style.italic = i;

  const u = attr(n, "u");
  if (u && u !== "none") style.underline = u as UnderlineStyle;

  const strike = attr(n, "strike");
  if (strike && strike !== "noStrike") style.strikethrough = true;

  const sz = attrNum(n, "sz");
  if (sz !== undefined) style.fontSize = hunPtToPoints(sz);

  const baseline = attrNum(n, "baseline");
  if (baseline !== undefined) style.baseline = perMilleToPercent(baseline);

  const lang = attr(n, "lang");
  if (lang) style.language = lang;

  const dirty = attrBool(n, "dirty");
  if (dirty !== undefined) style.dirty = dirty;

  // Color
  const solidFill = get(n, "a:solidFill");
  if (solidFill) {
    const color = parseColor(solidFill);
    if (color) style.color = color;
  }

  // Font family
  const latin = get(n, "a:latin");
  if (latin) {
    const typeface = attr(latin, "typeface");
    if (typeface && typeface !== "+mj-lt" && typeface !== "+mn-lt") {
      style.fontFamily = typeface;
    } else if (typeface === "+mj-lt") {
      style.fontTheme = "major";
    } else if (typeface === "+mn-lt") {
      style.fontTheme = "minor";
    }
  }

  // Hyperlink
  const hlinkClick = get(n, "a:hlinkClick");
  if (hlinkClick) {
    style.link = attr(hlinkClick, "r:id") ?? attr(hlinkClick, "id") ?? "";
  }

  return style;
}

// ─── Paragraph + run parsing ──────────────────────────────────────────────────

export function parseParagraph(pNode: unknown): Paragraph {
  if (!pNode || typeof pNode !== "object") {
    return { runs: [], style: {} };
  }

  const n = pNode as Record<string, unknown>;
  const pPr = get(n, "a:pPr");
  const style = parseParagraphStyle(pPr);

  const runs: ParagraphContent[] = [];

  // Process children in order: a:r, a:br, a:fld, a:endParaRPr
  // fast-xml-parser doesn't preserve order, so we collect all three types
  const rNodes = toArray(n["a:r"] as unknown[]);
  const brNodes = toArray(n["a:br"] as unknown[]);
  const fldNodes = toArray(n["a:fld"] as unknown[]);

  // We need to handle ordering — fast-xml-parser loses element order.
  // For now, emit runs → linebreaks → fields. This covers 95%+ of slides.
  // True ordering requires a SAX-based approach (future improvement).
  for (const r of rNodes) {
    const run = parseRun(r);
    if (run) runs.push(run);
  }

  for (const br of brNodes) {
    const lb = parseLineBreak(br);
    if (lb) runs.push(lb);
  }

  for (const fld of fldNodes) {
    const field = parseField(fld);
    if (field) runs.push(field);
  }

  // End-paragraph run properties for empty paragraphs (e.g. spacing only)
  if (runs.length === 0) {
    const endPr = get(n, "a:endParaRPr");
    if (endPr) {
      const endStyle = parseRunStyle(endPr);
      // Empty paragraph — push a zero-width run to carry the style
      runs.push({ type: "run", text: "", style: endStyle });
    }
  }

  return { runs, style };
}

function parseRun(rNode: unknown): TextRun | undefined {
  if (!rNode || typeof rNode !== "object") return undefined;
  const n = rNode as Record<string, unknown>;

  const t = n["a:t"];
  const text = textContent(t);

  const rPr = get(n, "a:rPr");
  const style = parseRunStyle(rPr);

  return { type: "run", text, style };
}

function parseLineBreak(brNode: unknown): LineBreak | undefined {
  if (brNode === undefined || brNode === null) return undefined;
  return { type: "lineBreak" };
}

function parseField(fldNode: unknown): FieldRun | undefined {
  if (!fldNode || typeof fldNode !== "object") return undefined;
  const n = fldNode as Record<string, unknown>;

  const fieldType = attr(n, "type") ?? "";
  const text = textContent(n["a:t"]);
  const rPr = get(n, "a:rPr");
  const style = parseRunStyle(rPr);

  return { type: "field", fieldType, text, style };
}

/**
 * Parse an entire text body (a:txBody or p:txBody).
 */
export function parseTextBody(txBodyNode: unknown): {
  paragraphs: Paragraph[];
  properties: BodyProperties;
} {
  if (!txBodyNode || typeof txBodyNode !== "object") {
    return { paragraphs: [], properties: {} };
  }

  const n = txBodyNode as Record<string, unknown>;
  const bodyPrNode = get(n, "a:bodyPr");
  const properties = parseBodyProperties(bodyPrNode);

  const pNodes = toArray(n["a:p"] as unknown[]);
  const paragraphs = pNodes.map(parseParagraph);

  return { paragraphs, properties };
}
