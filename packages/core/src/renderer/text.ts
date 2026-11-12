import { PlaceholderInfo } from "../model/nodes/base";
import type { TextBody, TextParagraph, TextRun } from "../model/nodes/shape";
import { normalizePlaceholderType } from "../model/presentation";
import { parseOoxmlBool } from "../ooxml/boolean";
import { isExternalTargetMode } from "../ooxml/rel";
import { angleToDeg, emuToPx, pctToDecimal } from "../ooxml/unit";
import { SafeXmlNode } from "../ooxml/xml";
import { isAllowedExternalUrl } from "../utils/url";
import { RenderContext } from "./context";
import { PPTX_ATTRS, PPTX_DATASET } from "./dom-attributes";
import { cssFontFamilyStack, resolveThemeFont } from "./font";
import { resolveSlideNavigationIndex, slideJumpTitle } from "./navigation";
import { resolveColor, resolveColorToCss, resolveFill } from "./style";
import { getEffectiveBodyPrChild } from "./text-body-props";

function findStyleAtLevel(styleNode: SafeXmlNode | undefined, level: number): SafeXmlNode {
  if (!styleNode || !styleNode.exists()) {
    return new SafeXmlNode(null);
  }
  const lvlNode = styleNode.child(`lvl${level + 1}pPr`);
  if (lvlNode.exists()) return lvlNode;
  return styleNode.child("defPPr");
}

function getPlaceholderCategory(
  placeholder: PlaceholderInfo | undefined,
): "title" | "body" | "other" {
  if (!placeholder || !placeholder.type) return "other";
  const t = placeholder.type;
  if (t === "title" || t === "ctrTitle") return "title";
  if (
    t === "body" ||
    t === "subTitle" ||
    t === "obj" ||
    t === "dt" ||
    t === "ftr" ||
    t === "sldNum"
  ) {
    return "body";
  }
  return "other";
}

function findPlaceholderNode(
  placeholders: SafeXmlNode[],
  info: PlaceholderInfo,
): SafeXmlNode | undefined {
  for (const ph of placeholders) {
    let phEl: SafeXmlNode | undefined;
    const nvSpPr = ph.child("nvSpPr");
    if (nvSpPr.exists()) {
      phEl = nvSpPr.child("nvPr").child("ph");
    }
    if (!phEl || !phEl.exists()) {
      const nvPicPr = ph.child("nvPicPr");
      if (nvPicPr.exists()) {
        phEl = nvPicPr.child("nvPr").child("ph");
      }
    }
    if (!phEl || !phEl.exists()) continue;

    const phType = phEl.attr("type");
    const phIdx = phEl.numAttr("idx");

    // Matches by idx, then exact type, then equivalent type (title≡ctrTitle, body≡subTitle≡obj).
    if (info.idx !== undefined && phIdx === info.idx) return ph;
    if (info.type && phType === info.type) return ph;
    if (info.type && normalizePlaceholderType(phType) === normalizePlaceholderType(info.type)) {
      return ph;
    }
  }
  return undefined;
}

function getPlaceholderLstStyle(phNode: SafeXmlNode): SafeXmlNode | undefined {
  const txBody = phNode.child("txBody");
  if (!txBody.exists()) return undefined;
  const lstStyle = txBody.child("lstStyle");
  return lstStyle.exists() ? lstStyle : undefined;
}

interface MergedParagraphStyle {
  align?: string;
  rtl?: boolean;
  marginLeft?: number;
  textIndent?: number;
  defaultTabSize?: number;
  lineHeight?: string;
  // Set for spcPts line spacing, which needs per-line wrappers for CJK fonts.
  lineHeightAbsolute?: boolean;
  spaceBefore?: number;
  spaceBeforePct?: number; // Fraction of the font size.
  spaceAfter?: number;
  spaceAfterPct?: number; // Fraction of the font size.
  bulletChar?: string;
  bulletFont?: string;
  bulletAutoNum?: string;
  bulletAutoNumStartAt?: number;
  bulletSizePct?: number;
  bulletSizePt?: number;
  bulletColorFollowsText?: boolean;
  bulletNone?: boolean;
  bulletColorNode?: SafeXmlNode;
  defRPr?: SafeXmlNode;
  defRPrs?: SafeXmlNode[];
}

// Later calls override earlier values.
function mergeParagraphProps(target: MergedParagraphStyle, pPr: SafeXmlNode): void {
  if (!pPr.exists()) return;

  const algn = pPr.attr("algn");
  if (algn) target.align = algn;

  const rtl = pPr.attr("rtl");
  if (rtl !== undefined) target.rtl = parseOoxmlBool(rtl);

  const marL = pPr.numAttr("marL");
  if (marL !== undefined) target.marginLeft = emuToPx(marL);

  const indent = pPr.numAttr("indent");
  if (indent !== undefined) target.textIndent = emuToPx(indent);

  const defTabSz = pPr.numAttr("defTabSz");
  if (defTabSz !== undefined) target.defaultTabSize = emuToPx(defTabSz);

  // spcPct 100000 is single spacing. Line height must be unitless: a percentage
  // is resolved once against the parent's font-size and inherited as fixed px,
  // so larger runs inside would overlap.
  const lnSpc = pPr.child("lnSpc");
  if (lnSpc.exists()) {
    const spcPct = lnSpc.child("spcPct");
    if (spcPct.exists()) {
      const val = spcPct.numAttr("val");
      if (val !== undefined) {
        target.lineHeight = `${(val / 100000).toFixed(3)}`;
      }
    }
    const spcPts = lnSpc.child("spcPts");
    if (spcPts.exists()) {
      const val = spcPts.numAttr("val");
      if (val !== undefined) {
        target.lineHeight = `${val / 100}pt`;
        target.lineHeightAbsolute = true;
      }
    }
  }

  const spcBef = pPr.child("spcBef");
  if (spcBef.exists()) {
    const spcPts = spcBef.child("spcPts");
    if (spcPts.exists()) {
      const val = spcPts.numAttr("val");
      if (val !== undefined) target.spaceBefore = val / 100;
    }
    const spcPct = spcBef.child("spcPct");
    if (spcPct.exists()) {
      const val = spcPct.numAttr("val");
      if (val !== undefined) target.spaceBeforePct = val / 100000;
    }
  }

  const spcAft = pPr.child("spcAft");
  if (spcAft.exists()) {
    const spcPts = spcAft.child("spcPts");
    if (spcPts.exists()) {
      const val = spcPts.numAttr("val");
      if (val !== undefined) target.spaceAfter = val / 100;
    }
    const spcPct = spcAft.child("spcPct");
    if (spcPct.exists()) {
      const val = spcPct.numAttr("val");
      if (val !== undefined) target.spaceAfterPct = val / 100000;
    }
  }

  const buChar = pPr.child("buChar");
  if (buChar.exists()) {
    target.bulletChar = buChar.attr("char") || "";
    target.bulletNone = false;
  }
  const buAutoNum = pPr.child("buAutoNum");
  if (buAutoNum.exists()) {
    target.bulletAutoNum = buAutoNum.attr("type") || "arabicPeriod";
    const startAt = buAutoNum.numAttr("startAt");
    if (startAt !== undefined) target.bulletAutoNumStartAt = startAt;
    target.bulletNone = false;
  }
  const buNone = pPr.child("buNone");
  if (buNone.exists()) {
    target.bulletNone = true;
    target.bulletChar = undefined;
    target.bulletAutoNum = undefined;
  }
  const buFont = pPr.child("buFont");
  if (buFont.exists()) {
    target.bulletFont = buFont.attr("typeface");
  }
  const buSzPct = pPr.child("buSzPct");
  if (buSzPct.exists()) {
    const val = buSzPct.numAttr("val");
    if (val !== undefined) {
      target.bulletSizePct = val / 100000;
      target.bulletSizePt = undefined;
    }
  }
  const buSzPts = pPr.child("buSzPts");
  if (buSzPts.exists()) {
    const val = buSzPts.numAttr("val");
    if (val !== undefined) {
      target.bulletSizePt = val / 100;
      target.bulletSizePct = undefined;
    }
  }
  const buSzTx = pPr.child("buSzTx");
  if (buSzTx.exists()) {
    target.bulletSizePct = undefined;
    target.bulletSizePt = undefined;
  }
  // `a:buClrTx` must override inherited bullet colors.
  const buClrTx = pPr.child("buClrTx");
  if (buClrTx.exists()) {
    target.bulletColorFollowsText = true;
    target.bulletColorNode = undefined;
  }
  const buClr = pPr.child("buClr");
  if (buClr.exists()) {
    target.bulletColorNode = buClr;
    target.bulletColorFollowsText = false;
  }

  const defRPr = pPr.child("defRPr");
  if (defRPr.exists()) {
    target.defRPr = defRPr;
    target.defRPrs ??= [];
    target.defRPrs.push(defRPr);
  }
}

interface MergedRunStyle {
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  color?: string;
  fontFamily?: string;
  fontFamilyStack?: string[];
  hlinkClick?: string;
  hlinkSlideIndex?: number;
  hlinkTooltip?: string;
  letterSpacingPt?: number;
  // Minimum font size in points to kern at; 0 always kerns.
  kern?: number;
  cap?: string;
  // In 1/1000 of a percent; positive is superscript.
  baseline?: number;
  textGradientCss?: string;
  textPatternCss?: string;
  highlightColor?: string;
  underlineColor?: string;
  underlineFollowsText?: boolean;
  textNoFill?: boolean;
  textOutlineWidth?: number;
  textOutlineColor?: string;
  // Applied as a mask-image, which fades the whole run.
  textOutlineGradientCss?: string;
  textShadow?: string;
}

function getRunColorKind(rPr: SafeXmlNode | undefined): "none" | "defaultTextScheme" | "explicit" {
  if (!rPr?.exists()) return "none";
  if (rPr.child("gradFill").exists()) return "explicit";
  const solidFill = rPr.child("solidFill");
  if (!solidFill.exists()) return "none";
  const scheme = solidFill.child("schemeClr").attr("val");
  return scheme === "tx1" || scheme === "tx2" ? "defaultTextScheme" : "explicit";
}

function mergeRunProps(target: MergedRunStyle, rPr: SafeXmlNode, ctx: RenderContext): void {
  if (!rPr.exists()) return;

  const sz = rPr.numAttr("sz");
  if (sz !== undefined) target.fontSize = sz / 100; // Hundredths of a point.

  const b = rPr.attr("b");
  if (b !== undefined) target.bold = parseOoxmlBool(b);

  const i = rPr.attr("i");
  if (i !== undefined) target.italic = parseOoxmlBool(i);

  const u = rPr.attr("u");
  if (u !== undefined && u !== "none") target.underline = true;
  if (u === "none") target.underline = false;

  const strike = rPr.attr("strike");
  if (strike !== undefined && strike !== "noStrike") target.strikethrough = true;
  if (strike === "noStrike") target.strikethrough = false;

  const highlight = rPr.child("highlight");
  if (highlight.exists()) {
    target.highlightColor = resolveColorToCss(highlight, ctx);
  }

  const uFill = rPr.child("uFill");
  if (uFill.exists()) {
    const uSolidFill = uFill.child("solidFill");
    if (uSolidFill.exists()) {
      target.underlineColor = resolveColorToCss(uSolidFill, ctx);
      target.underlineFollowsText = false;
    }
  }
  const uFillTx = rPr.child("uFillTx");
  if (uFillTx.exists()) {
    target.underlineFollowsText = true;
    target.underlineColor = undefined;
  }

  const solidFill = rPr.child("solidFill");
  if (solidFill.exists()) {
    delete target.textGradientCss;
    delete target.textPatternCss;
    delete target.textNoFill;
    const { color, alpha } = resolveColor(solidFill, ctx);
    const hex = color.startsWith("#") ? color : `#${color}`;
    if (alpha < 1) {
      const { r, g, b: bl } = hexToRgbInternal(hex);
      target.color = `rgba(${r},${g},${bl},${alpha.toFixed(3)})`;
    } else {
      target.color = hex;
    }
  }
  const gradFill = rPr.child("gradFill");
  if (gradFill.exists()) {
    delete target.color;
    delete target.textPatternCss;
    delete target.textNoFill;
    const css = resolveGradientForText(gradFill, ctx);
    if (css) target.textGradientCss = css;
  }
  const pattFill = rPr.child("pattFill");
  if (pattFill.exists()) {
    delete target.color;
    delete target.textGradientCss;
    delete target.textNoFill;
    const css = resolveFill(rPr, ctx);
    if (css) target.textPatternCss = css;
  }

  // Font family. Office often writes separate Latin/East Asian typefaces in the
  // same run; keep them as a CSS fallback stack so CJK glyphs use the EA face.
  const fontFamilyStack: string[] = [];
  const languageHints = [rPr.attr("lang"), rPr.attr("altLang")];
  for (const fontNodeName of ["latin", "ea", "cs"] as const) {
    const fontNode = rPr.child(fontNodeName);
    if (!fontNode.exists()) continue;
    const typeface = fontNode.attr("typeface");
    if (!typeface) continue;
    fontFamilyStack.push(resolveThemeFont(typeface, ctx, languageHints));
  }
  if (fontFamilyStack.length > 0) {
    target.fontFamily = fontFamilyStack[0];
    target.fontFamilyStack = fontFamilyStack;
  }

  const hlinkClick = rPr.child("hlinkClick");
  if (hlinkClick.exists()) {
    const rId = hlinkClick.attr("id") ?? hlinkClick.attr("r:id");
    const rel = rId ? ctx.slide.rels.get(rId) : undefined;
    const action = hlinkClick.attr("action");
    const slideIndex = resolveSlideNavigationIndex(ctx, action, rel);
    if (slideIndex !== undefined && ctx.onNavigate) {
      target.hlinkSlideIndex = slideIndex;
      target.hlinkTooltip = hlinkClick.attr("tooltip");
      if (target.underline === undefined) target.underline = true;
    } else if (rel && isExternalTargetMode(rel.targetMode) && isAllowedExternalUrl(rel.target)) {
      target.hlinkClick = rel.target;
      if (target.underline === undefined) target.underline = true;
    }
  }

  // `spc` and `kern` are in hundredths of a point.
  const spc = rPr.numAttr("spc");
  if (spc !== undefined) target.letterSpacingPt = spc / 100;

  const kern = rPr.numAttr("kern");
  if (kern !== undefined) target.kern = kern / 100;

  const cap = rPr.attr("cap");
  if (cap !== undefined) target.cap = cap;

  const baseline = rPr.numAttr("baseline");
  if (baseline !== undefined) target.baseline = baseline;

  const effectLst = rPr.child("effectLst");
  const outerShdw = effectLst.child("outerShdw");
  if (outerShdw.exists()) {
    const textOuterShadow = resolveTextOuterShadow(outerShdw, ctx);
    if (textOuterShadow) appendTextShadow(target, textOuterShadow);
  }

  const glow = effectLst.child("glow");
  if (glow.exists()) {
    const textGlowShadow = resolveTextGlowShadow(glow, ctx);
    if (textGlowShadow) appendTextShadow(target, textGlowShadow);
  }

  if (rPr.child("noFill").exists()) {
    delete target.color;
    delete target.textGradientCss;
    delete target.textPatternCss;
    target.textNoFill = true;
  }

  const ln = rPr.child("ln");
  if (ln.exists() && !ln.child("noFill").exists()) {
    const lnW = ln.numAttr("w");
    target.textOutlineWidth = lnW ? emuToPx(lnW) : 0.75;
    const lnSolid = ln.child("solidFill");
    if (lnSolid.exists()) {
      const { color: c, alpha: a } = resolveColor(lnSolid, ctx);
      target.textOutlineColor = colorToCssLocal(c, a);
    }
    const lnGrad = ln.child("gradFill");
    if (lnGrad.exists()) {
      target.textOutlineGradientCss = resolveGradientForText(lnGrad, ctx);
    }
  }
}

function mergeParagraphDefaultRunProps(
  target: MergedRunStyle,
  paragraphStyle: MergedParagraphStyle,
  ctx: RenderContext,
): void {
  for (const defRPr of paragraphStyle.defRPrs ?? []) {
    mergeRunProps(target, defRPr, ctx);
  }
}

function getParagraphDefaultRunStyle(
  paragraphStyle: MergedParagraphStyle,
  ctx: RenderContext,
): MergedRunStyle {
  const style: MergedRunStyle = {};
  mergeParagraphDefaultRunProps(style, paragraphStyle, ctx);
  return style;
}

function hexToRgbInternal(hex: string): { r: number; g: number; b: number } {
  const cleaned = hex.replace(/^#/, "");
  const num = parseInt(
    cleaned.length === 3
      ? cleaned[0] + cleaned[0] + cleaned[1] + cleaned[1] + cleaned[2] + cleaned[2]
      : cleaned,
    16,
  );
  return { r: (num >> 16) & 0xff, g: (num >> 8) & 0xff, b: num & 0xff };
}

function colorToCssLocal(color: string, alpha: number): string {
  const hex = color.startsWith("#") ? color : `#${color}`;
  if (alpha >= 1) return hex;
  const { r, g, b } = hexToRgbInternal(hex);
  return `rgba(${r},${g},${b},${alpha.toFixed(3)})`;
}

function resolveTextGlowShadow(glow: SafeXmlNode, ctx: RenderContext): string | undefined {
  const radiusPx = emuToPx(glow.numAttr("rad") ?? 0);
  if (!(radiusPx > 0)) return undefined;

  const { color, alpha } = resolveColor(glow, ctx);
  if (!color || alpha <= 0) return undefined;

  return `0px 0px ${radiusPx.toFixed(1)}px ${colorToCssLocal(color, alpha)}`;
}

function resolveTextOuterShadow(outerShdw: SafeXmlNode, ctx: RenderContext): string | undefined {
  const distPx = emuToPx(outerShdw.numAttr("dist") ?? 0);
  const blurPx = emuToPx(outerShdw.numAttr("blurRad") ?? 0);
  const dirDeg = (outerShdw.numAttr("dir") ?? 0) / 60000;
  const offsetX = distPx * Math.cos((dirDeg * Math.PI) / 180);
  const offsetY = distPx * Math.sin((dirDeg * Math.PI) / 180);

  const { color, alpha } = resolveColor(outerShdw, ctx);
  if (!color || alpha <= 0) return undefined;

  return `${offsetX.toFixed(1)}px ${offsetY.toFixed(1)}px ${blurPx.toFixed(1)}px ${colorToCssLocal(color, alpha)}`;
}

function appendTextShadow(target: MergedRunStyle, shadow: string): void {
  target.textShadow = target.textShadow ? `${target.textShadow}, ${shadow}` : shadow;
}

function getResolvedRunColor(rPr: SafeXmlNode | undefined, ctx: RenderContext): string | undefined {
  if (!rPr?.exists()) return undefined;
  const style: MergedRunStyle = {};
  mergeRunProps(style, rPr, ctx);
  return style.color;
}

function normalizeCssColorForCompare(color: string | undefined): string | undefined {
  if (!color) return undefined;
  const trimmed = color.trim().toLowerCase();
  if (trimmed.startsWith("#")) {
    const { r, g, b } = hexToRgbInternal(trimmed);
    return `${r},${g},${b},1`;
  }
  const match = trimmed.match(/^rgba?\(([^)]+)\)$/);
  if (!match) return trimmed;
  const parts = match[1].split(",").map((part) => part.trim());
  if (parts.length < 3) return trimmed;
  const [r, g, b] = parts;
  const alpha = parts[3] ?? "1";
  return `${Number(r)},${Number(g)},${Number(b)},${Number(alpha)}`;
}

function colorsEqualForCompare(a: string | undefined, b: string | undefined): boolean {
  const normalizedA = normalizeCssColorForCompare(a);
  const normalizedB = normalizeCssColorForCompare(b);
  return normalizedA !== undefined && normalizedA === normalizedB;
}

function hasSrgbSolidFill(rPr: SafeXmlNode | undefined): boolean {
  return rPr?.child("solidFill").child("srgbClr").exists() ?? false;
}

function runHasHyperlink(rPr: SafeXmlNode | undefined): boolean {
  return rPr?.child("hlinkClick").exists() ?? false;
}

function hasRenderableHyperlink(runStyle: MergedRunStyle): boolean {
  return Boolean(runStyle.hlinkClick) || runStyle.hlinkSlideIndex !== undefined;
}

function hasMatchingNonHyperlinkTextColor(
  run: TextRun,
  paragraph: TextParagraph,
  runColor: string | undefined,
  ctx: RenderContext,
): boolean {
  if (!runColor) return false;
  for (const otherRun of paragraph.runs) {
    if (otherRun === run || runHasHyperlink(otherRun.properties)) continue;
    const otherColor = getResolvedRunColor(otherRun.properties, ctx);
    if (colorsEqualForCompare(runColor, otherColor)) return true;
  }
  const endParaColor = getResolvedRunColor(paragraph.endParaRPr, ctx);
  return colorsEqualForCompare(runColor, endParaColor);
}

function shouldUseHyperlinkThemeColor(
  run: TextRun,
  paragraph: TextParagraph,
  runStyle: MergedRunStyle,
  runColorKind: ReturnType<typeof getRunColorKind>,
  hasExplicitRunColor: boolean,
  ctx: RenderContext,
): boolean {
  if (!hasRenderableHyperlink(runStyle)) return false;
  if (!hasExplicitRunColor || runColorKind === "defaultTextScheme") return true;
  if (runColorKind !== "explicit" || !hasSrgbSolidFill(run.properties)) return false;
  return hasMatchingNonHyperlinkTextColor(run, paragraph, runStyle.color, ctx);
}

function resolveGradientForText(gradFill: SafeXmlNode, ctx: RenderContext): string {
  const gsLst = gradFill.child("gsLst");
  const stops: { position: number; color: string }[] = [];
  for (const gs of gsLst.children("gs")) {
    const pos = gs.numAttr("pos") ?? 0;
    const posPercent = pctToDecimal(pos) * 100;
    const { color, alpha } = resolveColor(gs, ctx);
    stops.push({ position: posPercent, color: colorToCssLocal(color, alpha) });
  }
  if (stops.length === 0) return "";
  stops.sort((a, b) => a.position - b.position);
  const stopsStr = stops.map((s) => `${s.color} ${s.position.toFixed(1)}%`).join(", ");
  const lin = gradFill.child("lin");
  if (lin.exists()) {
    const angle = angleToDeg(lin.numAttr("ang") ?? 0);
    const cssAngle = (angle + 90) % 360;
    return `linear-gradient(${cssAngle.toFixed(1)}deg, ${stopsStr})`;
  }
  return `linear-gradient(180deg, ${stopsStr})`;
}

function applyClippedTextBackground(element: HTMLElement, css: string): void {
  element.style.background = css;
  // jsdom drops a valid multi-layer background shorthand with background-size.
  if (!element.style.background && css.includes(" 0 0 / 8px 8px, ")) {
    const layers = css.split(" 0 0 / 8px 8px, ");
    const backgroundColor = layers[layers.length - 1]?.trim();
    const backgroundImages = layers.slice(0, -1).map((layer) => layer.trim());
    if (backgroundImages.length > 0) {
      element.style.backgroundImage = backgroundImages.join(", ");
      element.style.backgroundSize = backgroundImages.map(() => "8px 8px").join(", ");
    }
    if (backgroundColor) {
      element.style.backgroundColor = backgroundColor;
    }
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (element.style as any).webkitBackgroundClip = "text";
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (element.style as any).backgroundClip = "text";
  element.style.color = "transparent";
}

function generateAutoNumber(type: string, num: number): string {
  switch (type) {
    case "arabicPeriod":
      return `${num}.`;
    case "arabicParenR":
      return `${num})`;
    case "arabicParenBoth":
      return `(${num})`;
    case "arabicPlain":
      return `${num}`;
    case "romanUcPeriod":
      return `${toRoman(num)}.`;
    case "romanLcPeriod":
      return `${toRoman(num).toLowerCase()}.`;
    case "alphaUcPeriod":
      return `${String.fromCharCode(64 + (((num - 1) % 26) + 1))}.`;
    case "alphaLcPeriod":
      return `${String.fromCharCode(96 + (((num - 1) % 26) + 1))}.`;
    case "alphaUcParenR":
      return `${String.fromCharCode(64 + (((num - 1) % 26) + 1))})`;
    case "alphaLcParenR":
      return `${String.fromCharCode(96 + (((num - 1) % 26) + 1))})`;
    default:
      return `${num}.`;
  }
}

function toRoman(num: number): string {
  const vals = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
  const syms = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"];
  let result = "";
  let remaining = num;
  for (let i = 0; i < vals.length; i++) {
    while (remaining >= vals[i]) {
      result += syms[i];
      remaining -= vals[i];
    }
  }
  return result;
}

interface RenderTextBodyOptions {
  cellTextColor?: string;
  // The cellText* values override inherited styles but yield to explicit run rPr.
  cellTextBold?: boolean;
  cellTextItalic?: boolean;
  cellTextFontFamily?: string | string[];
  // A shape style's fontRef color (e.g. SmartArt), yielding only to an explicit run color.
  fontRefColor?: string;
  isVerticalText?: boolean;
  defaultLineHeight?: string;
  trimOuterParagraphSpacing?: boolean;
  compactSingleLineSpacing?: boolean;
}

// Styles merge from presentation defaults through master, layout, and shape lstStyle to paragraph pPr and run rPr.
export function renderTextBody(
  textBody: TextBody,
  placeholder: PlaceholderInfo | undefined,
  ctx: RenderContext,
  container: HTMLElement,
  options?: RenderTextBodyOptions,
): void {
  const category = getPlaceholderCategory(placeholder);

  let fontScale = 1;
  let lnSpcReduction = 0;
  const normAutofit = getEffectiveBodyPrChild(textBody, "normAutofit");
  if (normAutofit?.exists()) {
    const fs = normAutofit.numAttr("fontScale");
    if (fs !== undefined) fontScale = fs / 100000;
    const lsr = normAutofit.numAttr("lnSpcReduction");
    if (lsr !== undefined) lnSpcReduction = lsr / 100000;
  }
  if (fontScale !== 1) container.dataset[PPTX_DATASET.fontScale] = String(fontScale);

  const bulletCounters = new Map<string, number>();
  const visibleParagraphIndexes = textBody.paragraphs
    .map((paragraph, index) => ({
      index,
      visible: paragraph.runs.some((run) => run.text != null && run.text.length > 0),
    }))
    .filter((item) => item.visible)
    .map((item) => item.index);
  const firstVisibleParagraphIndex = visibleParagraphIndexes[0];
  const lastVisibleParagraphIndex = visibleParagraphIndexes[visibleParagraphIndexes.length - 1];
  const singleVisibleParagraph =
    visibleParagraphIndexes.length === 1 ? visibleParagraphIndexes[0] : undefined;

  for (const [paragraphIndex, paragraph] of textBody.paragraphs.entries()) {
    const paraDiv = document.createElement("div");
    paraDiv.dataset[PPTX_DATASET.paragraph] = String(paragraphIndex);
    paraDiv.style.width = "100%";
    paraDiv.style.minWidth = "0px";
    paraDiv.style.maxWidth = "100%";
    paraDiv.style.boxSizing = "border-box";
    paraDiv.style.overflowWrap = "anywhere";
    const level = paragraph.level;
    if (options?.isVerticalText) {
      paraDiv.style.wordBreak = "keep-all";
    }
    const hasLineBreaks = paragraph.runs.some((r) => r.text === "\n");
    const hasVisibleRuns = paragraph.runs.some((r) => r.text != null && r.text.length > 0);
    const isSingleLineAutoFitParagraph =
      options?.compactSingleLineSpacing &&
      paragraphIndex === singleVisibleParagraph &&
      !hasLineBreaks;

    const merged: MergedParagraphStyle = {};

    mergeParagraphProps(merged, findStyleAtLevel(ctx.presentation.defaultTextStyle, level));

    mergeParagraphProps(merged, findStyleAtLevel(ctx.master.defaultTextStyle, level));

    const masterTextStyle =
      category === "title"
        ? ctx.master.textStyles.titleStyle
        : category === "body"
          ? ctx.master.textStyles.bodyStyle
          : ctx.master.textStyles.otherStyle;
    mergeParagraphProps(merged, findStyleAtLevel(masterTextStyle, level));

    if (placeholder) {
      const masterPh = findPlaceholderNode(ctx.master.placeholders, placeholder);
      if (masterPh) {
        const lstStyle = getPlaceholderLstStyle(masterPh);
        mergeParagraphProps(merged, findStyleAtLevel(lstStyle, level));
      }
    }

    if (placeholder) {
      const layoutPh = findPlaceholderNode(
        ctx.layout.placeholders.map((e) => e.node),
        placeholder,
      );
      if (layoutPh) {
        const lstStyle = getPlaceholderLstStyle(layoutPh);
        mergeParagraphProps(merged, findStyleAtLevel(lstStyle, level));
      }
    }

    mergeParagraphProps(merged, findStyleAtLevel(textBody.listStyle, level));

    if (paragraph.properties) {
      mergeParagraphProps(merged, paragraph.properties);
    }

    if (merged.align) {
      const alignMap: Record<string, string> = {
        l: "left",
        ctr: "center",
        r: "right",
        just: "justify",
        justLow: "justify",
        dist: "justify",
        thaiDist: "justify",
      };
      paraDiv.style.textAlign = alignMap[merged.align] || "left";
    }
    if (merged.rtl !== undefined) {
      paraDiv.style.direction = merged.rtl ? "rtl" : "ltr";
    }
    if (merged.marginLeft !== undefined) {
      paraDiv.style.paddingLeft = `${merged.marginLeft}px`;
    }
    if (merged.textIndent !== undefined) {
      paraDiv.style.textIndent = `${merged.textIndent}px`;
    }
    let effectiveLineHeight = merged.lineHeight ?? options?.defaultLineHeight;
    if (effectiveLineHeight) {
      if (lnSpcReduction > 0) {
        const parsed = parseFloat(effectiveLineHeight);
        if (!isNaN(parsed)) {
          if (effectiveLineHeight.includes("pt")) {
            effectiveLineHeight = `${(parsed * (1 - lnSpcReduction)).toFixed(2)}pt`;
          } else {
            effectiveLineHeight = `${(parsed * (1 - lnSpcReduction)).toFixed(3)}`;
          }
        }
      }
      if (options?.isVerticalText && !merged.lineHeightAbsolute) {
        effectiveLineHeight = "1";
      } else if (isSingleLineAutoFitParagraph && merged.lineHeightAbsolute) {
        effectiveLineHeight = "normal";
      }
      paraDiv.style.lineHeight = effectiveLineHeight!;
    }
    // Percentage-based spacing is relative to defRPr or the first run's size.
    let effectiveFontSize = 12;
    const defaultRunStyle = getParagraphDefaultRunStyle(merged, ctx);
    if (defaultRunStyle.fontSize !== undefined) effectiveFontSize = defaultRunStyle.fontSize;
    const firstRunSize =
      paragraph.runs.length > 0 ? paragraph.runs[0].properties?.numAttr("sz") : undefined;
    if (firstRunSize !== undefined) {
      effectiveFontSize = firstRunSize / 100;
    } else if (!hasVisibleRuns && paragraph.endParaRPr) {
      // Empty paragraphs (no runs, or only empty-text runs; common in Google
      // Slides exports) take their line height from endParaRPr, like PowerPoint.
      const sz = paragraph.endParaRPr.numAttr("sz");
      if (sz !== undefined) effectiveFontSize = sz / 100;
    }
    // Line boxes include a strut from the block's own font size, so an inherited page font would overflow small labels.
    paraDiv.style.fontSize = `${effectiveFontSize * fontScale}pt`;

    // Content outside a run span, like text typed into an empty paragraph, inherits
    // this color instead of the page's. Empty paragraphs take it from endParaRPr.
    const paraDefaultStyle: MergedRunStyle = { ...defaultRunStyle };
    if (!hasVisibleRuns && paragraph.endParaRPr) {
      mergeRunProps(paraDefaultStyle, paragraph.endParaRPr, ctx);
    }
    // Toggles only on empty paragraphs: runs set them when on but not when
    // off, so a non-bold run would otherwise inherit a bold paragraph.
    if (!hasVisibleRuns) {
      const endBold = paragraph.endParaRPr?.attr("b") !== undefined;
      const endItalic = paragraph.endParaRPr?.attr("i") !== undefined;
      if (endBold ? paraDefaultStyle.bold : (options?.cellTextBold ?? paraDefaultStyle.bold)) {
        paraDiv.style.fontWeight = "bold";
      }
      if (
        endItalic ? paraDefaultStyle.italic : (options?.cellTextItalic ?? paraDefaultStyle.italic)
      ) {
        paraDiv.style.fontStyle = "italic";
      }
      const decorations: string[] = [];
      if (paraDefaultStyle.underline) decorations.push("underline");
      if (paraDefaultStyle.strikethrough) decorations.push("line-through");
      if (decorations.length > 0) paraDiv.style.textDecoration = decorations.join(" ");
    }
    paraDiv.style.color =
      options?.fontRefColor ?? options?.cellTextColor ?? paraDefaultStyle.color ?? "#000000";
    // Same for font family: without this, text typed into an empty paragraph
    // inherits the page font, then switches to the theme font on commit.
    {
      const paraFont =
        options?.cellTextFontFamily ??
        paraDefaultStyle.fontFamilyStack ??
        paraDefaultStyle.fontFamily ??
        ctx.theme.minorFont.latin ??
        ctx.theme.minorFont.ea;
      if (paraFont) {
        const resolvedFont = Array.isArray(paraFont)
          ? paraFont.map((font) => resolveThemeFont(font, ctx))
          : resolveThemeFont(paraFont, ctx);
        paraDiv.style.fontFamily = cssFontFamilyStack(resolvedFont);
      }
    }

    const trimSpaceBefore =
      options?.trimOuterParagraphSpacing && paragraphIndex === firstVisibleParagraphIndex;
    const trimSpaceAfter =
      options?.trimOuterParagraphSpacing && paragraphIndex === lastVisibleParagraphIndex;

    if (trimSpaceBefore) {
      paraDiv.style.marginTop = "0px";
    } else if (merged.spaceBefore !== undefined) {
      paraDiv.style.marginTop = `${merged.spaceBefore}pt`;
    } else if (merged.spaceBeforePct !== undefined) {
      paraDiv.style.marginTop = `${merged.spaceBeforePct * effectiveFontSize}pt`;
    }
    if (trimSpaceAfter) {
      paraDiv.style.marginBottom = "0px";
    } else if (merged.spaceAfter !== undefined) {
      paraDiv.style.marginBottom = `${merged.spaceAfter}pt`;
    } else if (merged.spaceAfterPct !== undefined) {
      paraDiv.style.marginBottom = `${merged.spaceAfterPct * effectiveFontSize}pt`;
    }

    // PowerPoint shows no bullets on metadata placeholders or on empty paragraphs.
    const suppressBullet =
      !hasVisibleRuns ||
      placeholder?.type === "sldNum" ||
      placeholder?.type === "dt" ||
      placeholder?.type === "ftr" ||
      placeholder?.type === "title" ||
      placeholder?.type === "ctrTitle" ||
      placeholder?.type === "subTitle";
    let bulletPrefix = "";
    if (!suppressBullet && merged.bulletNone !== true) {
      if (merged.bulletChar) {
        bulletPrefix = merged.bulletChar;
      } else if (merged.bulletAutoNum) {
        const counterKey = `${level}:${merged.bulletAutoNum}`;
        const num = merged.bulletAutoNumStartAt ?? bulletCounters.get(counterKey) ?? 1;
        bulletPrefix = generateAutoNumber(merged.bulletAutoNum, num);
        bulletCounters.set(counterKey, num + 1);
      }
    }

    if (bulletPrefix) {
      const bulletSpan = document.createElement("span");
      bulletSpan.dataset[PPTX_DATASET.bullet] = "";
      bulletSpan.contentEditable = "false";
      bulletSpan.textContent = bulletPrefix + " ";
      const marginLeft = merged.marginLeft;
      const textIndent = merged.textIndent;
      const useHangingBulletGutter =
        marginLeft !== undefined && marginLeft > 0 && textIndent !== undefined && textIndent < 0;
      if (useHangingBulletGutter) {
        const markerLeft = Math.max(0, marginLeft + textIndent);
        const markerWidth = Math.max(0, marginLeft - markerLeft);
        paraDiv.style.textIndent = "0px";
        if (merged.align === "ctr" || merged.align === "r") {
          paraDiv.style.paddingLeft = "0px";
          bulletSpan.style.display = "inline-block";
          bulletSpan.style.width = `${markerWidth}px`;
          bulletSpan.style.whiteSpace = "pre";
        } else {
          paraDiv.style.position = "relative";
          bulletSpan.style.position = "absolute";
          bulletSpan.style.left = `${markerLeft}px`;
          bulletSpan.style.top = "0px";
          bulletSpan.style.width = `${markerWidth}px`;
          bulletSpan.style.whiteSpace = "pre";
        }
      }
      if (merged.bulletFont) {
        bulletSpan.style.fontFamily = cssFontFamilyStack(resolveThemeFont(merged.bulletFont, ctx));
      }
      const bulletFontSize = merged.bulletSizePt ?? effectiveFontSize * (merged.bulletSizePct ?? 1);
      bulletSpan.style.fontSize = `${bulletFontSize * fontScale}pt`;
      let bulletColor: string | undefined;
      const firstVisibleRunTextColor = (): string | undefined => {
        const firstVisibleRun = paragraph.runs.find(
          (run) => run.text != null && run.text.length > 0,
        );
        if (!firstVisibleRun) return undefined;
        const runStyle = getParagraphDefaultRunStyle(merged, ctx);
        if (firstVisibleRun.properties) {
          mergeRunProps(runStyle, firstVisibleRun.properties, ctx);
        }
        return (
          runStyle.color ??
          options?.fontRefColor ??
          options?.cellTextColor ??
          (runStyle.textNoFill ? "transparent" : undefined)
        );
      };
      if (merged.bulletColorNode && merged.bulletColorNode.exists()) {
        bulletColor = resolveColorToCss(merged.bulletColorNode, ctx);
      }
      if (bulletColor === undefined) {
        bulletColor = firstVisibleRunTextColor();
      }
      if (bulletColor === undefined && textBody.listStyle) {
        const lstStyleLevel = findStyleAtLevel(textBody.listStyle, level);
        if (lstStyleLevel.exists()) {
          const lstDefRPr = lstStyleLevel.child("defRPr");
          if (lstDefRPr.exists()) {
            const fallbackStyle: MergedRunStyle = {};
            mergeRunProps(fallbackStyle, lstDefRPr, ctx);
            if (fallbackStyle.color !== undefined) {
              bulletColor = fallbackStyle.color;
            }
          }
        }
      }
      bulletSpan.style.color =
        bulletColor ?? options?.fontRefColor ?? options?.cellTextColor ?? "#000000";
      paraDiv.appendChild(bulletSpan);
    }

    // Paragraphs without visible text still occupy one line in PowerPoint; Google
    // Slides exports stack runs with empty <a:t/> as vertical spacing.
    if (!hasVisibleRuns && !hasLineBreaks) {
      const emptyLine = document.createElement("br");
      emptyLine.setAttribute(PPTX_ATTRS.emptyLine, "");
      paraDiv.appendChild(emptyLine);
    }

    if (paragraph.runs.some((r) => r.text?.includes("\t"))) {
      const defaultTabPx = merged.defaultTabSize ?? 96; // OOXML's default 914400 EMU at 96 dpi.
      paraDiv.style.tabSize = `${defaultTabPx}px`;
    }
    // CJK fonts (e.g. Microsoft YaHei) have content areas taller than an absolute
    // line-height, so each line gets a fixed-height block.
    const useLineWrappers = merged.lineHeightAbsolute && hasLineBreaks && effectiveLineHeight;
    let currentLineDiv: HTMLElement | null = null;
    if (useLineWrappers) {
      currentLineDiv = document.createElement("div");
      currentLineDiv.style.height = effectiveLineHeight!;
      currentLineDiv.style.overflow = "visible";
      paraDiv.appendChild(currentLineDiv);
    }

    // CSS collapses spaces at the start of a line but PowerPoint renders them.
    let atLineStart = true;
    for (let runIndex = 0; runIndex < paragraph.runs.length; runIndex++) {
      const run = paragraph.runs[runIndex];
      if (run.text === "\n") {
        atLineStart = true;
        if (useLineWrappers) {
          currentLineDiv = document.createElement("div");
          currentLineDiv.style.height = effectiveLineHeight!;
          currentLineDiv.style.overflow = "visible";
          paraDiv.appendChild(currentLineDiv);
        } else {
          paraDiv.appendChild(document.createElement("br"));
        }
        continue;
      }

      const runStyle: MergedRunStyle = {};

      mergeParagraphDefaultRunProps(runStyle, merged, ctx);

      if (run.properties) {
        mergeRunProps(runStyle, run.properties, ctx);
      }

      // An empty <a:defRPr/> in pPr can overwrite the lstStyle defRPr that carries the color.
      if (runStyle.color === undefined && textBody.listStyle) {
        const lstStyleLevel = findStyleAtLevel(textBody.listStyle, level);
        if (lstStyleLevel.exists()) {
          const lstDefRPr = lstStyleLevel.child("defRPr");
          if (lstDefRPr.exists()) {
            const fallbackStyle: MergedRunStyle = {};
            mergeRunProps(fallbackStyle, lstDefRPr, ctx);
            if (fallbackStyle.color !== undefined) {
              runStyle.color = fallbackStyle.color;
            }
          }
        }
      }

      let element: HTMLElement;
      if (runStyle.hlinkSlideIndex !== undefined) {
        const span = document.createElement("span");
        const slideIndex = runStyle.hlinkSlideIndex;
        span.setAttribute("role", "link");
        span.tabIndex = 0;
        span.title = runStyle.hlinkTooltip || slideJumpTitle(slideIndex);
        span.style.cursor = "pointer";
        span.addEventListener("click", (e) => {
          e.stopPropagation();
          ctx.onNavigate?.({ slideIndex });
        });
        span.addEventListener("keydown", (e) => {
          if (e.key !== "Enter" && e.key !== " ") return;
          e.preventDefault();
          e.stopPropagation();
          ctx.onNavigate?.({ slideIndex });
        });
        element = span;
      } else if (runStyle.hlinkClick) {
        const a = document.createElement("a");
        a.href = runStyle.hlinkClick;
        a.target = "_blank";
        a.rel = "noopener noreferrer";
        element = a;
      } else {
        element = document.createElement("span");
      }

      element.dataset[PPTX_DATASET.run] = String(runIndex);

      if (run.text && run.text.includes("\t")) {
        element.textContent = run.text;
        element.style.whiteSpace = "pre";
      } else if (atLineStart && run.text?.startsWith(" ")) {
        element.textContent = run.text;
        element.style.whiteSpace = "pre-wrap";
      } else if (run.text && / {2}/.test(run.text)) {
        // " &nbsp;" pairs stop browsers from collapsing spaces, while single
        // spaces between words stay stretchable for justify.
        const escaped = run.text
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/ {2}/g, " \u00a0");
        element.innerHTML = escaped;
      } else {
        element.textContent = run.text;
      }
      if (run.text && run.text.length > 0) {
        atLineStart = false;
      }

      const fontSize = runStyle.fontSize || 12;
      element.style.fontSize = `${fontSize * fontScale}pt`;
      const hasExplicitRunBold = run.properties?.attr("b") !== undefined;
      if (hasExplicitRunBold ? runStyle.bold : (options?.cellTextBold ?? runStyle.bold)) {
        element.style.fontWeight = "bold";
      }
      const hasExplicitRunItalic = run.properties?.attr("i") !== undefined;
      if (hasExplicitRunItalic ? runStyle.italic : (options?.cellTextItalic ?? runStyle.italic)) {
        element.style.fontStyle = "italic";
      }

      const decorations: string[] = [];
      if (runStyle.underline) decorations.push("underline");
      if (runStyle.strikethrough) decorations.push("line-through");
      if (decorations.length > 0) {
        element.style.textDecoration = decorations.join(" ");
      }
      if (runStyle.highlightColor) {
        element.style.backgroundColor = runStyle.highlightColor;
      }

      const runColorKind = getRunColorKind(run.properties);
      const hasExplicitRunColor = runColorKind !== "none";
      let effectiveColor: string | undefined;
      if (options?.fontRefColor) {
        effectiveColor = hasExplicitRunColor ? runStyle.color : options.fontRefColor;
      } else if (options?.cellTextColor && !hasExplicitRunColor) {
        effectiveColor = options.cellTextColor;
      } else {
        effectiveColor = runStyle.color;
      }

      // Office writes default text fills on hyperlink runs (tx1/tx2, or the
      // surrounding text's color) but displays them in the theme hlink color.
      if (
        shouldUseHyperlinkThemeColor(
          run,
          paragraph,
          runStyle,
          runColorKind,
          hasExplicitRunColor,
          ctx,
        )
      ) {
        const hlinkHex = ctx.theme.colorScheme.get("hlink");
        if (hlinkHex) {
          effectiveColor = hlinkHex.startsWith("#") ? hlinkHex : `#${hlinkHex}`;
        }
      }

      if (effectiveColor) {
        element.style.color = effectiveColor;
      } else {
        // Keeps text from inheriting the page's color.
        element.style.color = "#000000";
      }
      if (runStyle.underlineFollowsText && effectiveColor) {
        element.style.textDecorationColor = effectiveColor;
      }
      if (runStyle.underlineColor) {
        element.style.textDecorationColor = runStyle.underlineColor;
      }
      if (runStyle.textShadow) {
        element.style.textShadow = runStyle.textShadow;
      }

      if (runStyle.textGradientCss) {
        applyClippedTextBackground(element, runStyle.textGradientCss);
      }
      if (runStyle.textPatternCss) {
        applyClippedTextBackground(element, runStyle.textPatternCss);
      }

      if (runStyle.textNoFill || runStyle.textOutlineWidth) {
        const strokeW = runStyle.textOutlineWidth ?? 0.75;
        if (runStyle.textNoFill && runStyle.textOutlineGradientCss) {
          const outlineColor = "#ffffff"; // The mask applies the gradient.
          element.style.color = "transparent";
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (element.style as any).webkitTextStrokeWidth = `${strokeW}px`;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (element.style as any).webkitTextStrokeColor = outlineColor;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (element.style as any).paintOrder = "stroke fill";
          const maskGrad = runStyle.textOutlineGradientCss;
          element.style.maskImage = maskGrad;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (element.style as any).webkitMaskImage = maskGrad;
        } else if (runStyle.textNoFill && runStyle.textOutlineColor) {
          element.style.color = "transparent";
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (element.style as any).webkitTextStrokeWidth = `${strokeW}px`;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (element.style as any).webkitTextStrokeColor = runStyle.textOutlineColor;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (element.style as any).paintOrder = "stroke fill";
        } else if (runStyle.textNoFill) {
          element.style.color = "transparent";
        } else if (runStyle.textOutlineColor) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (element.style as any).webkitTextStrokeWidth = `${strokeW}px`;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (element.style as any).webkitTextStrokeColor = runStyle.textOutlineColor;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (element.style as any).paintOrder = "stroke fill";
        }
      }

      const hasExplicitRunFont =
        run.properties?.child("latin").exists() ||
        run.properties?.child("ea").exists() ||
        run.properties?.child("cs").exists();
      const effectiveFont = hasExplicitRunFont
        ? (runStyle.fontFamilyStack ?? runStyle.fontFamily)
        : (options?.cellTextFontFamily ?? runStyle.fontFamilyStack ?? runStyle.fontFamily);
      if (effectiveFont) {
        const resolvedFont = Array.isArray(effectiveFont)
          ? effectiveFont.map((font) => resolveThemeFont(font, ctx))
          : resolveThemeFont(effectiveFont, ctx);
        element.style.fontFamily = cssFontFamilyStack(resolvedFont);
      } else {
        const fallback = ctx.theme.minorFont.latin || ctx.theme.minorFont.ea;
        if (fallback) {
          element.style.fontFamily = cssFontFamilyStack(fallback);
        }
      }

      if (runStyle.letterSpacingPt !== undefined) {
        element.style.letterSpacing = `${runStyle.letterSpacingPt}pt`;
      }
      if (runStyle.kern !== undefined) {
        const effectivePt = (runStyle.fontSize || 12) * fontScale;
        element.style.fontKerning = effectivePt >= runStyle.kern ? "normal" : "none";
      }

      if (runStyle.cap === "all") {
        element.style.textTransform = "uppercase";
      } else if (runStyle.cap === "small") {
        element.style.fontVariant = "small-caps";
      }

      if (runStyle.baseline !== undefined && runStyle.baseline !== 0) {
        const shiftPct = runStyle.baseline / 1000;
        element.style.verticalAlign = `${shiftPct}%`;
        if (Math.abs(shiftPct) >= 20) {
          element.style.fontSize = `${fontSize * fontScale * 0.65}pt`;
        }
      }

      const appendTarget = currentLineDiv ?? paraDiv;
      appendTarget.appendChild(element);
    }

    // After a trailing line break, endParaRPr sizes the blank last line; bottom-anchored text depends on it.
    if (paragraph.endParaRPr) {
      const lastRun = paragraph.runs[paragraph.runs.length - 1];
      if (lastRun?.text === "\n") {
        const epSz = paragraph.endParaRPr.numAttr("sz");
        if (epSz !== undefined) {
          const spacer = document.createElement("span");
          spacer.textContent = "\u200B";
          spacer.style.fontSize = `${(epSz / 100) * fontScale}pt`;
          const target = currentLineDiv ?? paraDiv;
          target.appendChild(spacer);
        }
      }
    }

    container.appendChild(paraDiv);
  }
}
