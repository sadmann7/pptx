import { angleToDeg, emuToPx, pctToDecimal } from "../ooxml/unit";
import { SafeXmlNode } from "../ooxml/xml";
import {
  applyColorModifiers,
  hexToRgb,
  hslToRgb,
  presetColorToHex,
  rgbToHex,
} from "../utils/color";
import type { ColorModifier } from "../utils/color";
import { RenderContext } from "./context";

function buildColorCacheKey(colorNode: SafeXmlNode): string {
  const parts: string[] = [colorNode.localName, colorNode.attr("val") ?? ""];
  for (const child of colorNode.allChildren()) {
    const tag = child.localName;
    const val = child.attr("val") ?? child.attr("amt");
    if (tag) parts.push(`${tag}:${val ?? ""}`);
    // Wrapper nodes hold their color in a nested child.
    for (const grandchild of child.allChildren()) {
      const gtag = grandchild.localName;
      const gval = grandchild.attr("val") ?? grandchild.attr("amt");
      if (gtag) parts.push(`${gtag}:${gval ?? ""}`);
    }
  }
  return parts.join("|");
}

function collectModifiers(colorNode: SafeXmlNode): ColorModifier[] {
  const modifiers: ColorModifier[] = [];
  for (const child of colorNode.allChildren()) {
    const name = child.localName;
    const val = child.numAttr("val") ?? child.numAttr("amt");
    if (val !== undefined && name) {
      modifiers.push({ name, val });
    } else if (
      name === "inv" ||
      name === "gray" ||
      name === "comp" ||
      name === "gamma" ||
      name === "invGamma"
    ) {
      modifiers.push({ name, val: 0 });
    }
  }
  return modifiers;
}

// The master's colorMap remaps logical names like "tx1" to theme slots like "dk1".
function resolveSchemeColor(schemeName: string, ctx: RenderContext): string {
  let mappedName = schemeName;
  if (ctx.layout.colorMapOverride) {
    const override = ctx.layout.colorMapOverride.get(schemeName);
    if (override) mappedName = override;
  }
  if (mappedName === schemeName) {
    const mapped = ctx.master.colorMap.get(schemeName);
    if (mapped) mappedName = mapped;
  }

  const hex = ctx.theme.colorScheme.get(mappedName);
  if (hex) return hex;

  const fallback = ctx.theme.colorScheme.get(schemeName);
  return fallback || "000000";
}

export function resolveColor(
  colorNode: SafeXmlNode,
  ctx: RenderContext,
): { color: string; alpha: number } {
  const cacheKey = buildColorCacheKey(colorNode);
  const cached = ctx.colorCache.get(cacheKey);
  if (cached) return cached;

  const result = resolveColorUncached(colorNode, ctx);
  ctx.colorCache.set(cacheKey, result);
  return result;
}

function resolveColorUncached(
  colorNode: SafeXmlNode,
  ctx: RenderContext,
  placeholderColorNode?: SafeXmlNode,
): { color: string; alpha: number } {
  for (const child of colorNode.allChildren()) {
    const tag = child.localName;
    const modifiers = collectModifiers(child);

    switch (tag) {
      case "srgbClr": {
        const hex = child.attr("val") || "000000";
        return applyColorModifiers(hex, modifiers);
      }

      case "schemeClr": {
        const scheme = child.attr("val") || "tx1";
        if (scheme.toLowerCase() === "phclr" && placeholderColorNode?.exists()) {
          const base = resolveColor(placeholderColorNode, ctx);
          const baseHex = base.color.startsWith("#") ? base.color.slice(1) : base.color;
          const adjusted = applyColorModifiers(baseHex, modifiers);
          return { color: adjusted.color, alpha: adjusted.alpha * base.alpha };
        }
        const hex = resolveSchemeColor(scheme, ctx);
        return applyColorModifiers(hex, modifiers);
      }

      case "sysClr": {
        const hex = child.attr("lastClr") || child.attr("val") || "000000";
        return applyColorModifiers(hex, modifiers);
      }

      case "prstClr": {
        const name = child.attr("val") || "black";
        const hex = presetColorToHex(name) || "#000000";
        return applyColorModifiers(hex.replace("#", ""), modifiers);
      }

      case "hslClr": {
        const hue = (child.numAttr("hue") ?? 0) / 60000; // 60000ths of a degree.
        const sat = (child.numAttr("sat") ?? 0) / 100000;
        const lum = (child.numAttr("lum") ?? 0) / 100000;
        const rgb = hslToRgb(hue, sat, lum);
        const hex = rgbToHex(rgb.r, rgb.g, rgb.b).replace("#", "");
        return applyColorModifiers(hex, modifiers);
      }

      case "scrgbClr": {
        // scrgbClr channels run 0 to 100000.
        const r = Math.round(((child.numAttr("r") ?? 0) / 100000) * 255);
        const g = Math.round(((child.numAttr("g") ?? 0) / 100000) * 255);
        const b = Math.round(((child.numAttr("b") ?? 0) / 100000) * 255);
        const hex = rgbToHex(r, g, b).replace("#", "");
        return applyColorModifiers(hex, modifiers);
      }

      default:
        break;
    }
  }

  const selfTag = colorNode.localName;
  if (selfTag === "srgbClr") {
    const hex = colorNode.attr("val") || "000000";
    return applyColorModifiers(hex, collectModifiers(colorNode));
  }
  if (selfTag === "schemeClr") {
    const scheme = colorNode.attr("val") || "tx1";
    if (scheme.toLowerCase() === "phclr" && placeholderColorNode?.exists()) {
      const base = resolveColor(placeholderColorNode, ctx);
      const baseHex = base.color.startsWith("#") ? base.color.slice(1) : base.color;
      const adjusted = applyColorModifiers(baseHex, collectModifiers(colorNode));
      return { color: adjusted.color, alpha: adjusted.alpha * base.alpha };
    }
    const hex = resolveSchemeColor(scheme, ctx);
    return applyColorModifiers(hex, collectModifiers(colorNode));
  }
  if (selfTag === "sysClr") {
    const hex = colorNode.attr("lastClr") || colorNode.attr("val") || "000000";
    return applyColorModifiers(hex, collectModifiers(colorNode));
  }
  if (selfTag === "prstClr") {
    const name = colorNode.attr("val") || "black";
    const hex = presetColorToHex(name) || "#000000";
    return applyColorModifiers(hex.replace("#", ""), collectModifiers(colorNode));
  }
  if (selfTag === "hslClr") {
    const hue = (colorNode.numAttr("hue") ?? 0) / 60000;
    const sat = (colorNode.numAttr("sat") ?? 0) / 100000;
    const lum = (colorNode.numAttr("lum") ?? 0) / 100000;
    const rgb = hslToRgb(hue, sat, lum);
    const hex = rgbToHex(rgb.r, rgb.g, rgb.b).replace("#", "");
    return applyColorModifiers(hex, collectModifiers(colorNode));
  }
  if (selfTag === "scrgbClr") {
    const r = Math.round(((colorNode.numAttr("r") ?? 0) / 100000) * 255);
    const g = Math.round(((colorNode.numAttr("g") ?? 0) / 100000) * 255);
    const b = Math.round(((colorNode.numAttr("b") ?? 0) / 100000) * 255);
    const hex = rgbToHex(r, g, b).replace("#", "");
    return applyColorModifiers(hex, collectModifiers(colorNode));
  }

  return { color: "#000000", alpha: 1 };
}

export function resolveColorToCss(node: SafeXmlNode, ctx: RenderContext): string {
  const { color, alpha } = resolveColor(node, ctx);
  return colorToCss(color, alpha);
}

function colorToCss(color: string, alpha: number): string {
  const hex = color.startsWith("#") ? color : `#${color}`;
  const { r, g, b } = hexToRgb(hex);
  if (alpha >= 1) {
    return hex;
  }
  return `rgba(${r},${g},${b},${alpha.toFixed(3)})`;
}

function resolveColorWithPlaceholder(
  colorNode: SafeXmlNode,
  ctx: RenderContext,
  placeholderColorNode?: SafeXmlNode,
): { color: string; alpha: number } {
  if (!placeholderColorNode?.exists()) return resolveColor(colorNode, ctx);
  return resolveColorUncached(colorNode, ctx, placeholderColorNode);
}

// Returns 'transparent' for noFill and '' when the fill is inherited or a blipFill the image renderer paints.
export function resolveFill(spPr: SafeXmlNode, ctx: RenderContext): string {
  const solidFill = spPr.child("solidFill");
  if (solidFill.exists()) {
    const { color, alpha } = resolveColor(solidFill, ctx);
    return colorToCss(color, alpha);
  }

  const gradFill = spPr.child("gradFill");
  if (gradFill.exists()) {
    return resolveGradient(gradFill, ctx);
  }

  const blipFill = spPr.child("blipFill");
  if (blipFill.exists()) {
    return "";
  }

  const pattFill = spPr.child("pattFill");
  if (pattFill.exists()) {
    return resolvePatternFill(pattFill, ctx);
  }

  const grpFill = spPr.child("grpFill");
  if (grpFill.exists()) {
    if (ctx.groupFillNode) {
      return resolveFill(ctx.groupFillNode, ctx);
    }
    return "";
  }

  const noFill = spPr.child("noFill");
  if (noFill.exists()) {
    return "transparent";
  }

  return "";
}

function resolvePatternFill(
  pattFill: SafeXmlNode,
  ctx: RenderContext,
  placeholderColorNode?: SafeXmlNode,
): string {
  const preset = pattFill.attr("prst") ?? "solid";

  let fg = "#000000";
  let bg = "#ffffff";

  const fgClr = pattFill.child("fgClr");
  if (fgClr.exists()) {
    const { color, alpha } = resolveColorWithPlaceholder(fgClr, ctx, placeholderColorNode);
    fg = colorToCss(color, alpha);
  }

  const bgClr = pattFill.child("bgClr");
  if (bgClr.exists()) {
    const { color, alpha } = resolveColorWithPlaceholder(bgClr, ctx, placeholderColorNode);
    bg = colorToCss(color, alpha);
  }

  const s = 8;

  const pat = (gradient: string): string => `${gradient} 0 0 / ${s}px ${s}px, ${bg}`;
  const pat2 = (g1: string, g2: string): string =>
    `${g1} 0 0 / ${s}px ${s}px, ${g2} 0 0 / ${s}px ${s}px, ${bg}`;

  switch (preset) {
    case "solid":
    case "solidDmnd":
      return fg;

    case "pct5":
    case "pct10":
    case "pct20":
    case "pct25":
      return pat(`radial-gradient(${fg} 1px, transparent 1px)`);
    case "pct30":
    case "pct40":
    case "pct50":
      return pat(`radial-gradient(${fg} 1.5px, transparent 1.5px)`);
    case "pct60":
    case "pct70":
    case "pct75":
    case "pct80":
    case "pct90":
      return pat(`radial-gradient(${fg} 2.5px, transparent 2.5px)`);

    case "horz":
    case "ltHorz":
    case "narHorz":
    case "dkHorz":
      return pat(
        `repeating-linear-gradient(0deg, ${fg} 0px, ${fg} 1px, transparent 1px, transparent ${s}px)`,
      );

    case "vert":
    case "ltVert":
    case "narVert":
    case "dkVert":
      return pat(
        `repeating-linear-gradient(90deg, ${fg} 0px, ${fg} 1px, transparent 1px, transparent ${s}px)`,
      );

    case "dnDiag":
    case "ltDnDiag":
    case "narDnDiag":
    case "dkDnDiag":
    case "wdDnDiag":
      return pat(
        `repeating-linear-gradient(45deg, ${fg} 0px, ${fg} 1px, transparent 1px, transparent ${s}px)`,
      );

    case "upDiag":
    case "ltUpDiag":
    case "narUpDiag":
    case "dkUpDiag":
    case "wdUpDiag":
      return pat(
        `repeating-linear-gradient(-45deg, ${fg} 0px, ${fg} 1px, transparent 1px, transparent ${s}px)`,
      );

    case "smGrid":
    case "lgGrid":
    case "cross":
      return pat2(
        `repeating-linear-gradient(0deg, ${fg} 0px, ${fg} 1px, transparent 1px, transparent ${s}px)`,
        `repeating-linear-gradient(90deg, ${fg} 0px, ${fg} 1px, transparent 1px, transparent ${s}px)`,
      );

    case "smCheck":
    case "lgCheck":
    case "diagCross":
    case "openDmnd":
      return pat2(
        `repeating-linear-gradient(45deg, ${fg} 0px, ${fg} 1px, transparent 1px, transparent ${s}px)`,
        `repeating-linear-gradient(-45deg, ${fg} 0px, ${fg} 1px, transparent 1px, transparent ${s}px)`,
      );

    case "dotGrid":
    case "dotDmnd":
      return pat(`radial-gradient(${fg} 1px, transparent 1px)`);

    case "trellis":
    case "weave":
      return pat2(
        `repeating-linear-gradient(45deg, ${fg} 0px, ${fg} 2px, transparent 2px, transparent ${s}px)`,
        `repeating-linear-gradient(-45deg, ${fg} 0px, ${fg} 2px, transparent 2px, transparent ${s}px)`,
      );

    case "dashDnDiag":
    case "dashUpDiag":
    case "dashHorz":
    case "dashVert": {
      const angle = preset.includes("Dn")
        ? "45deg"
        : preset.includes("Up")
          ? "-45deg"
          : preset.includes("Horz")
            ? "0deg"
            : "90deg";
      return pat(
        `repeating-linear-gradient(${angle}, ${fg} 0px, ${fg} 3px, transparent 3px, transparent ${s}px)`,
      );
    }

    case "sphere":
    case "shingle":
    case "plaid":
    case "divot":
    case "zigZag":
      return pat(`radial-gradient(${fg} 2px, transparent 2px)`);

    default:
      return bg;
  }
}

function resolveGradient(
  gradFill: SafeXmlNode,
  ctx: RenderContext,
  placeholderColorNode?: SafeXmlNode,
): string {
  const gsLst = gradFill.child("gsLst");
  const stops: { position: number; color: string }[] = [];

  for (const gs of gsLst.children("gs")) {
    const pos = gs.numAttr("pos") ?? 0;
    const posPercent = pctToDecimal(pos) * 100;
    const { color, alpha } = resolveColorWithPlaceholder(gs, ctx, placeholderColorNode);
    stops.push({ position: posPercent, color: colorToCss(color, alpha) });
  }

  if (stops.length === 0) {
    return "";
  }

  stops.sort((a, b) => a.position - b.position);

  const stopsStr = formatGradientStops(stops);

  const lin = gradFill.child("lin");
  if (lin.exists()) {
    const angle = angleToDeg(lin.numAttr("ang") ?? 0);
    // OOXML angle 0 runs left to right; CSS 0deg runs bottom to top.
    const cssAngle = (angle + 90) % 360;
    return `linear-gradient(${cssAngle.toFixed(1)}deg, ${stopsStr})`;
  }

  const path = gradFill.child("path");
  if (path.exists()) {
    const pathType = path.attr("path");
    if (pathType === "circle" || pathType === "shape" || pathType === "rect") {
      // OOXML path stops run from the fillToRect center to the edge, like CSS radial stops.
      const fillToRect = parseFillToRect(path);
      const { cx, cy } = getFillToRectCenter(fillToRect);
      const pathStopsStr = formatGradientStops(
        getFocusedGradientStops({
          stops,
          cx,
          cy,
          fillToRect,
        }),
      );

      if (pathType === "rect") {
        // Rectangular gradients have cross/X contours (Chebyshev distance) that CSS
        // can't express; approximate with a radial gradient. The SVG path in shape.ts
        // renders the real blend.
        return `radial-gradient(closest-side at ${(cx * 100).toFixed(1)}% ${(cy * 100).toFixed(1)}%, ${pathStopsStr})`;
      }

      return `radial-gradient(ellipse at ${(cx * 100).toFixed(1)}% ${(cy * 100).toFixed(1)}%, ${pathStopsStr})`;
    }
  }

  return `linear-gradient(180deg, ${stopsStr})`;
}

// `lnRef` from `<p:style>` supplies the color when `<a:ln>` has no solidFill, common for connectors.
export function resolveLineStyle(
  ln: SafeXmlNode,
  ctx: RenderContext,
  lnRef?: SafeXmlNode,
): { width: number; color: string; dash: string; dashKind: string } {
  const widthEmu = ln.numAttr("w") ?? 0;
  let width = emuToPx(widthEmu);

  let color = "transparent";
  const solidFill = ln.child("solidFill");
  if (solidFill.exists()) {
    const phClr = solidFill.child("schemeClr");
    const usesPlaceholder = phClr.exists() && (phClr.attr("val") ?? "").toLowerCase() === "phclr";
    if (usesPlaceholder && lnRef && lnRef.exists()) {
      // Theme line styles often use schemeClr=phClr and expect the concrete color from lnRef.
      const base = resolveColor(lnRef, ctx);
      const baseHex = base.color.startsWith("#") ? base.color.slice(1) : base.color;
      const adjusted = applyColorModifiers(baseHex, collectModifiers(phClr));
      color = colorToCss(adjusted.color, adjusted.alpha * base.alpha);
    } else {
      const resolved = resolveColor(solidFill, ctx);
      color = colorToCss(resolved.color, resolved.alpha);
    }
  } else if (lnRef && lnRef.exists() && (lnRef.numAttr("idx") ?? 0) > 0) {
    const idx = lnRef.numAttr("idx") ?? 0;
    if (idx > 0 && ctx.theme.lineStyles && ctx.theme.lineStyles.length >= idx) {
      const themeLn = ctx.theme.lineStyles[idx - 1];
      if (width === 0) {
        const themeW = themeLn.numAttr("w") ?? 0;
        width = emuToPx(themeW);
      }
      const resolved = resolveColor(lnRef, ctx);
      color = colorToCss(resolved.color, resolved.alpha);
    } else {
      const resolved = resolveColor(lnRef, ctx);
      color = colorToCss(resolved.color, resolved.alpha);
      if (width === 0 && idx > 0) {
        width = idx * 0.75; // Approximation without a theme line style.
      }
    }
  }

  // Width fallback should still use lnRef/theme even when explicit solidFill is present on <a:ln>.
  if (width === 0 && lnRef && lnRef.exists()) {
    const idx = lnRef.numAttr("idx") ?? 0;
    if (idx > 0 && ctx.theme.lineStyles && ctx.theme.lineStyles.length >= idx) {
      const themeLn = ctx.theme.lineStyles[idx - 1];
      const themeW = themeLn.numAttr("w") ?? 0;
      width = emuToPx(themeW);
    } else if (idx > 0) {
      width = idx * 0.75;
    }
  }
  if (width === 0 && color !== "transparent" && !ln.child("noFill").exists()) {
    // OOXML defaults an explicit visible line with no `w` to 0.75pt, which is 1px at 96dpi.
    width = 1;
  }

  let dash = "solid";
  let dashKind = "solid";
  const prstDash = ln.child("prstDash");
  if (prstDash.exists()) {
    const val = prstDash.attr("val") || "solid";
    dashKind = val;
    dash = ooxmlDashToCss(val);
  }

  if (dash === "solid" && lnRef && lnRef.exists()) {
    const idx = lnRef.numAttr("idx") ?? 0;
    if (idx > 0 && ctx.theme.lineStyles && ctx.theme.lineStyles.length >= idx) {
      const themeLn = ctx.theme.lineStyles[idx - 1];
      const themeDash = themeLn.child("prstDash");
      if (themeDash.exists()) {
        dashKind = themeDash.attr("val") || "solid";
        dash = ooxmlDashToCss(dashKind);
      }
    }
  }

  return { width, color, dash, dashKind };
}

function ooxmlDashToCss(val: string): string {
  switch (val) {
    case "solid":
      return "solid";
    case "dot":
    case "sysDot":
      return "dotted";
    case "dash":
    case "sysDash":
    case "lgDash":
      return "dashed";
    case "dashDot":
    case "lgDashDot":
    case "lgDashDotDot":
    case "sysDashDot":
    case "sysDashDotDot":
      return "dashed";
    default:
      return "solid";
  }
}

export interface GradientFillData {
  type: "linear" | "radial";
  stops: Array<{ position: number; color: string }>;
  // OOXML gradients visually match linearRGB more closely than sRGB.
  colorInterpolation?: "linearRGB" | "sRGB";
  angle: number;
  cx?: number;
  cy?: number;
  pathType?: string;
  fillToRect?: GradientFocusRect;
}

export interface GradientFocusRect {
  l: number;
  t: number;
  r: number;
  b: number;
}

function parseFillToRect(path: SafeXmlNode): GradientFocusRect | undefined {
  const ftr = path.child("fillToRect");
  if (!ftr.exists()) return undefined;
  return {
    l: (ftr.numAttr("l") ?? 0) / 100000,
    t: (ftr.numAttr("t") ?? 0) / 100000,
    r: (ftr.numAttr("r") ?? 0) / 100000,
    b: (ftr.numAttr("b") ?? 0) / 100000,
  };
}

function getFillToRectCenter(fillToRect: GradientFocusRect | undefined): {
  cx: number;
  cy: number;
} {
  if (!fillToRect) return { cx: 0.5, cy: 0.5 };
  return {
    cx: (fillToRect.l + (1 - fillToRect.r)) / 2,
    cy: (fillToRect.t + (1 - fillToRect.b)) / 2,
  };
}

function clampUnit(value: number): number {
  return Math.max(0, Math.min(1, value));
}

export function getGradientFocusOffset(
  gradientFillData: Pick<GradientFillData, "fillToRect" | "cx" | "cy">,
  options: { axis?: "x" | "y" | "radial"; width?: number; height?: number } = {},
): number {
  const fillToRect = gradientFillData.fillToRect;
  if (!fillToRect) return 0;

  const focusW = Math.max(0, 1 - fillToRect.l - fillToRect.r);
  const focusH = Math.max(0, 1 - fillToRect.t - fillToRect.b);
  if (focusW <= 0 && focusH <= 0) return 0;

  const cx = gradientFillData.cx ?? 0.5;
  const cy = gradientFillData.cy ?? 0.5;
  const maxX = Math.max(Math.abs(cx), Math.abs(1 - cx));
  const maxY = Math.max(Math.abs(cy), Math.abs(1 - cy));

  if (options.axis === "x") {
    return maxX > 0 ? clampUnit(focusW / 2 / maxX) : 0;
  }
  if (options.axis === "y") {
    return maxY > 0 ? clampUnit(focusH / 2 / maxY) : 0;
  }

  const width = options.width ?? 1;
  const height = options.height ?? 1;
  const innerRadius = Math.hypot((focusW / 2) * width, (focusH / 2) * height);
  const outerRadius = Math.hypot(maxX * width, maxY * height);
  return outerRadius > 0 ? clampUnit(innerRadius / outerRadius) : 0;
}

export function getFocusedGradientStops(
  gradientFillData: Pick<GradientFillData, "stops" | "fillToRect" | "cx" | "cy">,
  options: { axis?: "x" | "y" | "radial"; width?: number; height?: number } = {},
): Array<{ position: number; color: string }> {
  const focusOffset = getGradientFocusOffset(gradientFillData, options);
  if (focusOffset <= 0) return gradientFillData.stops;

  return gradientFillData.stops.map((stop) => ({
    ...stop,
    position: focusOffset * 100 + stop.position * (1 - focusOffset),
  }));
}

function formatGradientStops(stops: Array<{ position: number; color: string }>): string {
  return stops.map((s) => `${s.color} ${s.position.toFixed(1)}%`).join(", ");
}

function resolveGradientFillNode(
  gradFill: SafeXmlNode,
  ctx: RenderContext,
  placeholderColorNode?: SafeXmlNode,
): GradientFillData | null {
  const gsLst = gradFill.child("gsLst");
  const stops: Array<{ position: number; color: string }> = [];

  for (const gs of gsLst.children("gs")) {
    const pos = gs.numAttr("pos") ?? 0;
    const posPercent = pctToDecimal(pos) * 100;
    const { color, alpha } = resolveColorWithPlaceholder(gs, ctx, placeholderColorNode);
    stops.push({ position: posPercent, color: colorToCss(color, alpha) });
  }

  if (stops.length === 0) return null;
  stops.sort((a, b) => a.position - b.position);

  const lin = gradFill.child("lin");
  if (lin.exists()) {
    const angle = angleToDeg(lin.numAttr("ang") ?? 0);
    return { type: "linear", stops, angle, colorInterpolation: "linearRGB" };
  }

  const path = gradFill.child("path");
  if (path.exists()) {
    const pathType = path.attr("path");
    if (pathType === "circle" || pathType === "shape" || pathType === "rect") {
      const fillToRect = parseFillToRect(path);
      const { cx, cy } = getFillToRectCenter(fillToRect);
      return {
        type: "radial",
        stops,
        angle: 0,
        cx,
        cy,
        pathType: pathType,
        fillToRect,
        colorInterpolation: "linearRGB",
      };
    }
  }

  return { type: "linear", stops, angle: 0, colorInterpolation: "linearRGB" };
}

export function resolveGradientFill(
  spPr: SafeXmlNode,
  ctx: RenderContext,
): GradientFillData | null {
  let gradFill = spPr.child("gradFill");

  if (!gradFill.exists() && spPr.child("grpFill").exists() && ctx.groupFillNode) {
    gradFill = ctx.groupFillNode.child("gradFill");
  }

  if (!gradFill.exists()) return null;

  return resolveGradientFillNode(gradFill, ctx);
}

export function resolveThemeFillReference(
  fillRef: SafeXmlNode,
  ctx: RenderContext,
): { fillCss: string; gradientFillData: GradientFillData | null } {
  const idx = fillRef.numAttr("idx") ?? 0;
  return resolveThemeFillReferenceFromList(fillRef, ctx, ctx.theme.fillStyles, idx);
}

export function resolveThemeBackgroundFillReference(
  bgRef: SafeXmlNode,
  ctx: RenderContext,
): { fillCss: string; gradientFillData: GradientFillData | null } {
  const idx = bgRef.numAttr("idx") ?? 0;
  if (idx >= 1001) {
    return resolveThemeFillReferenceFromList(bgRef, ctx, ctx.theme.bgFillStyles ?? [], idx - 1000);
  }

  return resolveThemeFillReferenceFromList(bgRef, ctx, ctx.theme.fillStyles, idx);
}

function resolveThemeFillReferenceFromList(
  fillRef: SafeXmlNode,
  ctx: RenderContext,
  fillStyles: SafeXmlNode[] | undefined,
  idx: number,
): { fillCss: string; gradientFillData: GradientFillData | null } {
  if (idx <= 0 || (fillStyles?.length ?? 0) < idx) {
    return { fillCss: resolveColorToCss(fillRef, ctx), gradientFillData: null };
  }

  const themeFill = fillStyles?.[idx - 1];
  if (!themeFill?.exists()) {
    return { fillCss: resolveColorToCss(fillRef, ctx), gradientFillData: null };
  }

  if (themeFill.localName === "solidFill") {
    const resolved = resolveColorWithPlaceholder(themeFill, ctx, fillRef);
    return { fillCss: colorToCss(resolved.color, resolved.alpha), gradientFillData: null };
  }

  if (themeFill.localName === "gradFill") {
    return {
      fillCss: resolveGradient(themeFill, ctx, fillRef),
      gradientFillData: resolveGradientFillNode(themeFill, ctx, fillRef),
    };
  }

  if (themeFill.localName === "pattFill") {
    return { fillCss: resolvePatternFill(themeFill, ctx, fillRef), gradientFillData: null };
  }

  if (themeFill.localName === "noFill") {
    return { fillCss: "transparent", gradientFillData: null };
  }

  return { fillCss: resolveColorToCss(fillRef, ctx), gradientFillData: null };
}

export interface GradientStrokeData {
  stops: Array<{ position: number; color: string }>;
  angle: number;
  width: number;
  colorInterpolation?: "linearRGB" | "sRGB";
}

export function resolveGradientStroke(
  ln: SafeXmlNode,
  ctx: RenderContext,
): GradientStrokeData | null {
  const gradFill = ln.child("gradFill");
  if (!gradFill.exists()) return null;

  const gsLst = gradFill.child("gsLst");
  const stops: Array<{ position: number; color: string }> = [];

  for (const gs of gsLst.children("gs")) {
    const pos = gs.numAttr("pos") ?? 0;
    const posPercent = pctToDecimal(pos) * 100;
    const { color, alpha } = resolveColor(gs, ctx);
    const cssColor = colorToCss(color, alpha);
    stops.push({ position: posPercent, color: cssColor });
  }

  if (stops.length === 0) return null;
  stops.sort((a, b) => a.position - b.position);

  const lin = gradFill.child("lin");
  let angle = 0;
  if (lin.exists()) {
    angle = angleToDeg(lin.numAttr("ang") ?? 0);
  }

  const widthEmu = ln.numAttr("w") ?? 0;
  let width = emuToPx(widthEmu);
  // Keeps a gradient stroke without `w` visible.
  if (width <= 0) width = 1;

  return { stops, angle, width, colorInterpolation: "linearRGB" };
}
