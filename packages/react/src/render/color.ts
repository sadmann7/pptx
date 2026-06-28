import type { Color, ColorMap, Effect, Fill, Stroke, ThemeColors } from "@pptx/parser";
import { resolveColor } from "@pptx/parser";

export function toCSS(color: Color | undefined, theme: ThemeColors, colorMap?: ColorMap): string {
  if (!color) return "transparent";
  return resolveColor(color, theme, colorMap);
}

export function fillToCSS(fill: Fill | undefined, theme: ThemeColors, colorMap?: ColorMap): string {
  if (!fill || fill.type === "none") return "transparent";
  if (fill.type === "solid") return toCSS(fill.color, theme, colorMap);
  if (fill.type === "gradient") {
    const stops = fill.stops
      .map((s) => `${toCSS(s.color, theme, colorMap)} ${(s.position * 100).toFixed(1)}%`)
      .join(", ");
    const angle = fill.angle ?? 0;
    // CSS gradient angle: 0° = bottom-to-top; OOXML: 0° = left-to-right
    const cssAngle = (90 - angle + 360) % 360;
    return `linear-gradient(${cssAngle}deg, ${stops})`;
  }
  // pattern falls back to foreground color
  if (fill.type === "pattern" && fill.fgColor) return toCSS(fill.fgColor, theme, colorMap);
  // image fills are handled separately via backgroundImage CSS property
  if (fill.type === "image") return "transparent";
  return "transparent";
}

/** Returns a CSS background-image value for image fills, or undefined for others. */
export function fillToBackgroundImage(fill: Fill | undefined): string | undefined {
  if (fill?.type === "image" && fill.src && !fill.src.startsWith("rId:")) {
    return `url("${fill.src}")`;
  }
  return undefined;
}

/** SVG fill attribute value */
export function fillToSVG(fill: Fill | undefined, theme: ThemeColors, colorMap?: ColorMap): string {
  if (!fill || fill.type === "none") return "none";
  if (fill.type === "solid") return toCSS(fill.color, theme, colorMap);
  // SVG gradient would require defs — fall back to first stop for now
  if (fill.type === "gradient" && fill.stops.length > 0) {
    return toCSS(fill.stops[0]!.color, theme, colorMap);
  }
  if (fill.type === "pattern" && fill.fgColor) return toCSS(fill.fgColor, theme, colorMap);
  return "none";
}

const PT_TO_PX = 96 / 72;

/**
 * Convert an array of element effects to a CSS `box-shadow` value.
 * Outer shadows → `box-shadow: dx dy blur color`.
 */
export function effectsToBoxShadow(
  effects: Effect[] | undefined,
  theme: ThemeColors,
  colorMap?: ColorMap,
): string | undefined {
  if (!effects?.length) return undefined;
  const parts: string[] = [];
  for (const e of effects) {
    if (e.type !== "outerShadow") continue;
    const color = toCSS(e.color, theme, colorMap);
    const dist = (e.distance ?? 0) * PT_TO_PX;
    const rad = ((e.direction ?? 0) * Math.PI) / 180;
    const dx = (dist * Math.cos(rad)).toFixed(1);
    const dy = (dist * Math.sin(rad)).toFixed(1);
    const blur = ((e.blurRadius ?? 0) * PT_TO_PX).toFixed(1);
    parts.push(`${dx}px ${dy}px ${blur}px ${color}`);
  }
  return parts.length ? parts.join(", ") : undefined;
}

/**
 * Convert an array of element effects to a CSS `filter: drop-shadow()` string.
 * Useful for non-rectangular shapes where box-shadow won't follow the outline.
 */
export function effectsToFilter(
  effects: Effect[] | undefined,
  theme: ThemeColors,
  colorMap?: ColorMap,
): string | undefined {
  if (!effects?.length) return undefined;
  const parts: string[] = [];
  for (const e of effects) {
    if (e.type !== "outerShadow") continue;
    const color = toCSS(e.color, theme, colorMap);
    const dist = (e.distance ?? 0) * PT_TO_PX;
    const rad = ((e.direction ?? 0) * Math.PI) / 180;
    const dx = (dist * Math.cos(rad)).toFixed(1);
    const dy = (dist * Math.sin(rad)).toFixed(1);
    const blur = ((e.blurRadius ?? 0) * PT_TO_PX).toFixed(1);
    parts.push(`drop-shadow(${dx}px ${dy}px ${blur}px ${color})`);
  }
  return parts.length ? parts.join(" ") : undefined;
}

/** SVG stroke attributes */
export function strokeToSVGAttrs(
  stroke: Stroke | undefined,
  theme: ThemeColors,
  colorMap?: ColorMap,
): {
  stroke: string;
  strokeWidth: string;
  strokeDasharray?: string;
} {
  if (!stroke || stroke.fill.type === "none") {
    return { stroke: "none", strokeWidth: "0" };
  }

  const color = fillToSVG(stroke.fill, theme, colorMap);
  const width = stroke.width != null ? `${stroke.width}pt` : "1pt";

  let strokeDasharray: string | undefined;
  switch (stroke.dashStyle) {
    case "dot":
      strokeDasharray = "2 2";
      break;
    case "dash":
      strokeDasharray = "6 2";
      break;
    case "dashDot":
      strokeDasharray = "6 2 2 2";
      break;
    case "lgDash":
      strokeDasharray = "12 4";
      break;
    case "lgDashDot":
      strokeDasharray = "12 4 2 4";
      break;
  }

  return { stroke: color, strokeWidth: width, strokeDasharray };
}
