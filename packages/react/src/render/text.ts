import type {
  BodyProperties,
  ColorMap,
  ParagraphStyle,
  RunStyle,
  ThemeColors,
  ThemeFonts,
} from "@pptx/parser";
import type React from "react";
import { toCSS } from "./color";

// ─── Font resolution ──────────────────────────────────────────────────────────

const FONT_FAMILY_ALIASES: Record<string, string[]> = {
  calibri: ["Calibri", "Aptos", "Arial", "Helvetica", "sans-serif"],
  "calibri light": ["Calibri Light", "Aptos Display", "Aptos", "Arial", "Helvetica", "sans-serif"],
  aptos: ["Aptos", "Arial", "Helvetica", "sans-serif"],
  "aptos display": ["Aptos Display", "Aptos", "Arial", "Helvetica", "sans-serif"],
};

const CSS_GENERIC_FONTS = new Set([
  "serif",
  "sans-serif",
  "monospace",
  "cursive",
  "fantasy",
  "system-ui",
]);

function expandFontAliases(font: string): string[] {
  const key = font.trim().toLowerCase();
  return FONT_FAMILY_ALIASES[key] ?? [font.trim()];
}

function cssFontToken(font: string): string {
  const normalized = font.trim().toLowerCase();
  if (CSS_GENERIC_FONTS.has(normalized)) return normalized;
  return `"${font.trim().replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
}

export function buildFontStack(fonts: string[]): string {
  const expanded = fonts.flatMap(expandFontAliases);
  const seen = new Set<string>();
  const unique = expanded.filter((f) => {
    const key = f.trim().toLowerCase();
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
  const hasGeneric = unique.some((f) => CSS_GENERIC_FONTS.has(f.trim().toLowerCase()));
  if (!hasGeneric) unique.push("sans-serif");
  return unique.map(cssFontToken).join(", ");
}

/**
 * Resolve a fontTheme reference ("major"/"minor") plus explicit font families
 * into a CSS font-family stack. Falls back to theme minor font.
 */
export function resolveFontFamily(
  style: RunStyle,
  defaultRunStyle: RunStyle | undefined,
  themeFonts: ThemeFonts | undefined,
): string | undefined {
  const fontFamily = style.fontFamily ?? defaultRunStyle?.fontFamily;
  const fontEa = style.fontEa ?? defaultRunStyle?.fontEa;
  const fontCs = style.fontCs ?? defaultRunStyle?.fontCs;
  const fontTheme = style.fontTheme ?? defaultRunStyle?.fontTheme;

  const stack: string[] = [];

  if (fontFamily) {
    stack.push(fontFamily);
  } else if (fontTheme && themeFonts) {
    const set = fontTheme === "major" ? themeFonts.majorFont : themeFonts.minorFont;
    stack.push(set.latin);
  }

  if (fontEa) stack.push(fontEa);
  else if (fontTheme && themeFonts) {
    const set = fontTheme === "major" ? themeFonts.majorFont : themeFonts.minorFont;
    if (set.ea) stack.push(set.ea);
  }

  if (fontCs) stack.push(fontCs);

  if (stack.length === 0) return undefined;
  return buildFontStack(stack);
}

// ─── Body / container ────────────────────────────────────────────────────────

export function bodyStyle(
  props: BodyProperties,
  _theme: ThemeColors,
  _colorMap?: ColorMap,
): React.CSSProperties {
  const autofit = props.autofit;
  const grows = autofit === "spAutoFit";
  const css: React.CSSProperties = {
    overflow: grows ? "visible" : "hidden",
    width: "100%",
    height: autofit === "spAutoFit" ? "auto" : "100%",
    minHeight: autofit === "spAutoFit" ? undefined : "100%",
    boxSizing: "border-box",
    display: "flex",
    flexDirection: "column",
  };

  css.paddingLeft = `${props.insetLeft ?? 7.2}pt`;
  css.paddingRight = `${props.insetRight ?? 7.2}pt`;
  css.paddingTop = `${props.insetTop ?? 3.6}pt`;
  css.paddingBottom = `${props.insetBottom ?? 3.6}pt`;

  switch (props.verticalAlignment) {
    case "mid":
      css.justifyContent = "center";
      break;
    case "bottom":
    case "just":
      css.justifyContent = "flex-end";
      break;
    default:
      css.justifyContent = "flex-start";
      break;
  }

  if (props.wrap === "none") css.whiteSpace = "nowrap";

  if (props.columns && props.columns > 1) {
    css.columnCount = props.columns;
    if (props.columnSpacing != null) {
      css.columnGap = `${props.columnSpacing}pt`;
    }
  }

  return css;
}

// ─── Paragraph ────────────────────────────────────────────────────────────────

export function paragraphStyle(
  style: ParagraphStyle,
  _theme: ThemeColors,
  _colorMap?: ColorMap,
  lnSpcReduction?: number,
): React.CSSProperties {
  const css: React.CSSProperties = {
    marginTop: 0,
    marginRight: 0,
    marginBottom: 0,
    marginLeft: 0,
    padding: 0,
    lineHeight: 1,
  };

  if (style.alignment) {
    css.textAlign = style.alignment as React.CSSProperties["textAlign"];
  }

  if (style.marginLeft != null) css.marginLeft = `${style.marginLeft}pt`;
  if (style.marginRight != null) css.marginRight = `${style.marginRight}pt`;
  if (style.indent != null) css.textIndent = `${style.indent}pt`;

  if (style.spaceBefore) {
    css.marginTop =
      style.spaceBefore.unit === "pt"
        ? `${style.spaceBefore.value}pt`
        : `${style.spaceBefore.value}%`;
  }

  if (style.spaceAfter) {
    css.marginBottom =
      style.spaceAfter.unit === "pt" ? `${style.spaceAfter.value}pt` : `${style.spaceAfter.value}%`;
  }

  if (style.lineSpacing) {
    if (style.lineSpacing.unit === "pct") {
      let ratio = style.lineSpacing.value / 100;
      if (lnSpcReduction) ratio *= 1 - lnSpcReduction;
      css.lineHeight = `${ratio.toFixed(3)}`;
    } else {
      let ptVal = style.lineSpacing.value;
      if (lnSpcReduction) ptVal *= 1 - lnSpcReduction;
      css.lineHeight = `${ptVal}pt`;
    }
  }

  return css;
}

// ─── Run ─────────────────────────────────────────────────────────────────────

export function runStyle(
  style: RunStyle,
  theme: ThemeColors,
  defaultRunStyle?: RunStyle,
  colorMap?: ColorMap,
  themeFonts?: ThemeFonts,
  fontScale?: number,
): React.CSSProperties {
  const css: React.CSSProperties = {};
  const scale = fontScale ?? 1;

  const bold = style.bold ?? defaultRunStyle?.bold;
  const italic = style.italic ?? defaultRunStyle?.italic;
  const fontSize = style.fontSize ?? defaultRunStyle?.fontSize;
  const color = style.color ?? defaultRunStyle?.color;
  const underline = style.underline ?? defaultRunStyle?.underline;
  const strikethrough = style.strikethrough ?? defaultRunStyle?.strikethrough;
  const baseline = style.baseline ?? defaultRunStyle?.baseline;
  const letterSpacing = style.letterSpacing ?? defaultRunStyle?.letterSpacing;
  const kern = style.kern ?? defaultRunStyle?.kern;
  const cap = style.cap ?? defaultRunStyle?.cap;
  const highlight = style.highlight ?? defaultRunStyle?.highlight;

  if (bold) css.fontWeight = "bold";
  if (italic) css.fontStyle = "italic";
  if (fontSize != null) css.fontSize = `${fontSize * scale}pt`;
  if (color) css.color = toCSS(color, theme, colorMap);

  const fontCss = resolveFontFamily(style, defaultRunStyle, themeFonts);
  if (fontCss) css.fontFamily = fontCss;

  if (underline && underline !== "none") {
    css.textDecoration = strikethrough ? "underline line-through" : "underline";
  } else if (strikethrough) {
    css.textDecoration = "line-through";
  }

  if (letterSpacing != null) css.letterSpacing = `${letterSpacing}pt`;

  if (kern != null) {
    const effectivePt = (fontSize ?? 12) * scale;
    css.fontKerning = effectivePt >= kern ? "normal" : "none";
  }

  if (cap === "all") css.textTransform = "uppercase";
  else if (cap === "small") css.fontVariant = "small-caps";

  if (highlight) {
    css.backgroundColor = toCSS(highlight, theme, colorMap);
  }

  if (baseline != null) {
    if (baseline > 0) {
      css.verticalAlign = "super";
      css.fontSize = `${((fontSize ?? 12) * scale * 0.65).toFixed(1)}pt`;
    } else if (baseline < 0) {
      css.verticalAlign = "sub";
      css.fontSize = `${((fontSize ?? 12) * scale * 0.65).toFixed(1)}pt`;
    }
  }

  return css;
}
