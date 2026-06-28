import type { Effect, Theme, ThemeColors, ThemeFonts } from "../types";
import { parseEffects } from "./fill";
import { attr, get, toArray } from "../xml";

const FALLBACK_COLORS: ThemeColors = {
  dk1: "000000",
  dk2: "1f3864",
  lt1: "ffffff",
  lt2: "e7e6e6",
  accent1: "4472c4",
  accent2: "ed7d31",
  accent3: "a9d18e",
  accent4: "ffc000",
  accent5: "5b9bd5",
  accent6: "70ad47",
  hlink: "0563c1",
  folHlink: "954f72",
};

/**
 * Parse ppt/theme/theme1.xml (or any theme XML) into a Theme object.
 */
export function parseTheme(themeXml: Record<string, unknown>): Theme {
  const themeEl = get(themeXml, "a:theme") as Record<string, unknown> | undefined;
  const name = attr(themeEl, "name") ?? "Office Theme";

  const fmtScheme = get(themeEl, "a:themeElements");
  const clrScheme = get(fmtScheme, "a:clrScheme") as Record<string, unknown> | undefined;
  const fontScheme = get(fmtScheme, "a:fontScheme") as Record<string, unknown> | undefined;

  const fmtSchemeNode = get(fmtScheme, "a:fmtScheme") as
    | Record<string, unknown>
    | undefined;
  const effectStyles = parseEffectStyleLst(fmtSchemeNode);

  return {
    name,
    colors: parseThemeColors(clrScheme),
    fonts: parseThemeFonts(fontScheme),
    ...(effectStyles.length ? { effectStyles } : {}),
  };
}

function parseThemeColors(clrScheme: Record<string, unknown> | undefined): ThemeColors {
  if (!clrScheme) return FALLBACK_COLORS;

  const resolve = (key: string): string => {
    const node = clrScheme[key];
    if (!node || typeof node !== "object") return "000000";
    const n = node as Record<string, unknown>;
    // Colors in the scheme are stored as <a:srgbClr> or <a:sysClr>
    const srgb = attr(n["a:srgbClr"], "val");
    if (srgb) return srgb.toLowerCase();
    const sysLast = attr(n["a:sysClr"], "lastClr");
    if (sysLast) return sysLast.toLowerCase();
    return "000000";
  };

  return {
    dk1: resolve("a:dk1"),
    dk2: resolve("a:dk2"),
    lt1: resolve("a:lt1"),
    lt2: resolve("a:lt2"),
    accent1: resolve("a:accent1"),
    accent2: resolve("a:accent2"),
    accent3: resolve("a:accent3"),
    accent4: resolve("a:accent4"),
    accent5: resolve("a:accent5"),
    accent6: resolve("a:accent6"),
    hlink: resolve("a:hlink"),
    folHlink: resolve("a:folHlink"),
  };
}

function parseThemeFonts(fontScheme: Record<string, unknown> | undefined): ThemeFonts {
  if (!fontScheme) return { major: "Calibri Light", minor: "Calibri" };

  const majorFonts = get(fontScheme, "a:majorFont") as Record<string, unknown> | undefined;
  const minorFonts = get(fontScheme, "a:minorFont") as Record<string, unknown> | undefined;

  const major = attr(get(majorFonts, "a:latin"), "typeface") ?? "Calibri Light";
  const minor = attr(get(minorFonts, "a:latin"), "typeface") ?? "Calibri";

  return { major, minor };
}

/**
 * Parse effectStyleLst from theme fmtScheme.
 * Returns an array where index i corresponds to effectRef idx=i+1
 * (OOXML effectRef idx is 1-based but stored 0-based here for convenience).
 */
function parseEffectStyleLst(
  fmtSchemeNode: Record<string, unknown> | undefined,
): Effect[][] {
  if (!fmtSchemeNode) return [];
  const effectStyleLst = get(fmtSchemeNode, "a:effectStyleLst") as
    | Record<string, unknown>
    | undefined;
  if (!effectStyleLst) return [];
  const styleNodes = toArray(
    effectStyleLst["a:effectStyle"] as unknown[],
  );
  return styleNodes.map((styleNode) => {
    const sn = styleNode as Record<string, unknown>;
    const effectLst = sn["a:effectLst"] ?? sn["effectLst"];
    return parseEffects(effectLst);
  });
}
