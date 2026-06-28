import type { Fill, Stroke } from "../types";
import { parseColor, parseGradientStops } from "../color";
import { attr, attrNum, get } from "../xml";
import { angleToDegs, emuToPoints } from "../emu";

/**
 * Parse any fill node (spPr child or similar).
 * Handles: solidFill, gradFill, pattFill, noFill.
 */
export function parseFill(node: unknown): Fill | undefined {
  if (!node || typeof node !== "object") return undefined;
  const n = node as Record<string, unknown>;

  if ("a:noFill" in n || "noFill" in n) {
    return { type: "none" };
  }

  if ("a:solidFill" in n || "solidFill" in n) {
    const sf = (n["a:solidFill"] ?? n["solidFill"]) as Record<string, unknown>;
    const color = parseColor(sf);
    if (!color) return undefined;
    return { type: "solid", color };
  }

  if ("a:gradFill" in n || "gradFill" in n) {
    const gf = (n["a:gradFill"] ?? n["gradFill"]) as Record<string, unknown>;
    const gsLst = get(gf, "a:gsLst", "a:gs") ?? get(gf, "gsLst", "gs");
    const stops = parseGradientStops(gsLst);
    const linNode = get(gf, "a:lin") ?? get(gf, "lin");
    const angle = linNode ? angleToDegs(attr(linNode, "ang")) : undefined;
    return { type: "gradient", stops, angle };
  }

  if ("a:pattFill" in n || "pattFill" in n) {
    const pf = (n["a:pattFill"] ?? n["pattFill"]) as Record<string, unknown>;
    const preset = attr(pf, "prst") ?? "pct5";
    const fgColor = parseColor(get(pf, "a:fgClr") ?? get(pf, "fgClr"));
    const bgColor = parseColor(get(pf, "a:bgClr") ?? get(pf, "bgClr"));
    return { type: "pattern", preset, fgColor, bgColor };
  }

  return undefined;
}

/**
 * Parse <a:ln> (line / stroke) node.
 */
export function parseStroke(lnNode: unknown): Stroke | undefined {
  if (!lnNode || typeof lnNode !== "object") return undefined;
  const n = lnNode as Record<string, unknown>;

  const fill = parseFill(n);
  const widthEmu = attrNum(n, "w");
  const width = widthEmu !== undefined ? emuToPoints(widthEmu) : undefined;

  const prstDash = get(n, "a:prstDash");
  const dashStyle = prstDash ? (attr(prstDash, "val") ?? "solid") : "solid";

  const joinNode = n["a:miter"]
    ? "miter"
    : n["a:bevel"]
      ? "bevel"
      : n["a:round"]
        ? "round"
        : undefined;
  const capNode = attr(n, "cap");

  return {
    fill: fill ?? { type: "none" },
    width,
    dashStyle: dashStyle as Stroke["dashStyle"],
    joinStyle: joinNode as Stroke["joinStyle"],
    capStyle: capNode as Stroke["capStyle"],
  };
}

/**
 * Parse background fill from <p:bg> node.
 */
export function parseBackground(bgNode: unknown): Fill | undefined {
  if (!bgNode || typeof bgNode !== "object") return undefined;
  const n = bgNode as Record<string, unknown>;
  const bgPr = n["p:bgPr"] ?? n["bgPr"];
  if (!bgPr) return undefined;
  return parseFill(bgPr);
}
