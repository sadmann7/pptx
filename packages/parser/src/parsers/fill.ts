import type {
  ArrowEnd,
  ArrowEndSize,
  ArrowEndType,
  Effect,
  Fill,
  OuterShadow,
  Stroke,
  Theme,
} from "../types";
import { parseColor, parseGradientStops } from "../color";
import { attr, attrNum, get } from "../xml";
import { angleToDegs, emuToPoints } from "../emu";
import type { PptxZip, Relationship } from "../zip";
import { readMediaAsUrl } from "../zip";

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

  // Image fill — records the relationship ID so callers can resolve asynchronously.
  // The src is temporarily "rId:<id>" and must be resolved by the async parseBackgroundAsync.
  if ("a:blipFill" in n || "blipFill" in n) {
    const bf = (n["a:blipFill"] ?? n["blipFill"]) as Record<string, unknown>;
    const blip = get(bf, "a:blip") as Record<string, unknown> | undefined;
    const rId = blip ? (attr(blip, "r:embed") ?? attr(blip, "embed") ?? "") : "";
    if (rId) return { type: "image", src: `rId:${rId}` };
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

  const headEnd = parseArrowEnd(get(n, "a:headEnd"));
  const tailEnd = parseArrowEnd(get(n, "a:tailEnd"));

  return {
    fill: fill ?? { type: "none" },
    width,
    dashStyle: dashStyle as Stroke["dashStyle"],
    joinStyle: joinNode as Stroke["joinStyle"],
    capStyle: capNode as Stroke["capStyle"],
    ...(headEnd ? { headEnd } : {}),
    ...(tailEnd ? { tailEnd } : {}),
  };
}

/**
 * Parse <a:effectLst> to extract shape effects (outer shadow, etc.)
 */
export function parseEffects(effectLstNode: unknown): Effect[] {
  if (!effectLstNode || typeof effectLstNode !== "object") return [];
  const n = effectLstNode as Record<string, unknown>;
  const effects: Effect[] = [];

  const outerShdw = n["a:outerShdw"] ?? n["outerShdw"];
  if (outerShdw && typeof outerShdw === "object") {
    const sn = outerShdw as Record<string, unknown>;
    const color = parseColor(sn); // shadow node itself may have solidFill child
    const blurRad = attrNum(sn, "blurRad");
    const dist = attrNum(sn, "dist");
    const dir = attrNum(sn, "dir");
    const algn = attr(sn, "algn");
    const shadow: OuterShadow = {
      type: "outerShadow",
      color: color ?? { type: "solid", hex: "000000", alpha: 40 },
      blurRadius: blurRad !== undefined ? emuToPoints(blurRad) : undefined,
      distance: dist !== undefined ? emuToPoints(dist) : undefined,
      direction: dir !== undefined ? dir / 60000 : undefined,
      alignment: algn,
    };
    effects.push(shadow);
  }

  return effects;
}

function parseArrowEnd(node: unknown): ArrowEnd | undefined {
  if (!node || typeof node !== "object") return undefined;
  const type = (attr(node, "type") ?? "none") as ArrowEndType;
  if (type === "none") return undefined;
  return {
    type,
    width: (attr(node, "w") ?? "med") as ArrowEndSize,
    length: (attr(node, "len") ?? "med") as ArrowEndSize,
  };
}

/**
 * Parse background fill from <p:bg> node.
 * Handles p:bgPr (direct fills including blipFill images) and
 * p:bgRef (theme background fill references by index).
 */
export async function parseBackground(
  bgNode: unknown,
  zip?: PptxZip,
  rels?: Map<string, Relationship>,
  theme?: Theme,
): Promise<Fill | undefined> {
  if (!bgNode || typeof bgNode !== "object") return undefined;
  const n = bgNode as Record<string, unknown>;

  // p:bgPr — direct background properties (solid, gradient, image, etc.)
  const bgPr = n["p:bgPr"] ?? n["bgPr"];
  if (bgPr) {
    const fill = parseFill(bgPr);
    if (!fill) return undefined;
    // Resolve rId:XYZ placeholder created by parseFill for blipFill
    if (fill.type === "image" && fill.src.startsWith("rId:") && zip && rels) {
      const rId = fill.src.slice(4);
      const rel = rels.get(rId);
      if (rel) {
        const media = await readMediaAsUrl(zip, rel.target);
        return { type: "image", src: media.src };
      }
      return undefined;
    }
    return fill;
  }

  // p:bgRef — reference to theme background fill by 1-based index (1001 = first).
  // The bgRef element may also carry a color child (schemeClr, srgbClr, etc.)
  // that overrides the "phClr" (placeholder color) in the referenced fill style.
  // In practice, most theme bgFillStyleLst entries are solidFill with phClr,
  // so we treat the bgRef color as the actual solid fill color.
  const bgRef = n["p:bgRef"] ?? n["bgRef"];
  if (bgRef) {
    const bgRefNode = bgRef as Record<string, unknown>;
    const idx = parseInt(attr(bgRefNode, "idx") ?? "0", 10);

    // The bgRef color override (applied when theme fill uses phClr)
    const overrideColor = parseColor(bgRefNode);

    if (idx >= 1001 && theme?.bgFillStyles) {
      const templateFill = theme.bgFillStyles[idx - 1001];
      if (templateFill) {
        // If the template fill uses phClr (placeholder color) or the override
        // color is specified, substitute the override color.
        if (overrideColor && templateFill.type === "solid") {
          return { type: "solid", color: overrideColor };
        }
        // Otherwise return the template fill as-is (gradient, image, etc.)
        return templateFill;
      }
    }

    // Even without bgFillStyles, if there's a color override, return a solid fill.
    if (overrideColor) {
      return { type: "solid", color: overrideColor };
    }
  }

  return undefined;
}
