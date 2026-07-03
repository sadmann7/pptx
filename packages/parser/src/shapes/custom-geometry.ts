/**
 * Parse OOXML custom geometry (a:custGeom) into SVG path strings.
 */

import { SafeXmlNode } from "../parser/xml-parser";
import type { GuideDefinition } from "./guide-evaluator";
import { evaluateGuides } from "./guide-evaluator";

function inferPathExtent(pathNode: SafeXmlNode): { w: number; h: number } {
  let maxX = 0;
  let maxY = 0;

  for (const cmd of pathNode.allChildren()) {
    if (cmd.localName === "moveTo" || cmd.localName === "lnTo") {
      const pt = cmd.child("pt");
      maxX = Math.max(maxX, pt.numAttr("x") ?? 0);
      maxY = Math.max(maxY, pt.numAttr("y") ?? 0);
      continue;
    }
    if (cmd.localName === "cubicBezTo" || cmd.localName === "quadBezTo") {
      for (const pt of cmd.children("pt")) {
        maxX = Math.max(maxX, pt.numAttr("x") ?? 0);
        maxY = Math.max(maxY, pt.numAttr("y") ?? 0);
      }
      continue;
    }
    if (cmd.localName === "arcTo") {
      maxX = Math.max(maxX, cmd.numAttr("wR") ?? 0);
      maxY = Math.max(maxY, cmd.numAttr("hR") ?? 0);
    }
  }

  return {
    w: Math.max(1, maxX),
    h: Math.max(1, maxY),
  };
}

/** Collect `<a:gd>` definitions from an avLst/gdLst node in document order. */
function collectGuideDefinitions(listNode: SafeXmlNode): GuideDefinition[] {
  if (!listNode.exists()) return [];
  const defs: GuideDefinition[] = [];
  for (const gd of listNode.children("gd")) {
    const name = gd.attr("name");
    const fmla = gd.attr("fmla");
    if (name && fmla) defs.push({ name, fmla });
  }
  return defs;
}

/**
 * Render a custom geometry element to an SVG path d-attribute string.
 *
 * Coordinates in `a:pt` (and arcTo radii/angles) may be literal numbers or
 * references to guides defined in `a:avLst`/`a:gdLst` (ECMA-376
 * ST_AdjCoordinate / ST_AdjAngle). Guides are evaluated in the shape's EMU
 * coordinate space (`sourceExtent`), while numeric path points live in the
 * local space declared by the path's `w`/`h` attributes.
 *
 * @param custGeom - SafeXmlNode wrapping the `a:custGeom` element
 * @param width - Target width in pixels
 * @param height - Target height in pixels
 * @param sourceExtent - Shape extent in EMU (from `a:xfrm > a:ext`)
 * @returns SVG path d-attribute string
 */
export function renderCustomGeometry(
  custGeom: SafeXmlNode,
  width: number,
  height: number,
  sourceExtent?: { w: number; h: number },
): string {
  const pathLst = custGeom.child("pathLst");
  if (!pathLst.exists()) return "";

  const paths = pathLst.children("path");
  const segments: string[] = [];

  // Guides are defined once per custGeom and evaluated in shape EMU space.
  const guideDefs = [
    ...collectGuideDefinitions(custGeom.child("avLst")),
    ...collectGuideDefinitions(custGeom.child("gdLst")),
  ];
  const guideSpace = {
    w: sourceExtent?.w || 0,
    h: sourceExtent?.h || 0,
  };

  for (const pathNode of paths) {
    const fallbackExtent = inferPathExtent(pathNode);
    const pathW = pathNode.numAttr("w") ?? sourceExtent?.w ?? fallbackExtent.w;
    const pathH = pathNode.numAttr("h") ?? sourceExtent?.h ?? fallbackExtent.h;

    // Without a shape extent, assume guides live in the path's local space.
    const guideW = guideSpace.w || pathW;
    const guideH = guideSpace.h || pathH;
    const guides = guideDefs.length > 0 ? evaluateGuides(guideDefs, guideW, guideH) : null;

    // Guide values (shape EMU space) → path local space conversion factors.
    const guideToPathX = guideW > 0 ? pathW / guideW : 1;
    const guideToPathY = guideH > 0 ? pathH / guideH : 1;

    /** Resolve an ST_AdjCoordinate attribute into path local space. */
    const coord = (raw: string | undefined, axis: "x" | "y"): number => {
      if (raw === undefined) return 0;
      const n = Number(raw);
      if (!Number.isNaN(n)) return n;
      const value = guides?.get(raw) ?? 0;
      return value * (axis === "x" ? guideToPathX : guideToPathY);
    };

    /** Resolve an ST_AdjAngle attribute (no spatial scaling). */
    const angle = (raw: string | undefined): number => {
      if (raw === undefined) return 0;
      const n = Number(raw);
      if (!Number.isNaN(n)) return n;
      return guides?.get(raw) ?? 0;
    };

    const scaleX = pathW > 0 ? width / pathW : 1;
    const scaleY = pathH > 0 ? height / pathH : 1;

    // Track current position for arcTo calculations
    let curX = 0;
    let curY = 0;

    const commands = pathNode.allChildren();
    for (const cmd of commands) {
      switch (cmd.localName) {
        case "moveTo": {
          const pt = cmd.child("pt");
          const x = coord(pt.attr("x"), "x") * scaleX;
          const y = coord(pt.attr("y"), "y") * scaleY;
          segments.push(`M${x},${y}`);
          curX = x;
          curY = y;
          break;
        }

        case "lnTo": {
          const pt = cmd.child("pt");
          const x = coord(pt.attr("x"), "x") * scaleX;
          const y = coord(pt.attr("y"), "y") * scaleY;
          segments.push(`L${x},${y}`);
          curX = x;
          curY = y;
          break;
        }

        case "cubicBezTo": {
          const pts = cmd.children("pt");
          if (pts.length >= 3) {
            const x1 = coord(pts[0].attr("x"), "x") * scaleX;
            const y1 = coord(pts[0].attr("y"), "y") * scaleY;
            const x2 = coord(pts[1].attr("x"), "x") * scaleX;
            const y2 = coord(pts[1].attr("y"), "y") * scaleY;
            const x3 = coord(pts[2].attr("x"), "x") * scaleX;
            const y3 = coord(pts[2].attr("y"), "y") * scaleY;
            segments.push(`C${x1},${y1} ${x2},${y2} ${x3},${y3}`);
            curX = x3;
            curY = y3;
          }
          break;
        }

        case "quadBezTo": {
          const pts = cmd.children("pt");
          if (pts.length >= 2) {
            const x1 = coord(pts[0].attr("x"), "x") * scaleX;
            const y1 = coord(pts[0].attr("y"), "y") * scaleY;
            const x2 = coord(pts[1].attr("x"), "x") * scaleX;
            const y2 = coord(pts[1].attr("y"), "y") * scaleY;
            segments.push(`Q${x1},${y1} ${x2},${y2}`);
            curX = x2;
            curY = y2;
          }
          break;
        }

        case "arcTo": {
          const wRRaw = coord(cmd.attr("wR"), "x");
          const hRRaw = coord(cmd.attr("hR"), "y");
          const wR = wRRaw * scaleX;
          const hR = hRRaw * scaleY;
          const stAngRaw = angle(cmd.attr("stAng"));
          const swAngRaw = angle(cmd.attr("swAng"));

          // OOXML angles are in 60000ths of a degree
          const stAng = stAngRaw / 60000;
          const swAng = swAngRaw / 60000;

          if (wR === 0 || hR === 0 || swAng === 0) {
            // Degenerate arc, skip
            break;
          }

          // OOXML arcTo angles are visual (geometric ray) angles in path coordinate space.
          // Convert to parametric using UNSCALED radii before computing positions.
          const stVisRad = (stAng * Math.PI) / 180;
          const stAngRad = Math.atan2(wRRaw * Math.sin(stVisRad), hRRaw * Math.cos(stVisRad));

          const endVisRad = ((stAng + swAng) * Math.PI) / 180;
          const endAngRad = Math.atan2(wRRaw * Math.sin(endVisRad), hRRaw * Math.cos(endVisRad));

          // Compute center and endpoint in unscaled path space, then scale
          const curXU = curX / scaleX;
          const curYU = curY / scaleY;
          const cx = curXU - wRRaw * Math.cos(stAngRad);
          const cy = curYU - hRRaw * Math.sin(stAngRad);
          const endX = (cx + wRRaw * Math.cos(endAngRad)) * scaleX;
          const endY = (cy + hRRaw * Math.sin(endAngRad)) * scaleY;

          // SVG arc flags
          const largeArc = Math.abs(swAng) > 180 ? 1 : 0;
          const sweep = swAng > 0 ? 1 : 0;

          segments.push(`A${wR},${hR} 0 ${largeArc},${sweep} ${endX},${endY}`);
          curX = endX;
          curY = endY;
          break;
        }

        case "close": {
          segments.push("Z");
          break;
        }

        default:
          // Unknown command, skip
          break;
      }
    }
  }

  return segments.join(" ");
}
