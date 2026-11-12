import { emuToPx } from "../ooxml/unit";
import { SafeXmlNode } from "../ooxml/xml";
import type { PlaceholderEntry } from "./layout";

export interface MasterData {
  colorMap: Map<string, string>;
  background?: SafeXmlNode;
  textStyles: {
    titleStyle?: SafeXmlNode;
    bodyStyle?: SafeXmlNode;
    otherStyle?: SafeXmlNode;
  };
  defaultTextStyle?: SafeXmlNode;
  placeholders: SafeXmlNode[];
  placeholderEntries?: PlaceholderEntry[];
  spTree: SafeXmlNode;
  rels: Map<string, import("../ooxml/rel").RelEntry>;
}

function isPlaceholder(node: SafeXmlNode): boolean {
  for (const wrapper of ["nvSpPr", "nvPicPr", "nvGraphicFramePr", "nvCxnSpPr"]) {
    const nvWrapper = node.child(wrapper);
    if (!nvWrapper.exists()) continue;
    const nvPr = nvWrapper.child("nvPr");
    if (nvPr.child("ph").exists()) {
      return true;
    }
  }
  return false;
}

function getShapeXfrmInEmu(
  node: SafeXmlNode,
): { offX: number; offY: number; cx: number; cy: number } | null {
  const spPrXfrm = node.child("spPr").child("xfrm");
  const xfrm = spPrXfrm.exists() ? spPrXfrm : node.child("xfrm");
  if (!xfrm.exists()) return null;
  const off = xfrm.child("off");
  const ext = xfrm.child("ext");
  const offX = off.numAttr("x") ?? 0;
  const offY = off.numAttr("y") ?? 0;
  const cx = ext.numAttr("cx") ?? 0;
  const cy = ext.numAttr("cy") ?? 0;
  return { offX, offY, cx, cy };
}

function getGroupXfrmInEmu(grpSp: SafeXmlNode): {
  offX: number;
  offY: number;
  cx: number;
  cy: number;
  chOffX: number;
  chOffY: number;
  chExtCx: number;
  chExtCy: number;
} | null {
  const grpSpPr = grpSp.child("grpSpPr");
  if (!grpSpPr.exists()) return null;
  const xfrm = grpSpPr.child("xfrm");
  if (!xfrm.exists()) return null;
  const off = xfrm.child("off");
  const ext = xfrm.child("ext");
  const chOff = xfrm.child("chOff");
  const chExt = xfrm.child("chExt");
  const offX = off.numAttr("x") ?? 0;
  const offY = off.numAttr("y") ?? 0;
  const cx = ext.numAttr("cx") ?? 0;
  const cy = ext.numAttr("cy") ?? 0;
  const chOffX = chOff.exists() ? (chOff.numAttr("x") ?? 0) : 0;
  const chOffY = chOff.exists() ? (chOff.numAttr("y") ?? 0) : 0;
  const chExtCx = chExt.exists() ? (chExt.numAttr("cx") ?? cx) : cx;
  const chExtCy = chExt.exists() ? (chExt.numAttr("cy") ?? cy) : cy;
  return {
    offX,
    offY,
    cx,
    cy,
    chOffX,
    chOffY,
    chExtCx: chExtCx > 0 ? chExtCx : 1,
    chExtCy: chExtCy > 0 ? chExtCy : 1,
  };
}

function extractPlaceholderEntriesRecursive(
  spTree: SafeXmlNode,
  groupTransform: { offX: number; offY: number; scaleX: number; scaleY: number } | null,
): PlaceholderEntry[] {
  const out: PlaceholderEntry[] = [];
  for (const child of spTree.allChildren()) {
    if (child.localName === "grpSp") {
      const gx = getGroupXfrmInEmu(child);
      if (gx && gx.chExtCx > 0 && gx.chExtCy > 0) {
        const scaleX = gx.cx / gx.chExtCx;
        const scaleY = gx.cy / gx.chExtCy;
        const baseOffX = gx.offX - gx.chOffX * scaleX;
        const baseOffY = gx.offY - gx.chOffY * scaleY;
        const nextTransform = groupTransform
          ? {
              offX: groupTransform.offX + baseOffX * groupTransform.scaleX,
              offY: groupTransform.offY + baseOffY * groupTransform.scaleY,
              scaleX: groupTransform.scaleX * scaleX,
              scaleY: groupTransform.scaleY * scaleY,
            }
          : { offX: baseOffX, offY: baseOffY, scaleX, scaleY };
        out.push(...extractPlaceholderEntriesRecursive(child, nextTransform));
      } else {
        out.push(...extractPlaceholderEntriesRecursive(child, groupTransform));
      }
      continue;
    }

    if (!isPlaceholder(child)) continue;
    const sx = getShapeXfrmInEmu(child);
    if (!sx) {
      out.push({ node: child });
      continue;
    }

    if (groupTransform) {
      const absOffX = groupTransform.offX + sx.offX * groupTransform.scaleX;
      const absOffY = groupTransform.offY + sx.offY * groupTransform.scaleY;
      const absCx = sx.cx * groupTransform.scaleX;
      const absCy = sx.cy * groupTransform.scaleY;
      out.push({
        node: child,
        absoluteXfrm: {
          position: { x: emuToPx(absOffX), y: emuToPx(absOffY) },
          size: { w: emuToPx(absCx), h: emuToPx(absCy) },
        },
      });
    } else {
      out.push({
        node: child,
        absoluteXfrm: {
          position: { x: emuToPx(sx.offX), y: emuToPx(sx.offY) },
          size: { w: emuToPx(sx.cx), h: emuToPx(sx.cy) },
        },
      });
    }
  }
  return out;
}

function parseAllAttributes(node: SafeXmlNode): Map<string, string> {
  const result = new Map<string, string>();
  const el = node.element;
  if (!el) return result;
  const attrs = el.attributes;
  for (let i = 0; i < attrs.length; i++) {
    const attr = attrs[i];
    result.set(attr.localName, attr.value);
  }
  return result;
}

export function parseMaster(root: SafeXmlNode): MasterData {
  const cSld = root.child("cSld");

  const bg = cSld.child("bg");
  const background = bg.exists() ? bg : undefined;

  const spTree = cSld.child("spTree");

  const clrMap = root.child("clrMap");
  const colorMap = parseAllAttributes(clrMap);

  const txStyles = root.child("txStyles");
  const titleStyle = txStyles.child("titleStyle");
  const bodyStyle = txStyles.child("bodyStyle");
  const otherStyle = txStyles.child("otherStyle");

  const defaultTextStyle = root.child("defaultTextStyle");

  const placeholderEntries = extractPlaceholderEntriesRecursive(spTree, null);
  const placeholders = placeholderEntries.map((entry) => entry.node);

  return {
    colorMap,
    background,
    textStyles: {
      titleStyle: titleStyle.exists() ? titleStyle : undefined,
      bodyStyle: bodyStyle.exists() ? bodyStyle : undefined,
      otherStyle: otherStyle.exists() ? otherStyle : undefined,
    },
    defaultTextStyle: defaultTextStyle.exists() ? defaultTextStyle : undefined,
    placeholders,
    placeholderEntries,
    spTree,
    rels: new Map(), // Populated by buildPresentation.
  };
}
