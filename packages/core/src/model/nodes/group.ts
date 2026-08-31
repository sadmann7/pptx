/**
 * Parses p:grpSp (group shape) elements into GroupNodeData,
 * recursively handling nested shapes.
 */

import { emuToPx } from "../../ooxml/unit";
import { SafeXmlNode } from "../../ooxml/xml";
import { BaseNodeData, parseBaseProps, NodePosition, NodeSize } from "./base";

export interface GroupNodeData extends BaseNodeData {
  nodeType: "group";
  childOffset: NodePosition;
  childExtent: NodeSize;
  /** @internal Raw XML nodes, opaque to consumers. Use serializePresentation() for JSON-safe data. */
  children: SafeXmlNode[];
}

/** Tag names of elements that can be children in a group's spTree. */
const GROUP_CHILD_TAGS = new Set(["sp", "pic", "grpSp", "graphicFrame", "cxnSp"]);

/**
 * Parses a group shape XML node (`p:grpSp`) into GroupNodeData.
 */
export function parseGroupNode(grpNode: SafeXmlNode): GroupNodeData {
  const base = parseBaseProps(grpNode);

  // --- Child coordinate space from grpSpPr > a:xfrm ---
  // OOXML: when chOff/chExt omitted, child box equals group box (chOff=0,0, chExt=ext).
  const grpSpPr = grpNode.child("grpSpPr");
  const xfrm = grpSpPr.child("xfrm");
  const chOff = xfrm.child("chOff");
  const chExt = xfrm.child("chExt");

  const childOffset: NodePosition = chOff.exists()
    ? { x: emuToPx(chOff.numAttr("x") ?? 0), y: emuToPx(chOff.numAttr("y") ?? 0) }
    : { x: 0, y: 0 };

  const childExtent: NodeSize = (() => {
    if (!chExt.exists()) return { w: base.size.w, h: base.size.h };
    const cx = chExt.numAttr("cx");
    const cy = chExt.numAttr("cy");
    return {
      w: cx !== undefined && cx > 0 ? emuToPx(cx) : base.size.w,
      h: cy !== undefined && cy > 0 ? emuToPx(cy) : base.size.h,
    };
  })();

  // --- Collect direct child shape nodes ---
  const children: SafeXmlNode[] = [];
  for (const child of grpNode.allChildren()) {
    if (GROUP_CHILD_TAGS.has(child.localName)) {
      children.push(child);
    }
  }

  return {
    ...base,
    nodeType: "group",
    childOffset,
    childExtent,
    children,
  };
}
