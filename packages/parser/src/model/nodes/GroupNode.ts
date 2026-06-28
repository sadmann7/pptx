import { SafeXmlNode } from "../../parser/XmlParser";
import { parseBaseProps } from "./BaseNode";
import type { BaseNodeData, Position, Size } from "./BaseNode";
import { emuToPx } from "../../parser/units";

export interface GroupNodeData extends BaseNodeData {
  nodeType: "group";
  childOffset: Position;
  childExtent: Size;
  children: SafeXmlNode[];
}

const GROUP_CHILD_TAGS = new Set(["sp", "pic", "grpSp", "graphicFrame", "cxnSp"]);

export function parseGroupNode(grpNode: SafeXmlNode): GroupNodeData {
  const base = parseBaseProps(grpNode);
  const grpSpPr = grpNode.child("grpSpPr");
  const xfrm = grpSpPr.child("xfrm");
  const chOff = xfrm.child("chOff");
  const chExt = xfrm.child("chExt");

  const childOffset: Position = chOff.exists()
    ? { x: emuToPx(chOff.numAttr("x") ?? 0), y: emuToPx(chOff.numAttr("y") ?? 0) }
    : { x: 0, y: 0 };

  const childExtent: Size = (() => {
    if (!chExt.exists()) return { w: base.size.w, h: base.size.h };
    const cx = chExt.numAttr("cx"),
      cy = chExt.numAttr("cy");
    return {
      w: cx !== undefined && cx > 0 ? emuToPx(cx) : base.size.w,
      h: cy !== undefined && cy > 0 ? emuToPx(cy) : base.size.h,
    };
  })();

  const children: SafeXmlNode[] = [];
  for (const child of grpNode.allChildren()) {
    if (GROUP_CHILD_TAGS.has(child.localName)) children.push(child);
  }

  return { ...base, nodeType: "group", childOffset, childExtent, children };
}
