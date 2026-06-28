import { SafeXmlNode } from "../../parser/XmlParser";
import { emuToPx, angleToDeg } from "../../parser/units";
import { parseOoxmlBool } from "../../parser/booleans";

export type NodeType = "shape" | "picture" | "table" | "group" | "chart" | "unknown";

export interface Position {
  x: number;
  y: number;
}
export interface Size {
  w: number;
  h: number;
}
export interface PlaceholderInfo {
  type?: string;
  idx?: number;
}
export interface HlinkAction {
  action?: string;
  rId?: string;
  tooltip?: string;
}

export interface BaseNodeData {
  id: string;
  name: string;
  nodeType: NodeType;
  position: Position;
  size: Size;
  rotation: number;
  flipH: boolean;
  flipV: boolean;
  placeholder?: PlaceholderInfo;
  hlinkClick?: HlinkAction;
  source: SafeXmlNode;
}

function findNvProps(node: SafeXmlNode): { cNvPr: SafeXmlNode; nvPr: SafeXmlNode } {
  for (const name of ["nvSpPr", "nvPicPr", "nvGrpSpPr", "nvGraphicFramePr", "nvCxnSpPr"]) {
    const wrapper = node.child(name);
    if (wrapper.exists()) {
      return { cNvPr: wrapper.child("cNvPr"), nvPr: wrapper.child("nvPr") };
    }
  }
  return { cNvPr: node.child("cNvPr"), nvPr: node.child("nvPr") };
}

function findXfrm(node: SafeXmlNode): SafeXmlNode {
  const spPr = node.child("spPr");
  if (spPr.exists()) {
    const x = spPr.child("xfrm");
    if (x.exists()) return x;
  }
  const grpSpPr = node.child("grpSpPr");
  if (grpSpPr.exists()) {
    const x = grpSpPr.child("xfrm");
    if (x.exists()) return x;
  }
  const direct = node.child("xfrm");
  if (direct.exists()) return direct;
  return node.child("__nonexistent__");
}

export function parseBaseProps(spNode: SafeXmlNode): Omit<BaseNodeData, "nodeType"> {
  const { cNvPr, nvPr } = findNvProps(spNode);
  const id = cNvPr.attr("id") ?? "";
  const name = cNvPr.attr("name") ?? "";
  const xfrm = findXfrm(spNode);
  const off = xfrm.child("off");
  const ext = xfrm.child("ext");

  const position: Position = {
    x: emuToPx(off.numAttr("x") ?? 0),
    y: emuToPx(off.numAttr("y") ?? 0),
  };
  const size: Size = { w: emuToPx(ext.numAttr("cx") ?? 0), h: emuToPx(ext.numAttr("cy") ?? 0) };
  const rotation = angleToDeg(xfrm.numAttr("rot") ?? 0);
  const flipH = parseOoxmlBool(xfrm.attr("flipH"));
  const flipV = parseOoxmlBool(xfrm.attr("flipV"));

  const ph = nvPr.child("ph");
  const placeholder = ph.exists() ? { type: ph.attr("type"), idx: ph.numAttr("idx") } : undefined;

  let hlinkClick: HlinkAction | undefined;
  const hlinkNode = cNvPr.child("hlinkClick");
  if (hlinkNode.exists()) {
    hlinkClick = {
      action: hlinkNode.attr("action"),
      rId: hlinkNode.attr("id") ?? hlinkNode.attr("r:id"),
      tooltip: hlinkNode.attr("tooltip"),
    };
  }

  return {
    id,
    name,
    position,
    size,
    rotation,
    flipH,
    flipV,
    placeholder,
    hlinkClick,
    source: spNode,
  };
}
