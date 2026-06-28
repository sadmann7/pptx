import { SafeXmlNode } from "../../parser/XmlParser";
import { emuToPx, angleToDeg } from "../../parser/units";
import { parseBaseProps } from "./BaseNode";
import type { BaseNodeData } from "./BaseNode";

export interface TextRun {
  text: string;
  properties?: SafeXmlNode;
}

export interface TextParagraph {
  properties?: SafeXmlNode;
  runs: TextRun[];
  level: number;
  endParaRPr?: SafeXmlNode;
}

export interface TextBody {
  bodyProperties?: SafeXmlNode;
  layoutBodyProperties?: SafeXmlNode;
  listStyle?: SafeXmlNode;
  paragraphs: TextParagraph[];
}

export interface LineEndInfo {
  type: string;
  w?: string;
  len?: string;
}
export interface TextBoxBounds {
  x: number;
  y: number;
  w: number;
  h: number;
  rotation?: number;
}

export interface ShapeNodeData extends BaseNodeData {
  nodeType: "shape";
  presetGeometry?: string;
  adjustments: Map<string, number>;
  customGeometry?: SafeXmlNode;
  fill?: SafeXmlNode;
  line?: SafeXmlNode;
  headEnd?: LineEndInfo;
  tailEnd?: LineEndInfo;
  textBody?: TextBody;
  textBoxBounds?: TextBoxBounds;
}

function parseParagraph(pNode: SafeXmlNode): TextParagraph {
  const pPr = pNode.child("pPr");
  const level = pPr.numAttr("lvl") ?? 0;
  const orderedRuns: TextRun[] = [];
  for (const child of pNode.allChildren()) {
    const ln = child.localName;
    if (ln === "r") {
      orderedRuns.push({
        text: child.child("t").text(),
        properties: child.child("rPr").exists() ? child.child("rPr") : undefined,
      });
    } else if (ln === "br") {
      orderedRuns.push({
        text: "\n",
        properties: child.child("rPr").exists() ? child.child("rPr") : undefined,
      });
    } else if (ln === "tab") {
      orderedRuns.push({ text: "\t" });
    } else if (ln === "fld") {
      orderedRuns.push({
        text: child.child("t").text(),
        properties: child.child("rPr").exists() ? child.child("rPr") : undefined,
      });
    }
  }
  const endParaRPrNode = pNode.child("endParaRPr");
  return {
    properties: pPr.exists() ? pPr : undefined,
    runs: orderedRuns,
    level,
    endParaRPr: endParaRPrNode.exists() ? endParaRPrNode : undefined,
  };
}

export function parseTextBody(txBody: SafeXmlNode): TextBody | undefined {
  if (!txBody.exists()) return undefined;
  const bodyPr = txBody.child("bodyPr");
  const lstStyle = txBody.child("lstStyle");
  const paragraphs: TextParagraph[] = [];
  for (const pNode of txBody.children("p")) {
    paragraphs.push(parseParagraph(pNode));
  }
  return {
    bodyProperties: bodyPr.exists() ? bodyPr : undefined,
    listStyle: lstStyle.exists() ? lstStyle : undefined,
    paragraphs,
  };
}

const FILL_TYPES = ["solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "noFill"] as const;
function findFill(spPr: SafeXmlNode): SafeXmlNode | undefined {
  for (const ft of FILL_TYPES) {
    const f = spPr.child(ft);
    if (f.exists()) return f;
  }
  return undefined;
}

function parseAdjustments(avLst: SafeXmlNode): Map<string, number> {
  const adjustments = new Map<string, number>();
  for (const gd of avLst.children("gd")) {
    const name = gd.attr("name");
    const fmla = gd.attr("fmla") ?? "";
    if (!name) continue;
    const match = fmla.match(/val\s+(-?\d+)/);
    if (match) adjustments.set(name, Number(match[1]));
  }
  return adjustments;
}

export function parseShapeNode(spNode: SafeXmlNode): ShapeNodeData {
  const base = parseBaseProps(spNode);
  const spPr = spNode.child("spPr");
  const prstGeom = spPr.child("prstGeom");
  const presetGeometry = prstGeom.attr("prst");
  const adjustments = parseAdjustments(prstGeom.child("avLst"));
  const custGeom = spPr.child("custGeom");
  const customGeometry = custGeom.exists() ? custGeom : undefined;
  const fill = findFill(spPr);
  const ln = spPr.child("ln");
  const line = ln.exists() ? ln : undefined;

  let headEnd: LineEndInfo | undefined;
  let tailEnd: LineEndInfo | undefined;
  if (ln.exists()) {
    const he = ln.child("headEnd");
    if (he.exists()) {
      const t = he.attr("type");
      if (t && t !== "none") headEnd = { type: t, w: he.attr("w"), len: he.attr("len") };
    }
    const te = ln.child("tailEnd");
    if (te.exists()) {
      const t = te.attr("type");
      if (t && t !== "none") tailEnd = { type: t, w: te.attr("w"), len: te.attr("len") };
    }
  }

  const txBody = spNode.child("txBody");
  const textBody = parseTextBody(txBody);

  let textBoxBounds: TextBoxBounds | undefined;
  const txXfrm = spNode.child("txXfrm");
  if (txXfrm.exists()) {
    const txOff = txXfrm.child("off");
    const txExt = txXfrm.child("ext");
    const xfrm = spPr.child("xfrm");
    const off = xfrm.child("off");
    const ext = xfrm.child("ext");
    const shapeW = ext.numAttr("cx") ?? 0;
    const shapeH = ext.numAttr("cy") ?? 0;
    const txW = txExt.numAttr("cx") ?? 0;
    const txH = txExt.numAttr("cy") ?? 0;
    if (shapeW > 0 && shapeH > 0) {
      const shapeX = off.numAttr("x") ?? 0;
      const shapeY = off.numAttr("y") ?? 0;
      const txX = txOff.numAttr("x") ?? 0;
      const txY = txOff.numAttr("y") ?? 0;
      const txRotDeg = angleToDeg(txXfrm.numAttr("rot") ?? 0);
      const localX = txX - shapeX;
      const localY = txY - shapeY;
      const isHalfTurn = Math.abs(Math.round(txRotDeg)) % 360 === 180;
      textBoxBounds = {
        x: emuToPx(isHalfTurn ? shapeW - (localX + txW) : localX),
        y: emuToPx(isHalfTurn ? shapeH - (localY + txH) : localY),
        w: emuToPx(txW),
        h: emuToPx(txH),
        rotation: txRotDeg,
      };
    }
  }

  return {
    ...base,
    nodeType: "shape",
    presetGeometry,
    adjustments,
    customGeometry,
    fill,
    line,
    headEnd,
    tailEnd,
    textBody,
    textBoxBounds,
  };
}
