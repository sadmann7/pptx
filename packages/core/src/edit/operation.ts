/**
 * Edit operations for presentations opened with `keepPackage: true`.
 *
 * Every operation follows the same contract:
 *   1. mutate the live part XML (the same documents `writePptx()` saves),
 *   2. update the typed model so renderers and consumers stay in sync,
 *   3. mark the touched parts dirty on the retained package.
 *
 * `applyEdit()` returns an `undo()` closure that restores both the XML and
 * the model. Undos capture element references and part snapshots, so they
 * must be applied in reverse order (a standard LIFO undo stack).
 */

import type { NodePosition, NodeSize } from "../model/nodes/base";
import type { GroupNodeData } from "../model/nodes/group";
import type { ShapeNodeData, TextParagraph, TextRun } from "../model/nodes/shape";
import { materializeSlide, PresentationData } from "../model/presentation";
import { parseSlide, SlideData, SlideNode } from "../model/slide";
import type { PptxPackage } from "../ooxml/package";
import { parseRels, RelEntry, resolveRelTarget } from "../ooxml/rel";
import { degToAngle, pxToEmu } from "../ooxml/unit";
import { parseXml, SafeXmlNode } from "../ooxml/xml";
import {
  A_NS,
  CT_NS,
  insertChild,
  P_NS,
  R_NS,
  RELS_NS,
  removeChild,
  serializePartText,
  setOrRemoveAttr,
} from "./xml-mutation";

// ---------------------------------------------------------------------------
// Operation types
// ---------------------------------------------------------------------------

/** Replace the text of one run inside a shape's text body. */
export interface SetTextRunOperation {
  type: "setTextRun";
  slideId: string;
  nodeId: string;
  paragraphIndex: number;
  runIndex: number;
  text: string;
}

/** Move / resize / rotate / flip a top-level slide node. */
export interface SetNodeTransformOperation {
  type: "setNodeTransform";
  slideId: string;
  nodeId: string;
  /** New position in px (96 DPI slide space). */
  position?: NodePosition;
  /** New size in px. */
  size?: NodeSize;
  /** New rotation in degrees. */
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
}

/** Replace a shape's fill with a solid color. */
export interface SetSolidFillOperation {
  type: "setSolidFill";
  slideId: string;
  nodeId: string;
  /** Hex color, with or without leading `#` (e.g. "FF0000"). */
  color: string;
}

/** Remove a top-level node from a slide. */
export interface DeleteNodeOperation {
  type: "deleteNode";
  slideId: string;
  nodeId: string;
}

/** Move a slide to a new position in the deck. */
export interface MoveSlideOperation {
  type: "moveSlide";
  slideId: string;
  toIndex: number;
}

/** Duplicate a slide; the copy is inserted directly after the source. */
export interface DuplicateSlideOperation {
  type: "duplicateSlide";
  slideId: string;
}

/** Remove a slide from the deck. */
export interface DeleteSlideOperation {
  type: "deleteSlide";
  slideId: string;
}

/** Replace the entire text body of a shape. */
export interface SetTextBodyOperation {
  type: "setTextBody";
  slideId: string;
  nodeId: string;
  paragraphs: SetTextBodyParagraph[];
}

export interface SetTextBodyParagraph {
  /** Index of the source paragraph whose `a:pPr` / `a:endParaRPr` to clone. -1 = bare paragraph. */
  sourceParagraphIndex: number;
  runs: SetTextBodyRun[];
}

export interface SetTextBodyRun {
  text: string;
  /** Source `[paragraphIndex, runIndex]` to copy `a:rPr` from. When omitted the
   *  first run of the source paragraph (or no rPr at all) is used. */
  sourceRun?: [number, number];
}

/**
 * Apply several operations as one atomic edit with a single undo (e.g.
 * moving a multi-selection). Sub-operations are applied in order and undone
 * in reverse. If one fails midway, the already-applied ones are rolled back.
 */
export interface BatchOperation {
  type: "batch";
  operations: EditOperation[];
}

export type EditOperation =
  | SetTextRunOperation
  | SetNodeTransformOperation
  | SetSolidFillOperation
  | DeleteNodeOperation
  | MoveSlideOperation
  | DuplicateSlideOperation
  | DeleteSlideOperation
  | SetTextBodyOperation
  | BatchOperation;

export interface EditResult {
  /** Slide ids whose rendered output may have changed. */
  affectedSlideIds: string[];
  /** Restores XML, model, and dirty state. Apply undos in reverse order. */
  undo: () => void;
  /** Id of the slide created by `duplicateSlide`. */
  createdSlideId?: string;
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

export async function applyEdit(
  presentation: PresentationData,
  op: EditOperation,
): Promise<EditResult> {
  switch (op.type) {
    case "setTextRun":
      return applySetTextRun(presentation, op);
    case "setTextBody":
      return applySetTextBody(presentation, op);
    case "setNodeTransform":
      return applySetNodeTransform(presentation, op);
    case "setSolidFill":
      return applySetSolidFill(presentation, op);
    case "deleteNode":
      return applyDeleteNode(presentation, op);
    case "moveSlide":
      return applyMoveSlide(presentation, op);
    case "duplicateSlide":
      return applyDuplicateSlide(presentation, op);
    case "deleteSlide":
      return applyDeleteSlide(presentation, op);
    case "batch":
      return applyBatch(presentation, op);
  }
}

async function applyBatch(presentation: PresentationData, op: BatchOperation): Promise<EditResult> {
  const results: EditResult[] = [];
  try {
    for (const subOp of op.operations) {
      results.push(await applyEdit(presentation, subOp));
    }
  } catch (error) {
    for (const result of results.reverse()) result.undo();
    throw error;
  }

  const affected = new Set<string>();
  for (const result of results) {
    for (const slideId of result.affectedSlideIds) affected.add(slideId);
  }

  return {
    affectedSlideIds: [...affected],
    undo: () => {
      for (let i = results.length - 1; i >= 0; i--) results[i].undo();
    },
  };
}

// ---------------------------------------------------------------------------
// Shared lookup helpers
// ---------------------------------------------------------------------------

const PRESENTATION_PATH = "ppt/presentation.xml";
const PRESENTATION_RELS_PATH = "ppt/_rels/presentation.xml.rels";
const CONTENT_TYPES_PATH = "[Content_Types].xml";
const SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";

function requirePkg(pres: PresentationData): PptxPackage {
  if (!pres.sourcePackage) {
    throw new Error(
      "applyEdit: presentation was parsed without package retention. " +
        "Parse the zip with { keepPackage: true } to enable editing.",
    );
  }
  return pres.sourcePackage;
}

function findSlide(pres: PresentationData, slideId: string): SlideData {
  const slide = pres.slides.find((s) => s.id === slideId);
  if (!slide) throw new Error(`applyEdit: unknown slide "${slideId}"`);
  return slide;
}

function findNode(pres: PresentationData, slide: SlideData, nodeId: string): SlideNode {
  materializeSlide(pres, slide);
  const node = slide.nodes.find((n) => n.id === nodeId);
  if (!node) {
    throw new Error(`applyEdit: no top-level node "${nodeId}" on slide "${slide.id}"`);
  }
  return node;
}

function requireElement(node: SafeXmlNode, what: string): Element {
  const el = node.element;
  if (!el) throw new Error(`applyEdit: ${what} is missing`);
  return el;
}

/** Update slide indexes and the index-keyed slide→layout map after reordering. */
function reindexSlides(pres: PresentationData): void {
  pres.slideToLayout.clear();
  for (let i = 0; i < pres.slides.length; i++) {
    const slide = pres.slides[i];
    slide.index = i;
    if (slide.layoutIndex) {
      pres.slideToLayout.set(i, slide.layoutIndex);
    }
  }
}

// ---------------------------------------------------------------------------
// setTextRun
// ---------------------------------------------------------------------------

/** Localnames that produce entries in TextParagraph.runs, in document order. */
const RUN_LOCAL_NAMES = new Set(["r", "br", "tab", "fld"]);

function applySetTextRun(pres: PresentationData, op: SetTextRunOperation): EditResult {
  const sourcePackage = requirePkg(pres);
  const slide = findSlide(pres, op.slideId);
  const node = findNode(pres, slide, op.nodeId);

  if (node.nodeType !== "shape" || !node.textBody) {
    throw new Error(`applyEdit: node "${op.nodeId}" has no editable text body`);
  }
  const shape = node as ShapeNodeData;

  const paragraph = shape.textBody?.paragraphs[op.paragraphIndex];
  const run = paragraph?.runs[op.runIndex];
  if (!paragraph || !run) {
    throw new Error(
      `applyEdit: no run at paragraph ${op.paragraphIndex}, run ${op.runIndex} in node "${op.nodeId}"`,
    );
  }

  // The model's runs array is built from an in-order scan of r/br/tab/fld
  // children, so runIndex maps onto the same scan of the XML.
  const pNode = node.source.child("txBody").children("p")[op.paragraphIndex];
  const runNodes = (pNode?.allChildren() ?? []).filter((c) => RUN_LOCAL_NAMES.has(c.localName));
  const runNode = runNodes[op.runIndex];
  if (!runNode) {
    throw new Error(`applyEdit: run XML not found for paragraph ${op.paragraphIndex}`);
  }
  if (runNode.localName !== "r" && runNode.localName !== "fld") {
    throw new Error(`applyEdit: cannot set text on a <a:${runNode.localName}> (break/tab) run`);
  }

  const runEl = requireElement(runNode, "run element");
  let tEl = runNode.child("t").element;
  if (!tEl) {
    tEl = runEl.ownerDocument.createElementNS(A_NS, "a:t");
    insertChild(runEl, tEl, null);
  }

  const prevText = run.text;
  tEl.textContent = op.text;
  run.text = op.text;
  sourcePackage.markDirty(slide.id);

  return {
    affectedSlideIds: [slide.id],
    undo: () => {
      tEl.textContent = prevText;
      run.text = prevText;
      sourcePackage.markDirty(slide.id);
    },
  };
}

// ---------------------------------------------------------------------------
// setTextBody
// ---------------------------------------------------------------------------

function applySetTextBody(pres: PresentationData, op: SetTextBodyOperation): EditResult {
  const sourcePackage = requirePkg(pres);
  const slide = findSlide(pres, op.slideId);
  const node = findNode(pres, slide, op.nodeId);

  if (node.nodeType !== "shape" || !node.textBody) {
    throw new Error(`applyEdit: node "${op.nodeId}" has no editable text body`);
  }
  const shape = node as ShapeNodeData;
  const textBody = shape.textBody!;

  // Snapshot the current state for undo.
  const prevParagraphs = textBody.paragraphs;

  // XML: <p:txBody> element containing the live <a:p> elements.
  const txBodyNode = node.source.child("txBody");
  const txBodyEl = requireElement(txBodyNode, "txBody");
  const doc = txBodyEl.ownerDocument;

  // Snapshot original <a:p> elements (for undo) and keep them by index for
  // cloning pPr/rPr from source references.
  const origPEls = Array.from(txBodyEl.getElementsByTagNameNS(A_NS, "p"));
  const origParagraphs = prevParagraphs;

  // Collect source run XML elements indexed by [pIdx][rIdx].
  const origRunEls: Element[][] = [];
  for (const pEl of origPEls) {
    const runEls: Element[] = [];
    for (const child of Array.from(pEl.children)) {
      if (RUN_LOCAL_NAMES.has(child.localName)) {
        runEls.push(child);
      }
    }
    origRunEls.push(runEls);
  }

  // Remove existing <a:p> elements from txBody.
  for (const pEl of origPEls) {
    removeChild(pEl);
  }

  // Build new <a:p> elements and model paragraphs.
  const newParagraphs: TextParagraph[] = [];

  for (const opPara of op.paragraphs) {
    const pEl = doc.createElementNS(A_NS, "a:p");

    // Clone a:pPr from source paragraph.
    let level = 0;
    let pPrNode: SafeXmlNode | undefined;
    let endParaRPrNode: SafeXmlNode | undefined;
    if (opPara.sourceParagraphIndex >= 0 && opPara.sourceParagraphIndex < origParagraphs.length) {
      const srcPara = origParagraphs[opPara.sourceParagraphIndex];
      level = srcPara.level;
      if (srcPara.properties?.element) {
        pEl.appendChild(srcPara.properties.element.cloneNode(true));
        pPrNode = new SafeXmlNode(pEl.getElementsByTagNameNS(A_NS, "pPr")[0]);
      }
      if (srcPara.endParaRPr?.element) {
        endParaRPrNode = new SafeXmlNode(srcPara.endParaRPr.element.cloneNode(true) as Element);
      }
    }

    const runs: TextRun[] = [];

    // Resolve default rPr to clone when sourceRun is not specified.
    function resolveDefaultRPr(srcParaIdx: number): Element | null {
      if (srcParaIdx < 0 || srcParaIdx >= origParagraphs.length) return null;
      const srcRuns = origRunEls[srcParaIdx];
      if (!srcRuns || srcRuns.length === 0) return null;
      const firstRunEl = srcRuns[0];
      const rPr = firstRunEl.getElementsByTagNameNS(A_NS, "rPr")[0];
      return rPr ?? null;
    }

    // When the source paragraph has no runs at all, PowerPoint takes the
    // formatting for newly typed text from a:endParaRPr. Convert it into an
    // a:rPr (same attributes and children, different element name) so new
    // runs keep the paragraph's pending-text size/color/font.
    function endParaRPrAsRPr(srcParaIdx: number): Element | null {
      if (srcParaIdx < 0 || srcParaIdx >= origParagraphs.length) return null;
      const src = origParagraphs[srcParaIdx].endParaRPr?.element;
      if (!src) return null;
      const rPr = doc.createElementNS(A_NS, "a:rPr");
      for (const attr of Array.from(src.attributes)) {
        if (attr.name.startsWith("xmlns")) continue;
        rPr.setAttribute(attr.name, attr.value);
      }
      for (const child of Array.from(src.childNodes)) {
        rPr.appendChild(child.cloneNode(true));
      }
      return rPr;
    }

    for (const opRun of opPara.runs) {
      // Resolve the source rPr element to clone styling from.
      let rPrEl: Element | null = null;
      if (opRun.sourceRun) {
        const [srcPI, srcRI] = opRun.sourceRun;
        const srcRunEl = origRunEls[srcPI]?.[srcRI];
        if (srcRunEl) {
          rPrEl = srcRunEl.getElementsByTagNameNS(A_NS, "rPr")[0] ?? null;
        }
      } else {
        rPrEl =
          resolveDefaultRPr(opPara.sourceParagraphIndex) ??
          endParaRPrAsRPr(opPara.sourceParagraphIndex);
      }

      // a:br for newlines, a:r for text.
      if (opRun.text === "\n") {
        const brEl = doc.createElementNS(A_NS, "a:br");
        if (rPrEl) brEl.appendChild(rPrEl.cloneNode(true));
        pEl.appendChild(brEl);
        runs.push({
          text: "\n",
          properties: rPrEl
            ? new SafeXmlNode(brEl.getElementsByTagNameNS(A_NS, "rPr")[0])
            : undefined,
        });
      } else {
        const rEl = doc.createElementNS(A_NS, "a:r");
        if (rPrEl) rEl.appendChild(rPrEl.cloneNode(true));
        const tEl = doc.createElementNS(A_NS, "a:t");
        tEl.textContent = opRun.text;
        rEl.appendChild(tEl);
        pEl.appendChild(rEl);
        runs.push({
          text: opRun.text,
          properties: rPrEl
            ? new SafeXmlNode(rEl.getElementsByTagNameNS(A_NS, "rPr")[0])
            : undefined,
        });
      }
    }

    // Append a:endParaRPr if the source paragraph had one.
    if (endParaRPrNode?.element) {
      pEl.appendChild(endParaRPrNode.element);
    }

    insertChild(txBodyEl, pEl, null);

    newParagraphs.push({
      properties: pPrNode,
      runs,
      level,
      endParaRPr: endParaRPrNode,
    });
  }

  // Update the typed model.
  textBody.paragraphs = newParagraphs;
  sourcePackage.markDirty(slide.id);

  return {
    affectedSlideIds: [slide.id],
    undo: () => {
      // Remove new <a:p> elements.
      const currentPEls = Array.from(txBodyEl.getElementsByTagNameNS(A_NS, "p"));
      for (const pEl of currentPEls) {
        removeChild(pEl);
      }
      // Restore original <a:p> elements.
      for (const pEl of origPEls) {
        insertChild(txBodyEl, pEl, null);
      }
      // Restore model.
      textBody.paragraphs = prevParagraphs;
      sourcePackage.markDirty(slide.id);
    },
  };
}

// ---------------------------------------------------------------------------
// setNodeTransform
// ---------------------------------------------------------------------------

interface XfrmHandle {
  xfrmEl: Element;
  offEl: Element;
  extEl: Element;
  created: boolean;
}

/**
 * Find the node's xfrm element, creating one when the node inherits its
 * transform (e.g. layout placeholders have no xfrm of their own; moving
 * them requires writing a new one into the slide part).
 */
function getOrCreateXfrm(node: SlideNode): XfrmHandle {
  const sourceEl = requireElement(node.source, "node element");
  const doc = sourceEl.ownerDocument;

  let container: Element;
  let xfrmQualified: string;
  let xfrmNs: string;
  let insertRef: Element | null;

  if (node.nodeType === "table" || node.nodeType === "chart") {
    // Graphic frames carry p:xfrm directly, between nvGraphicFramePr and a:graphic.
    container = sourceEl;
    xfrmQualified = "p:xfrm";
    xfrmNs = P_NS;
    insertRef = node.source.child("graphic").element;
  } else {
    const prName = node.nodeType === "group" ? "grpSpPr" : "spPr";
    container = requireElement(node.source.child(prName), `${prName} of node "${node.id}"`);
    xfrmQualified = "a:xfrm";
    xfrmNs = A_NS;
    // xfrm is the first element in the shape-properties sequence.
    insertRef = container.firstElementChild;
  }

  let xfrmEl: Element | null = null;
  for (const child of container.children) {
    if (child.localName === "xfrm") {
      xfrmEl = child;
      break;
    }
  }

  const created = xfrmEl === null;
  if (!xfrmEl) {
    xfrmEl = doc.createElementNS(xfrmNs, xfrmQualified);
    insertChild(container, xfrmEl, insertRef);
  }

  let offEl: Element | null = null;
  let extEl: Element | null = null;
  for (const child of xfrmEl.children) {
    if (child.localName === "off") offEl = child;
    else if (child.localName === "ext") extEl = child;
  }
  if (!offEl) {
    offEl = doc.createElementNS(A_NS, "a:off");
    insertChild(xfrmEl, offEl, xfrmEl.firstElementChild);
  }
  if (!extEl) {
    extEl = doc.createElementNS(A_NS, "a:ext");
    insertChild(xfrmEl, extEl, offEl.nextElementSibling);
  }

  return { xfrmEl, offEl, extEl, created };
}

function applySetNodeTransform(pres: PresentationData, op: SetNodeTransformOperation): EditResult {
  const sourcePackage = requirePkg(pres);
  const slide = findSlide(pres, op.slideId);
  const node = findNode(pres, slide, op.nodeId);

  const prevModel = {
    position: { ...node.position },
    size: { ...node.size },
    rotation: node.rotation,
    flipH: node.flipH,
    flipV: node.flipV,
  };

  const { xfrmEl, offEl, extEl, created } = getOrCreateXfrm(node);

  const prevAttrs = {
    x: offEl.getAttribute("x") ?? undefined,
    y: offEl.getAttribute("y") ?? undefined,
    cx: extEl.getAttribute("cx") ?? undefined,
    cy: extEl.getAttribute("cy") ?? undefined,
    rot: xfrmEl.getAttribute("rot") ?? undefined,
    flipH: xfrmEl.getAttribute("flipH") ?? undefined,
    flipV: xfrmEl.getAttribute("flipV") ?? undefined,
  };

  // A freshly created xfrm must carry the full resolved transform, not just
  // the fields being changed: the resolved model values are the source.
  const position = op.position ?? (created ? node.position : undefined);
  const size = op.size ?? (created ? node.size : undefined);

  if (position) {
    offEl.setAttribute("x", String(pxToEmu(position.x)));
    offEl.setAttribute("y", String(pxToEmu(position.y)));
    node.position = { ...position };
  }
  if (size) {
    extEl.setAttribute("cx", String(pxToEmu(size.w)));
    extEl.setAttribute("cy", String(pxToEmu(size.h)));
    node.size = { ...size };
  }
  if (op.rotation !== undefined) {
    setOrRemoveAttr(xfrmEl, "rot", op.rotation === 0 ? undefined : String(degToAngle(op.rotation)));
    node.rotation = op.rotation;
  }
  if (op.flipH !== undefined) {
    setOrRemoveAttr(xfrmEl, "flipH", op.flipH ? "1" : undefined);
    node.flipH = op.flipH;
  }
  if (op.flipV !== undefined) {
    setOrRemoveAttr(xfrmEl, "flipV", op.flipV ? "1" : undefined);
    node.flipV = op.flipV;
  }

  // Groups created without an xfrm need a child space matching the resolved
  // extent so children keep rendering 1:1.
  if (created && node.nodeType === "group") {
    const group = node as GroupNodeData;
    const doc = xfrmEl.ownerDocument;
    const chOff = doc.createElementNS(A_NS, "a:chOff");
    chOff.setAttribute("x", offEl.getAttribute("x") ?? "0");
    chOff.setAttribute("y", offEl.getAttribute("y") ?? "0");
    const chExt = doc.createElementNS(A_NS, "a:chExt");
    chExt.setAttribute("cx", extEl.getAttribute("cx") ?? "0");
    chExt.setAttribute("cy", extEl.getAttribute("cy") ?? "0");
    insertChild(xfrmEl, chOff, extEl.nextElementSibling);
    insertChild(xfrmEl, chExt, chOff.nextElementSibling);
    group.childOffset = { ...node.position };
    group.childExtent = { ...node.size };
  }

  sourcePackage.markDirty(slide.id);

  return {
    affectedSlideIds: [slide.id],
    undo: () => {
      if (created) {
        removeChild(xfrmEl);
      } else {
        setOrRemoveAttr(offEl, "x", prevAttrs.x);
        setOrRemoveAttr(offEl, "y", prevAttrs.y);
        setOrRemoveAttr(extEl, "cx", prevAttrs.cx);
        setOrRemoveAttr(extEl, "cy", prevAttrs.cy);
        setOrRemoveAttr(xfrmEl, "rot", prevAttrs.rot);
        setOrRemoveAttr(xfrmEl, "flipH", prevAttrs.flipH);
        setOrRemoveAttr(xfrmEl, "flipV", prevAttrs.flipV);
      }
      node.position = prevModel.position;
      node.size = prevModel.size;
      node.rotation = prevModel.rotation;
      node.flipH = prevModel.flipH;
      node.flipV = prevModel.flipV;
      sourcePackage.markDirty(slide.id);
    },
  };
}

// ---------------------------------------------------------------------------
// setSolidFill
// ---------------------------------------------------------------------------

const FILL_LOCAL_NAMES = ["solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "noFill"];
/** spPr children that come after the fill in the schema sequence. */
const AFTER_FILL_LOCAL_NAMES = new Set([
  "ln",
  "effectLst",
  "effectDag",
  "scene3d",
  "sp3d",
  "extLst",
]);

function applySetSolidFill(pres: PresentationData, op: SetSolidFillOperation): EditResult {
  const sourcePackage = requirePkg(pres);
  const slide = findSlide(pres, op.slideId);
  const node = findNode(pres, slide, op.nodeId);

  if (node.nodeType !== "shape") {
    throw new Error(`applyEdit: setSolidFill only supports shape nodes, got "${node.nodeType}"`);
  }
  const shape = node as ShapeNodeData;

  const colorMatch = /^#?([0-9a-fA-F]{6})$/.exec(op.color);
  if (!colorMatch) {
    throw new Error(`applyEdit: invalid color "${op.color}"; expected 6-digit hex`);
  }
  const color = colorMatch[1].toUpperCase();

  const spPrEl = requireElement(node.source.child("spPr"), `spPr of node "${node.id}"`);
  const doc = spPrEl.ownerDocument;

  let oldFillEl: Element | null = null;
  for (const child of spPrEl.children) {
    if (FILL_LOCAL_NAMES.includes(child.localName)) {
      oldFillEl = child;
      break;
    }
  }

  let insertRef: Element | null = null;
  if (oldFillEl) {
    insertRef = oldFillEl.nextElementSibling;
    removeChild(oldFillEl);
  } else {
    for (const child of spPrEl.children) {
      if (AFTER_FILL_LOCAL_NAMES.has(child.localName)) {
        insertRef = child;
        break;
      }
    }
  }

  const solidFillEl = doc.createElementNS(A_NS, "a:solidFill");
  const srgbEl = doc.createElementNS(A_NS, "a:srgbClr");
  srgbEl.setAttribute("val", color);
  solidFillEl.appendChild(srgbEl);
  insertChild(spPrEl, solidFillEl, insertRef);

  const prevFill = shape.fill;
  shape.fill = new SafeXmlNode(solidFillEl);
  sourcePackage.markDirty(slide.id);

  return {
    affectedSlideIds: [slide.id],
    undo: () => {
      const ref = solidFillEl.nextElementSibling;
      removeChild(solidFillEl);
      if (oldFillEl) {
        insertChild(spPrEl, oldFillEl, ref);
      }
      shape.fill = prevFill;
      sourcePackage.markDirty(slide.id);
    },
  };
}

// ---------------------------------------------------------------------------
// deleteNode
// ---------------------------------------------------------------------------

function applyDeleteNode(pres: PresentationData, op: DeleteNodeOperation): EditResult {
  const sourcePackage = requirePkg(pres);
  const slide = findSlide(pres, op.slideId);
  const node = findNode(pres, slide, op.nodeId);

  const el = requireElement(node.source, "node element");
  const parentEl = el.parentElement;
  if (!parentEl) throw new Error(`applyEdit: node "${op.nodeId}" is not attached to the slide`);

  const domRef = el.nextElementSibling;
  removeChild(el);

  const modelIndex = slide.nodes.indexOf(node);
  slide.nodes.splice(modelIndex, 1);
  sourcePackage.markDirty(slide.id);

  return {
    affectedSlideIds: [slide.id],
    undo: () => {
      insertChild(parentEl, el, domRef);
      slide.nodes.splice(modelIndex, 0, node);
      sourcePackage.markDirty(slide.id);
    },
  };
}

// ---------------------------------------------------------------------------
// Slide-level operations
// ---------------------------------------------------------------------------

interface SlideListContext {
  sourcePackage: PptxPackage;
  lstNode: SafeXmlNode;
  lstEl: Element;
  relsText: string;
  rels: Map<string, RelEntry>;
}

async function loadSlideListContext(pres: PresentationData): Promise<SlideListContext> {
  const sourcePackage = requirePkg(pres);
  const presRoot = sourcePackage.getXmlRoot(PRESENTATION_PATH);
  if (!presRoot?.exists()) {
    throw new Error("applyEdit: presentation.xml is not registered on the package");
  }
  const lstNode = presRoot.child("sldIdLst");
  const lstEl = lstNode.element;
  if (!lstEl) {
    throw new Error("applyEdit: unsupported package: presentation.xml has no sldIdLst");
  }
  const relsText = (await sourcePackage.readText(PRESENTATION_RELS_PATH)) ?? "";
  return { sourcePackage, lstNode, lstEl, relsText, rels: parseRels(relsText) };
}

/** Locate the p:sldId element referencing the given slide part. */
function findSldIdEntry(ctx: SlideListContext, slideId: string): { el: Element; rId: string } {
  for (const sldId of ctx.lstNode.children("sldId")) {
    const rId = sldId.attr("r:id");
    if (!rId) continue;
    const rel = ctx.rels.get(rId);
    if (rel && resolveRelTarget("ppt", rel.target) === slideId) {
      return { el: requireElement(sldId, "sldId element"), rId };
    }
  }
  throw new Error(`applyEdit: slide "${slideId}" not found in presentation.xml sldIdLst`);
}

function sldIdElements(ctx: SlideListContext): Element[] {
  return ctx.lstNode.children("sldId").map((n) => requireElement(n, "sldId element"));
}

function relsPathFor(partPath: string): string {
  const idx = partPath.lastIndexOf("/");
  const dir = idx >= 0 ? partPath.slice(0, idx) : "";
  const name = partPath.slice(idx + 1);
  return `${dir}/_rels/${name}.rels`;
}

// --- moveSlide -------------------------------------------------------------

async function applyMoveSlide(pres: PresentationData, op: MoveSlideOperation): Promise<EditResult> {
  const ctx = await loadSlideListContext(pres);
  const slide = findSlide(pres, op.slideId);
  const fromIndex = pres.slides.indexOf(slide);
  const toIndex = Math.max(0, Math.min(op.toIndex, pres.slides.length - 1));

  if (fromIndex === toIndex) {
    return { affectedSlideIds: [], undo: () => {} };
  }

  const { el } = findSldIdEntry(ctx, op.slideId);

  const moveTo = (targetIndex: number) => {
    el.remove();
    const remaining = sldIdElements(ctx).filter((e) => e !== el);
    insertChild(ctx.lstEl, el, remaining[targetIndex] ?? null);
    ctx.sourcePackage.markDirty(PRESENTATION_PATH);
  };

  moveTo(toIndex);
  pres.slides.splice(fromIndex, 1);
  pres.slides.splice(toIndex, 0, slide);
  reindexSlides(pres);

  return {
    affectedSlideIds: [slide.id],
    undo: () => {
      moveTo(fromIndex);
      pres.slides.splice(toIndex, 1);
      pres.slides.splice(fromIndex, 0, slide);
      reindexSlides(pres);
    },
  };
}

// --- duplicateSlide ----------------------------------------------------------

function nextSlidePartPath(sourcePackage: PptxPackage): string {
  let max = 0;
  for (const path of sourcePackage.paths()) {
    const match = /^ppt\/slides\/slide(\d+)\.xml$/.exec(path);
    if (match) max = Math.max(max, Number(match[1]));
  }
  return `ppt/slides/slide${max + 1}.xml`;
}

function nextRelId(rels: Map<string, RelEntry>): string {
  let max = 0;
  for (const rId of rels.keys()) {
    const num = Number(rId.replace(/\D/g, ""));
    if (Number.isFinite(num)) max = Math.max(max, num);
  }
  return `rId${max + 1}`;
}

function nextSldIdNumber(ctx: SlideListContext): number {
  // Slide ids must be >= 256 per the spec.
  let max = 255;
  for (const sldId of ctx.lstNode.children("sldId")) {
    const num = Number(sldId.attr("id"));
    if (Number.isFinite(num)) max = Math.max(max, num);
  }
  return max + 1;
}

async function applyDuplicateSlide(
  pres: PresentationData,
  op: DuplicateSlideOperation,
): Promise<EditResult> {
  const ctx = await loadSlideListContext(pres);
  const { sourcePackage } = ctx;
  const source = findSlide(pres, op.slideId);
  const sourceIndex = pres.slides.indexOf(source);
  const { el: sourceSldIdEl } = findSldIdEntry(ctx, op.slideId);

  const newPath = nextSlidePartPath(sourcePackage);
  const newRelsPath = relsPathFor(newPath);

  const slideText = await sourcePackage.readText(source.id);
  if (slideText === undefined) {
    throw new Error(`applyEdit: slide part "${source.id}" is missing from the package`);
  }
  const slideRelsText = await sourcePackage.readText(relsPathFor(source.id));
  const ctText = (await sourcePackage.readText(CONTENT_TYPES_PATH)) ?? "";

  // --- Package parts ---
  sourcePackage.setEntry(newPath, slideText);
  if (slideRelsText !== undefined) {
    sourcePackage.setEntry(newRelsPath, slideRelsText);
  }

  // --- [Content_Types].xml: add an Override for the new part ---
  const ctRoot = parseXml(ctText);
  const ctRootEl = requireElement(ctRoot, "[Content_Types].xml root");
  const sourceOverride = ctRoot
    .children("Override")
    .find((o) => o.attr("PartName") === `/${source.id}`);
  const overrideEl = ctRootEl.ownerDocument.createElementNS(CT_NS, "Override");
  overrideEl.setAttribute("PartName", `/${newPath}`);
  overrideEl.setAttribute("ContentType", sourceOverride?.attr("ContentType") ?? SLIDE_CONTENT_TYPE);
  insertChild(ctRootEl, overrideEl, null);
  sourcePackage.setEntry(CONTENT_TYPES_PATH, serializePartText(ctRootEl, ctText));

  // --- presentation.xml.rels: relationship for the new part ---
  const newRId = nextRelId(ctx.rels);
  const relsRoot = parseXml(ctx.relsText);
  const relsRootEl = requireElement(relsRoot, "presentation rels root");
  const relEl = relsRootEl.ownerDocument.createElementNS(RELS_NS, "Relationship");
  relEl.setAttribute("Id", newRId);
  relEl.setAttribute("Type", SLIDE_REL_TYPE);
  relEl.setAttribute("Target", newPath.replace(/^ppt\//, ""));
  insertChild(relsRootEl, relEl, null);
  sourcePackage.setEntry(PRESENTATION_RELS_PATH, serializePartText(relsRootEl, ctx.relsText));

  // --- presentation.xml: sldId entry directly after the source slide ---
  const sldIdEl = ctx.lstEl.ownerDocument.createElementNS(P_NS, "p:sldId");
  sldIdEl.setAttribute("id", String(nextSldIdNumber(ctx)));
  sldIdEl.setAttributeNS(R_NS, "r:id", newRId);
  insertChild(ctx.lstEl, sldIdEl, sourceSldIdEl.nextElementSibling);
  sourcePackage.markDirty(PRESENTATION_PATH);

  // --- Model: parse the copy through the regular slide pipeline ---
  const newRoot = parseXml(slideText);
  sourcePackage.registerXmlRoot(newPath, newRoot);
  const newRels = parseRels(slideRelsText ?? "");
  const newSlide = parseSlide(newRoot, sourceIndex + 1, newRels, newPath, pres.diagramDrawings);
  if (newSlide.layoutIndex) {
    newSlide.layoutIndex = resolveRelTarget("ppt/slides", newSlide.layoutIndex);
  }
  pres.slides.splice(sourceIndex + 1, 0, newSlide);
  reindexSlides(pres);
  materializeSlide(pres, newSlide);

  return {
    affectedSlideIds: [newPath],
    createdSlideId: newPath,
    undo: () => {
      sourcePackage.deleteEntry(newPath);
      if (slideRelsText !== undefined) sourcePackage.deleteEntry(newRelsPath);
      sourcePackage.setEntry(CONTENT_TYPES_PATH, ctText);
      sourcePackage.setEntry(PRESENTATION_RELS_PATH, ctx.relsText);
      removeChild(sldIdEl);
      sourcePackage.markDirty(PRESENTATION_PATH);
      pres.slides.splice(pres.slides.indexOf(newSlide), 1);
      reindexSlides(pres);
    },
  };
}

// --- deleteSlide -------------------------------------------------------------

async function applyDeleteSlide(
  pres: PresentationData,
  op: DeleteSlideOperation,
): Promise<EditResult> {
  if (pres.slides.length <= 1) {
    throw new Error("applyEdit: cannot delete the last slide in the deck");
  }

  const ctx = await loadSlideListContext(pres);
  const { sourcePackage } = ctx;
  const slide = findSlide(pres, op.slideId);
  const slideIndex = pres.slides.indexOf(slide);
  const { el: sldIdEl, rId } = findSldIdEntry(ctx, op.slideId);
  const slideRelsPath = relsPathFor(slide.id);

  // Snapshots for undo.
  const slideBytes = await sourcePackage.readBytes(slide.id);
  const slideRelsBytes = await sourcePackage.readBytes(slideRelsPath);
  const slideXmlRoot = sourcePackage.getXmlRoot(slide.id);
  const ctText = (await sourcePackage.readText(CONTENT_TYPES_PATH)) ?? "";

  // --- presentation.xml ---
  const sldIdRef = sldIdEl.nextElementSibling;
  removeChild(sldIdEl);
  sourcePackage.markDirty(PRESENTATION_PATH);

  // --- presentation.xml.rels ---
  const relsRoot = parseXml(ctx.relsText);
  const relsRootEl = requireElement(relsRoot, "presentation rels root");
  const relNode = relsRoot.children("Relationship").find((r) => r.attr("Id") === rId);
  if (relNode?.element) {
    removeChild(relNode.element);
  }
  sourcePackage.setEntry(PRESENTATION_RELS_PATH, serializePartText(relsRootEl, ctx.relsText));

  // --- [Content_Types].xml ---
  const ctRoot = parseXml(ctText);
  const ctRootEl = requireElement(ctRoot, "[Content_Types].xml root");
  const overrideNode = ctRoot
    .children("Override")
    .find((o) => o.attr("PartName") === `/${slide.id}`);
  if (overrideNode?.element) {
    removeChild(overrideNode.element);
    sourcePackage.setEntry(CONTENT_TYPES_PATH, serializePartText(ctRootEl, ctText));
  }

  // --- Parts (media stays; shared and harmless when orphaned) ---
  sourcePackage.deleteEntry(slide.id);
  sourcePackage.deleteEntry(slideRelsPath);

  // --- Model ---
  pres.slides.splice(slideIndex, 1);
  reindexSlides(pres);

  return {
    affectedSlideIds: [slide.id],
    undo: () => {
      if (slideBytes) sourcePackage.setEntry(slide.id, slideBytes);
      if (slideRelsBytes) sourcePackage.setEntry(slideRelsPath, slideRelsBytes);
      // Re-attach the live XML document so the restored slide stays editable.
      if (slideXmlRoot) sourcePackage.registerXmlRoot(slide.id, slideXmlRoot);
      sourcePackage.setEntry(PRESENTATION_RELS_PATH, ctx.relsText);
      sourcePackage.setEntry(CONTENT_TYPES_PATH, ctText);
      insertChild(ctx.lstEl, sldIdEl, sldIdRef);
      sourcePackage.markDirty(PRESENTATION_PATH);
      pres.slides.splice(slideIndex, 0, slide);
      reindexSlides(pres);
    },
  };
}
