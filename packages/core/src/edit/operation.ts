/**
 * Edit operations for presentations opened with `keepSourcePackage: true`.
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
import type { TableNodeData } from "../model/nodes/table";
import { materializeSlide, PresentationData } from "../model/presentation";
import { parseSlide, SlideData, SlideNode } from "../model/slide";
import type { PptxPackage } from "../ooxml/package";
import { parseRels, RelEntry, resolveRelTarget } from "../ooxml/rel";
import { degToAngle, emuToPx, pxToEmu } from "../ooxml/unit";
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
        "Parse the zip with { keepSourcePackage: true } to enable editing.",
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
  const element = node.element;
  if (!element) throw new Error(`applyEdit: ${what} is missing`);
  return element;
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

  const runElement = requireElement(runNode, "run element");
  let tElement = runNode.child("t").element;
  if (!tElement) {
    tElement = runElement.ownerDocument.createElementNS(A_NS, "a:t");
    insertChild(runElement, tElement, null);
  }

  const prevText = run.text;
  tElement.textContent = op.text;
  run.text = op.text;
  sourcePackage.markDirty(slide.id);

  return {
    affectedSlideIds: [slide.id],
    undo: () => {
      tElement.textContent = prevText;
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
  const txBodyElement = requireElement(txBodyNode, "txBody");
  const doc = txBodyElement.ownerDocument;

  // Snapshot original <a:p> elements (for undo) and keep them by index for
  // cloning pPr/rPr from source references.
  const origParagraphElements = Array.from(txBodyElement.getElementsByTagNameNS(A_NS, "p"));
  const origParagraphs = prevParagraphs;

  // Collect source run XML elements indexed by [pIdx][rIdx].
  const origRunElements: Element[][] = [];
  for (const paragraphElement of origParagraphElements) {
    const runElements: Element[] = [];
    for (const child of Array.from(paragraphElement.children)) {
      if (RUN_LOCAL_NAMES.has(child.localName)) {
        runElements.push(child);
      }
    }
    origRunElements.push(runElements);
  }

  // Remove existing <a:p> elements from txBody.
  for (const paragraphElement of origParagraphElements) {
    removeChild(paragraphElement);
  }

  // Build new <a:p> elements and model paragraphs.
  const newParagraphs: TextParagraph[] = [];

  for (const opPara of op.paragraphs) {
    const paragraphElement = doc.createElementNS(A_NS, "a:p");

    // Clone a:pPr from source paragraph.
    let level = 0;
    let pPrNode: SafeXmlNode | undefined;
    let endParaRPrNode: SafeXmlNode | undefined;
    if (opPara.sourceParagraphIndex >= 0 && opPara.sourceParagraphIndex < origParagraphs.length) {
      const srcPara = origParagraphs[opPara.sourceParagraphIndex];
      level = srcPara.level;
      if (srcPara.properties?.element) {
        paragraphElement.appendChild(srcPara.properties.element.cloneNode(true));
        pPrNode = new SafeXmlNode(paragraphElement.getElementsByTagNameNS(A_NS, "pPr")[0]);
      }
      if (srcPara.endParaRPr?.element) {
        endParaRPrNode = new SafeXmlNode(srcPara.endParaRPr.element.cloneNode(true) as Element);
      }
    }

    const runs: TextRun[] = [];

    // Resolve default rPr to clone when sourceRun is not specified.
    function resolveDefaultRPr(srcParaIdx: number): Element | null {
      if (srcParaIdx < 0 || srcParaIdx >= origParagraphs.length) return null;
      const srcRuns = origRunElements[srcParaIdx];
      if (!srcRuns || srcRuns.length === 0) return null;
      const firstRunElement = srcRuns[0];
      const rPr = firstRunElement.getElementsByTagNameNS(A_NS, "rPr")[0];
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
      let rPrElement: Element | null = null;
      if (opRun.sourceRun) {
        const [srcPI, srcRI] = opRun.sourceRun;
        const srcRunElement = origRunElements[srcPI]?.[srcRI];
        if (srcRunElement) {
          rPrElement = srcRunElement.getElementsByTagNameNS(A_NS, "rPr")[0] ?? null;
        }
      } else {
        rPrElement =
          resolveDefaultRPr(opPara.sourceParagraphIndex) ??
          endParaRPrAsRPr(opPara.sourceParagraphIndex);
      }

      // a:br for newlines, a:r for text.
      if (opRun.text === "\n") {
        const brElement = doc.createElementNS(A_NS, "a:br");
        if (rPrElement) brElement.appendChild(rPrElement.cloneNode(true));
        paragraphElement.appendChild(brElement);
        runs.push({
          text: "\n",
          properties: rPrElement
            ? new SafeXmlNode(brElement.getElementsByTagNameNS(A_NS, "rPr")[0])
            : undefined,
        });
      } else {
        const rElement = doc.createElementNS(A_NS, "a:r");
        if (rPrElement) rElement.appendChild(rPrElement.cloneNode(true));
        const tElement = doc.createElementNS(A_NS, "a:t");
        tElement.textContent = opRun.text;
        rElement.appendChild(tElement);
        paragraphElement.appendChild(rElement);
        runs.push({
          text: opRun.text,
          properties: rPrElement
            ? new SafeXmlNode(rElement.getElementsByTagNameNS(A_NS, "rPr")[0])
            : undefined,
        });
      }
    }

    // Append a:endParaRPr if the source paragraph had one.
    if (endParaRPrNode?.element) {
      paragraphElement.appendChild(endParaRPrNode.element);
    }

    insertChild(txBodyElement, paragraphElement, null);

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
      const currentParagraphElements = Array.from(txBodyElement.getElementsByTagNameNS(A_NS, "p"));
      for (const paragraphElement of currentParagraphElements) {
        removeChild(paragraphElement);
      }
      // Restore original <a:p> elements.
      for (const paragraphElement of origParagraphElements) {
        insertChild(txBodyElement, paragraphElement, null);
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
  xfrmElement: Element;
  offElement: Element;
  extElement: Element;
  created: boolean;
}

/**
 * Find the node's xfrm element, creating one when the node inherits its
 * transform (e.g. layout placeholders have no xfrm of their own; moving
 * them requires writing a new one into the slide part).
 */
function getOrCreateXfrm(node: SlideNode): XfrmHandle {
  const sourceElement = requireElement(node.source, "node element");
  const doc = sourceElement.ownerDocument;

  let container: Element;
  let xfrmQualified: string;
  let xfrmNs: string;
  let insertRef: Element | null;

  if (node.nodeType === "table" || node.nodeType === "chart") {
    // Graphic frames carry p:xfrm directly, between nvGraphicFramePr and a:graphic.
    container = sourceElement;
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

  let xfrmElement: Element | null = null;
  for (const child of container.children) {
    if (child.localName === "xfrm") {
      xfrmElement = child;
      break;
    }
  }

  const created = xfrmElement === null;
  if (!xfrmElement) {
    xfrmElement = doc.createElementNS(xfrmNs, xfrmQualified);
    insertChild(container, xfrmElement, insertRef);
  }

  let offElement: Element | null = null;
  let extElement: Element | null = null;
  for (const child of xfrmElement.children) {
    if (child.localName === "off") offElement = child;
    else if (child.localName === "ext") extElement = child;
  }
  if (!offElement) {
    offElement = doc.createElementNS(A_NS, "a:off");
    insertChild(xfrmElement, offElement, xfrmElement.firstElementChild);
  }
  if (!extElement) {
    extElement = doc.createElementNS(A_NS, "a:ext");
    insertChild(xfrmElement, extElement, offElement.nextElementSibling);
  }

  return { xfrmElement, offElement, extElement, created };
}

/**
 * Scale a table's grid to a new frame size, the way PowerPoint does when you
 * drag a table's resize handle.
 *
 * A table is laid out from its grid (the sum of `a:gridCol/@w` and of
 * `a:tr/@h`), not from the graphicFrame extent, so writing the extent alone
 * shrinks the selection box while the table keeps its old size. Returns an undo
 * closure, or null when there is no scalable grid.
 */
function scaleTableGrid(table: TableNodeData, size: NodeSize): (() => void) | null {
  const tbl = table.source.child("graphic").child("graphicData").child("tbl");
  const columnElements = tbl.child("tblGrid").children("gridCol");
  const rowElements = tbl.children("tr");

  const gridW = columnElements.reduce((sum, col) => sum + (col.numAttr("w") ?? 0), 0);
  const gridH = rowElements.reduce((sum, row) => sum + (row.numAttr("h") ?? 0), 0);
  const scaleX = gridW > 0 ? pxToEmu(size.w) / gridW : 0;
  const scaleY = gridH > 0 ? pxToEmu(size.h) / gridH : 0;
  if (scaleX === 0 && scaleY === 0) return null;

  const prevColumns = [...table.columns];
  const prevHeights = table.rows.map((row) => row.height);
  const restore: (() => void)[] = [];

  const scaleAttr = (node: SafeXmlNode, name: string, scale: number): number | null => {
    const element = node.element;
    const current = node.numAttr(name);
    if (!element || current === undefined || scale === 0) return null;
    const next = Math.round(current * scale);
    const prev = element.getAttribute(name) ?? undefined;
    restore.push(() => setOrRemoveAttr(element, name, prev));
    element.setAttribute(name, String(next));
    return next;
  };

  columnElements.forEach((col, index) => {
    const next = scaleAttr(col, "w", scaleX);
    if (next !== null && index < table.columns.length) table.columns[index] = emuToPx(next);
  });
  rowElements.forEach((row, index) => {
    const next = scaleAttr(row, "h", scaleY);
    if (next !== null && index < table.rows.length) table.rows[index].height = emuToPx(next);
  });

  return () => {
    for (const undo of restore.reverse()) undo();
    table.columns = prevColumns;
    table.rows.forEach((row, index) => {
      row.height = prevHeights[index];
    });
  };
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

  const { xfrmElement, offElement, extElement, created } = getOrCreateXfrm(node);

  const prevAttrs = {
    x: offElement.getAttribute("x") ?? undefined,
    y: offElement.getAttribute("y") ?? undefined,
    cx: extElement.getAttribute("cx") ?? undefined,
    cy: extElement.getAttribute("cy") ?? undefined,
    rot: xfrmElement.getAttribute("rot") ?? undefined,
    flipH: xfrmElement.getAttribute("flipH") ?? undefined,
    flipV: xfrmElement.getAttribute("flipV") ?? undefined,
  };

  // A freshly created xfrm must carry the full resolved transform, not just
  // the fields being changed: the resolved model values are the source.
  const position = op.position ?? (created ? node.position : undefined);
  const size = op.size ?? (created ? node.size : undefined);

  if (position) {
    offElement.setAttribute("x", String(pxToEmu(position.x)));
    offElement.setAttribute("y", String(pxToEmu(position.y)));
    node.position = { ...position };
  }
  let undoTableGrid: (() => void) | null = null;
  if (size) {
    extElement.setAttribute("cx", String(pxToEmu(size.w)));
    extElement.setAttribute("cy", String(pxToEmu(size.h)));
    if (node.nodeType === "table") undoTableGrid = scaleTableGrid(node as TableNodeData, size);
    node.size = { ...size };
  }
  if (op.rotation !== undefined) {
    setOrRemoveAttr(
      xfrmElement,
      "rot",
      op.rotation === 0 ? undefined : String(degToAngle(op.rotation)),
    );
    node.rotation = op.rotation;
  }
  if (op.flipH !== undefined) {
    setOrRemoveAttr(xfrmElement, "flipH", op.flipH ? "1" : undefined);
    node.flipH = op.flipH;
  }
  if (op.flipV !== undefined) {
    setOrRemoveAttr(xfrmElement, "flipV", op.flipV ? "1" : undefined);
    node.flipV = op.flipV;
  }

  // Groups created without an xfrm need a child space matching the resolved
  // extent so children keep rendering 1:1.
  if (created && node.nodeType === "group") {
    const group = node as GroupNodeData;
    const doc = xfrmElement.ownerDocument;
    const chOff = doc.createElementNS(A_NS, "a:chOff");
    chOff.setAttribute("x", offElement.getAttribute("x") ?? "0");
    chOff.setAttribute("y", offElement.getAttribute("y") ?? "0");
    const chExt = doc.createElementNS(A_NS, "a:chExt");
    chExt.setAttribute("cx", extElement.getAttribute("cx") ?? "0");
    chExt.setAttribute("cy", extElement.getAttribute("cy") ?? "0");
    insertChild(xfrmElement, chOff, extElement.nextElementSibling);
    insertChild(xfrmElement, chExt, chOff.nextElementSibling);
    group.childOffset = { ...node.position };
    group.childExtent = { ...node.size };
  }

  sourcePackage.markDirty(slide.id);

  return {
    affectedSlideIds: [slide.id],
    undo: () => {
      undoTableGrid?.();
      if (created) {
        removeChild(xfrmElement);
      } else {
        setOrRemoveAttr(offElement, "x", prevAttrs.x);
        setOrRemoveAttr(offElement, "y", prevAttrs.y);
        setOrRemoveAttr(extElement, "cx", prevAttrs.cx);
        setOrRemoveAttr(extElement, "cy", prevAttrs.cy);
        setOrRemoveAttr(xfrmElement, "rot", prevAttrs.rot);
        setOrRemoveAttr(xfrmElement, "flipH", prevAttrs.flipH);
        setOrRemoveAttr(xfrmElement, "flipV", prevAttrs.flipV);
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

  const spPrElement = requireElement(node.source.child("spPr"), `spPr of node "${node.id}"`);
  const doc = spPrElement.ownerDocument;

  let oldFillElement: Element | null = null;
  for (const child of spPrElement.children) {
    if (FILL_LOCAL_NAMES.includes(child.localName)) {
      oldFillElement = child;
      break;
    }
  }

  let insertRef: Element | null = null;
  if (oldFillElement) {
    insertRef = oldFillElement.nextElementSibling;
    removeChild(oldFillElement);
  } else {
    for (const child of spPrElement.children) {
      if (AFTER_FILL_LOCAL_NAMES.has(child.localName)) {
        insertRef = child;
        break;
      }
    }
  }

  const solidFillElement = doc.createElementNS(A_NS, "a:solidFill");
  const srgbElement = doc.createElementNS(A_NS, "a:srgbClr");
  srgbElement.setAttribute("val", color);
  solidFillElement.appendChild(srgbElement);
  insertChild(spPrElement, solidFillElement, insertRef);

  const prevFill = shape.fill;
  shape.fill = new SafeXmlNode(solidFillElement);
  sourcePackage.markDirty(slide.id);

  return {
    affectedSlideIds: [slide.id],
    undo: () => {
      const ref = solidFillElement.nextElementSibling;
      removeChild(solidFillElement);
      if (oldFillElement) {
        insertChild(spPrElement, oldFillElement, ref);
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

  const element = requireElement(node.source, "node element");
  const parentElement = element.parentElement;
  if (!parentElement)
    throw new Error(`applyEdit: node "${op.nodeId}" is not attached to the slide`);

  const domRef = element.nextElementSibling;
  removeChild(element);

  const modelIndex = slide.nodes.indexOf(node);
  slide.nodes.splice(modelIndex, 1);
  sourcePackage.markDirty(slide.id);

  return {
    affectedSlideIds: [slide.id],
    undo: () => {
      insertChild(parentElement, element, domRef);
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
  lstElement: Element;
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
  const lstElement = lstNode.element;
  if (!lstElement) {
    throw new Error("applyEdit: unsupported package: presentation.xml has no sldIdLst");
  }
  const relsText = (await sourcePackage.readText(PRESENTATION_RELS_PATH)) ?? "";
  return { sourcePackage, lstNode, lstElement, relsText, rels: parseRels(relsText) };
}

/** Locate the p:sldId element referencing the given slide part. */
function findSldIdEntry(ctx: SlideListContext, slideId: string): { element: Element; rId: string } {
  for (const sldId of ctx.lstNode.children("sldId")) {
    const rId = sldId.attr("r:id");
    if (!rId) continue;
    const rel = ctx.rels.get(rId);
    if (rel && resolveRelTarget("ppt", rel.target) === slideId) {
      return { element: requireElement(sldId, "sldId element"), rId };
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

  const { element } = findSldIdEntry(ctx, op.slideId);

  const moveTo = (targetIndex: number) => {
    element.remove();
    const remaining = sldIdElements(ctx).filter((e) => e !== element);
    insertChild(ctx.lstElement, element, remaining[targetIndex] ?? null);
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
  const { element: sourceSldIdElement } = findSldIdEntry(ctx, op.slideId);

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
  const ctRootElement = requireElement(ctRoot, "[Content_Types].xml root");
  const sourceOverride = ctRoot
    .children("Override")
    .find((o) => o.attr("PartName") === `/${source.id}`);
  const overrideElement = ctRootElement.ownerDocument.createElementNS(CT_NS, "Override");
  overrideElement.setAttribute("PartName", `/${newPath}`);
  overrideElement.setAttribute(
    "ContentType",
    sourceOverride?.attr("ContentType") ?? SLIDE_CONTENT_TYPE,
  );
  insertChild(ctRootElement, overrideElement, null);
  sourcePackage.setEntry(CONTENT_TYPES_PATH, serializePartText(ctRootElement, ctText));

  // --- presentation.xml.rels: relationship for the new part ---
  const newRId = nextRelId(ctx.rels);
  const relsRoot = parseXml(ctx.relsText);
  const relsRootElement = requireElement(relsRoot, "presentation rels root");
  const relElement = relsRootElement.ownerDocument.createElementNS(RELS_NS, "Relationship");
  relElement.setAttribute("Id", newRId);
  relElement.setAttribute("Type", SLIDE_REL_TYPE);
  relElement.setAttribute("Target", newPath.replace(/^ppt\//, ""));
  insertChild(relsRootElement, relElement, null);
  sourcePackage.setEntry(PRESENTATION_RELS_PATH, serializePartText(relsRootElement, ctx.relsText));

  // --- presentation.xml: sldId entry directly after the source slide ---
  const sldIdElement = ctx.lstElement.ownerDocument.createElementNS(P_NS, "p:sldId");
  sldIdElement.setAttribute("id", String(nextSldIdNumber(ctx)));
  sldIdElement.setAttributeNS(R_NS, "r:id", newRId);
  insertChild(ctx.lstElement, sldIdElement, sourceSldIdElement.nextElementSibling);
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
      removeChild(sldIdElement);
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
  const { element: sldIdElement, rId } = findSldIdEntry(ctx, op.slideId);
  const slideRelsPath = relsPathFor(slide.id);

  // Snapshots for undo.
  const slideBytes = await sourcePackage.readBytes(slide.id);
  const slideRelsBytes = await sourcePackage.readBytes(slideRelsPath);
  const slideXmlRoot = sourcePackage.getXmlRoot(slide.id);
  const ctText = (await sourcePackage.readText(CONTENT_TYPES_PATH)) ?? "";

  // --- presentation.xml ---
  const sldIdRef = sldIdElement.nextElementSibling;
  removeChild(sldIdElement);
  sourcePackage.markDirty(PRESENTATION_PATH);

  // --- presentation.xml.rels ---
  const relsRoot = parseXml(ctx.relsText);
  const relsRootElement = requireElement(relsRoot, "presentation rels root");
  const relNode = relsRoot.children("Relationship").find((r) => r.attr("Id") === rId);
  if (relNode?.element) {
    removeChild(relNode.element);
  }
  sourcePackage.setEntry(PRESENTATION_RELS_PATH, serializePartText(relsRootElement, ctx.relsText));

  // --- [Content_Types].xml ---
  const ctRoot = parseXml(ctText);
  const ctRootElement = requireElement(ctRoot, "[Content_Types].xml root");
  const overrideNode = ctRoot
    .children("Override")
    .find((o) => o.attr("PartName") === `/${slide.id}`);
  if (overrideNode?.element) {
    removeChild(overrideNode.element);
    sourcePackage.setEntry(CONTENT_TYPES_PATH, serializePartText(ctRootElement, ctText));
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
      insertChild(ctx.lstElement, sldIdElement, sldIdRef);
      sourcePackage.markDirty(PRESENTATION_PATH);
      pres.slides.splice(slideIndex, 0, slide);
      reindexSlides(pres);
    },
  };
}
