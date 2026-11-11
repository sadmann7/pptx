import type { NodePosition, NodeSize } from "../model/nodes/base";
import type { GroupNodeData } from "../model/nodes/group";
import { parseRenderableChild } from "../model/nodes/parser";
import type { PictureNodeData } from "../model/nodes/picture";
import type { ShapeNodeData, TextBody, TextParagraph, TextRun } from "../model/nodes/shape";
import { parseParagraph } from "../model/nodes/shape";
import type { TableNodeData } from "../model/nodes/table";
import {
  materializeSlide,
  PresentationData,
  resolveSlideNodePlaceholders,
} from "../model/presentation";
import { parseSlide, SlideData, SlideNode } from "../model/slide";
import { parseOoxmlBool } from "../ooxml/boolean";
import type { PptxPackage } from "../ooxml/package";
import { isExternalTargetMode, parseRels, RelEntry, resolveRelTarget } from "../ooxml/rel";
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

/** A cell of a table, by row and grid column (both zero-based). */
export interface TableCellAddress {
  row: number;
  column: number;
}

/** Replace the text of one run inside a shape's text body. */
export interface SetTextRunOperation {
  type: "setTextRun";
  slideId: string;
  nodeId: string;
  /** The table cell holding the text, when the node is a table. */
  cell?: TableCellAddress;
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

/** Replace the entire text body of a shape or table cell. */
export interface SetTextBodyOperation {
  type: "setTextBody";
  slideId: string;
  nodeId: string;
  /** The table cell holding the text, when the node is a table. */
  cell?: TableCellAddress;
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

/** A caret position inside a text body: a paragraph and a character offset in it. */
export interface TextPosition {
  paragraphIndex: number;
  /** Character offset into the paragraph's text (the joined text of its runs). */
  offset: number;
}

/** A span of text inside one shape's text body. `end` is exclusive. */
export interface TextRange {
  start: TextPosition;
  end: TextPosition;
}

/**
 * Character formatting. Omitted fields are left as they are; `null` removes
 * the run's own value so it inherits from the paragraph, layout, or master.
 */
export interface TextStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  /** Hex color, with or without leading `#` (e.g. "FF0000"). */
  color?: string | null;
  /** Latin typeface (e.g. "Arial"). */
  fontFamily?: string | null;
  /** Font size in points. */
  fontSize?: number | null;
}

/**
 * Apply character formatting to a shape's text. Without `range` the whole
 * text body is formatted, including the end-of-paragraph marks so new text
 * typed into an empty paragraph picks the formatting up.
 */
export interface SetTextStyleOperation {
  type: "setTextStyle";
  slideId: string;
  nodeId: string;
  /** The table cell holding the text, when the node is a table. */
  cell?: TableCellAddress;
  style: TextStyle;
  range?: TextRange;
}

export type ParagraphAlign = "left" | "center" | "right" | "justify";

/** Paragraph formatting. Omitted fields are left as they are. */
export interface ParagraphStyle {
  align?: ParagraphAlign;
}

/**
 * Apply paragraph formatting to a shape's text. Without `range` every
 * paragraph is formatted; with it, every paragraph the range touches.
 */
export interface SetParagraphStyleOperation {
  type: "setParagraphStyle";
  slideId: string;
  nodeId: string;
  /** The table cell holding the text, when the node is a table. */
  cell?: TableCellAddress;
  style: ParagraphStyle;
  range?: TextRange;
}

/**
 * Add a slide built from a layout, with the layout's content placeholders
 * (empty), the way PowerPoint's New Slide does.
 */
export interface AddSlideOperation {
  type: "addSlide";
  /**
   * Layout part path (a key of `presentation.layouts`, e.g.
   * "ppt/slideLayouts/slideLayout2.xml"). Defaults to the layout of
   * `afterSlideId`, or of the last slide.
   */
  layoutId?: string;
  /** Insert after this slide. Defaults to the end of the deck. */
  afterSlideId?: string;
}

/** Set a slide's own background, or clear it to inherit the layout's. */
export interface SetSlideBackgroundOperation {
  type: "setSlideBackground";
  slideId: string;
  /** Hex color, with or without leading `#`. `null` removes the slide's own background. */
  color: string | null;
}

/**
 * Copy a top-level node onto the same slide, just above the original in the
 * stacking order.
 */
export interface DuplicateNodeOperation {
  type: "duplicateNode";
  slideId: string;
  nodeId: string;
  /**
   * How far the copy sits from the original, in px.
   *
   * @default { x: 10, y: 10 }
   */
  offset?: NodePosition;
}

export type NodeOrder = "front" | "back" | "forward" | "backward";

/** Move a top-level node in the slide's stacking order. */
export interface ReorderNodeOperation {
  type: "reorderNode";
  slideId: string;
  nodeId: string;
  /** To the top or bottom of the stack, or one step up or down. */
  order: NodeOrder;
}

/** Insert a text box on top of a slide. */
export interface AddTextBoxOperation {
  type: "addTextBox";
  slideId: string;
  /** Size in px. Defaults to a third of the slide's width, one line tall. */
  size?: NodeSize;
  /** Position in px. Defaults to centered on the slide. */
  position?: NodePosition;
  /** Initial text. Defaults to empty. */
  text?: string;
}

/**
 * Insert a preset shape, styled from the theme like PowerPoint's: accent fill,
 * a darker outline, and light centered text.
 */
export interface AddShapeOperation {
  type: "addShape";
  slideId: string;
  /** Preset geometry (`a:prstGeom@prst`), e.g. "rect", "ellipse", "rightArrow". */
  geometry: string;
  /** Size in px. Defaults to a quarter of the slide's width, square. */
  size?: NodeSize;
  /** Position in px. Defaults to centered on the slide. */
  position?: NodePosition;
}

/** Set a shape's or picture's outline. Omitted fields are left as they are. */
export interface SetOutlineOperation {
  type: "setOutline";
  slideId: string;
  nodeId: string;
  /** Hex color, with or without leading `#`. `null` removes the outline. */
  color?: string | null;
  /** Width in points. */
  width?: number;
}

/** Insert a table on top of a slide, styled with the deck's default table style. */
export interface AddTableOperation {
  type: "addTable";
  slideId: string;
  rows: number;
  columns: number;
  /** Size in px. Defaults to two thirds of the slide's width, with 40px rows. */
  size?: NodeSize;
  /** Position in px. Defaults to centered on the slide. */
  position?: NodePosition;
}

/** Set a table's column widths and row heights. */
export interface SetTableGridOperation {
  type: "setTableGrid";
  slideId: string;
  nodeId: string;
  /** Width of every column in px, in order. */
  columnWidths?: number[];
  /** Height of every row in px, in order. Rows still grow to fit their text. */
  rowHeights?: number[];
}

export type TableCellVerticalAlign = "top" | "middle" | "bottom";

/** Cell formatting. Omitted fields are left as they are. */
export interface TableCellStyle {
  /** Hex fill color, or `null` for no fill. */
  fill?: string | null;
  /** Every border of the cells: a hex color and width in points, or `null` for none. */
  borders?: { color: string; width?: number } | null;
  verticalAlign?: TableCellVerticalAlign;
}

/** Format table cells: a rectangle between two cells, or every cell. */
export interface SetTableCellStyleOperation {
  type: "setTableCellStyle";
  slideId: string;
  nodeId: string;
  style: TableCellStyle;
  /** Opposite corners of the cells to format. Defaults to the whole table. */
  from?: TableCellAddress;
  to?: TableCellAddress;
}

/** Which parts of its table style a table shows. Omitted fields are left as they are. */
export interface TableOptions {
  headerRow?: boolean;
  totalRow?: boolean;
  bandedRows?: boolean;
  firstColumn?: boolean;
  lastColumn?: boolean;
  bandedColumns?: boolean;
}

export interface SetTableOptionsOperation {
  type: "setTableOptions";
  slideId: string;
  nodeId: string;
  options: TableOptions;
}

/** Merge a rectangle of table cells into one, keeping all their text. */
export interface MergeTableCellsOperation {
  type: "mergeTableCells";
  slideId: string;
  nodeId: string;
  /** Opposite corners of the cells to merge; the range grows to cover any merge it cuts. */
  from: TableCellAddress;
  to: TableCellAddress;
}

/** Split a merged table cell back into its cells. */
export interface SplitTableCellOperation {
  type: "splitTableCell";
  slideId: string;
  nodeId: string;
  cell: TableCellAddress;
}

/** Insert an empty row into a table, formatted like its neighbor. */
export interface InsertTableRowOperation {
  type: "insertTableRow";
  slideId: string;
  nodeId: string;
  /** Index the new row takes; the row count appends it. */
  index: number;
}

/** Remove a row from a table. A table keeps at least one row. */
export interface DeleteTableRowOperation {
  type: "deleteTableRow";
  slideId: string;
  nodeId: string;
  index: number;
}

/** Insert an empty column into a table, as wide as its neighbor. */
export interface InsertTableColumnOperation {
  type: "insertTableColumn";
  slideId: string;
  nodeId: string;
  /** Index the new column takes; the column count appends it. */
  index: number;
}

/** Remove a column from a table. A table keeps at least one column. */
export interface DeleteTableColumnOperation {
  type: "deleteTableColumn";
  slideId: string;
  nodeId: string;
  index: number;
}

/** Insert a picture on top of a slide. */
export interface AddImageOperation {
  type: "addImage";
  slideId: string;
  /** The image file's bytes. */
  data: Uint8Array;
  /** `"image/png"`, `"image/jpeg"`, or `"image/gif"`. */
  contentType: string;
  /** Size in px. Defaults to the image's own size, scaled down to fit the slide. */
  size?: NodeSize;
  /** Position in px. Defaults to centered on the slide. */
  position?: NodePosition;
  /** Shape name, also used as the picture's alt text. */
  name?: string;
}

/**
 * Applies several operations as one atomic edit with a single undo (e.g.
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
  | SetTextStyleOperation
  | SetParagraphStyleOperation
  | AddSlideOperation
  | SetSlideBackgroundOperation
  | DuplicateNodeOperation
  | ReorderNodeOperation
  | AddImageOperation
  | PasteNodesOperation
  | AddTextBoxOperation
  | AddShapeOperation
  | SetOutlineOperation
  | AddTableOperation
  | SetTableGridOperation
  | MergeTableCellsOperation
  | SplitTableCellOperation
  | SetTableCellStyleOperation
  | SetTableOptionsOperation
  | InsertTableRowOperation
  | DeleteTableRowOperation
  | InsertTableColumnOperation
  | DeleteTableColumnOperation
  | BatchOperation;

export interface EditResult {
  /** Slide ids whose rendered output may have changed. */
  affectedSlideIds: string[];
  /** Restores XML, model, and dirty state. Apply undos in reverse order. */
  undo: () => void;
  /** Id of the slide created by `duplicateSlide` or `addSlide`. */
  createdSlideId?: string;
  /**
   * Ids of the nodes created by `duplicateNode`, `addTextBox`, `addShape`,
   * `addTable`, `addImage`, or `pasteNodes`.
   */
  createdNodeIds?: string[];
}

/**
 * Applies an operation to a presentation opened with `keepSourcePackage: true`:
 * mutates the live part XML that `writePptx()` saves, updates the model, and
 * marks the touched parts dirty. Undos capture element references and part
 * snapshots, so apply them in reverse order.
 */
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
    case "setTextStyle":
      return applySetTextStyle(presentation, op);
    case "setParagraphStyle":
      return applySetParagraphStyle(presentation, op);
    case "addSlide":
      return applyAddSlide(presentation, op);
    case "setSlideBackground":
      return applySetSlideBackground(presentation, op);
    case "duplicateNode":
      return applyDuplicateNode(presentation, op);
    case "reorderNode":
      return applyReorderNode(presentation, op);
    case "addImage":
      return applyAddImage(presentation, op);
    case "pasteNodes":
      return applyPasteNodes(presentation, op);
    case "addTextBox":
      return applyAddTextBox(presentation, op);
    case "addShape":
      return applyAddShape(presentation, op);
    case "setOutline":
      return applySetOutline(presentation, op);
    case "addTable":
      return applyAddTable(presentation, op);
    case "setTableGrid":
      return applySetTableGrid(presentation, op);
    case "setTableCellStyle":
      return applySetTableCellStyle(presentation, op);
    case "setTableOptions":
      return applySetTableOptions(presentation, op);
    case "mergeTableCells":
      return applyMergeTableCells(presentation, op);
    case "splitTableCell":
      return applySplitTableCell(presentation, op);
    case "insertTableRow":
      return applyInsertTableRow(presentation, op);
    case "deleteTableRow":
      return applyDeleteTableRow(presentation, op);
    case "insertTableColumn":
      return applyInsertTableColumn(presentation, op);
    case "deleteTableColumn":
      return applyDeleteTableColumn(presentation, op);
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
    createdSlideId: results.reduce<string | undefined>(
      (id, result) => result.createdSlideId ?? id,
      undefined,
    ),
    createdNodeIds: results.some((result) => result.createdNodeIds)
      ? results.flatMap((result) => result.createdNodeIds ?? [])
      : undefined,
  };
}

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
  if (!slide) throw new Error(`applyEdit: unknown slide "${slideId}".`);
  return slide;
}

function findNode(pres: PresentationData, slide: SlideData, nodeId: string): SlideNode {
  materializeSlide(pres, slide);
  const node = slide.nodes.find((n) => n.id === nodeId);
  if (!node) {
    throw new Error(`applyEdit: no top-level node "${nodeId}" on slide "${slide.id}".`);
  }
  return node;
}

/**
 * The node's element in the slide's shape tree. A node's source can sit
 * deeper, such as the fallback picture inside an OLE object's frame.
 */
function topLevelElement(node: SlideNode): Element {
  let element = requireElement(node.source, "node element");
  while (element.parentElement && element.parentElement.localName !== "spTree") {
    element = element.parentElement;
  }
  return element;
}

function requireElement(node: SafeXmlNode, what: string): Element {
  const element = node.element;
  if (!element) throw new Error(`applyEdit: ${what} is missing.`);
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

interface TextTarget {
  textBody: TextBody;
  txBody: SafeXmlNode;
}

function findTextTarget(
  pres: PresentationData,
  slide: SlideData,
  nodeId: string,
  cell: TableCellAddress | undefined,
): TextTarget {
  const node = findNode(pres, slide, nodeId);
  if (node.nodeType === "table" && cell) {
    const table = node as TableNodeData;
    const modelCell = table.rows[cell.row]?.cells[cell.column];
    const tc = node.source
      .child("graphic")
      .child("graphicData")
      .child("tbl")
      .children("tr")
      [cell.row]?.children("tc")[cell.column];
    const txBody = tc?.child("txBody");
    if (
      !modelCell ||
      modelCell.hMerge ||
      modelCell.vMerge ||
      !modelCell.textBody ||
      !txBody?.exists()
    ) {
      throw new Error(
        `applyEdit: no editable cell at row ${cell.row}, column ${cell.column} in table "${nodeId}".`,
      );
    }
    return { textBody: modelCell.textBody, txBody };
  }
  if (node.nodeType !== "shape" || !node.textBody || cell) {
    throw new Error(`applyEdit: node "${nodeId}" has no editable text body.`);
  }
  return { textBody: (node as ShapeNodeData).textBody!, txBody: node.source.child("txBody") };
}

/** Localnames that produce entries in TextParagraph.runs, in document order. */
const RUN_LOCAL_NAMES = new Set(["r", "br", "tab", "fld"]);

function applySetTextRun(pres: PresentationData, op: SetTextRunOperation): EditResult {
  const sourcePackage = requirePkg(pres);
  const slide = findSlide(pres, op.slideId);
  const { textBody, txBody } = findTextTarget(pres, slide, op.nodeId, op.cell);

  const paragraph = textBody.paragraphs[op.paragraphIndex];
  const run = paragraph?.runs[op.runIndex];
  if (!paragraph || !run) {
    throw new Error(
      `applyEdit: no run at paragraph ${op.paragraphIndex}, run ${op.runIndex} in node "${op.nodeId}".`,
    );
  }

  // The model's runs array is built from an in-order scan of r/br/tab/fld
  // children, so runIndex maps onto the same scan of the XML.
  const pNode = txBody.children("p")[op.paragraphIndex];
  const runNodes = (pNode?.allChildren() ?? []).filter((c) => RUN_LOCAL_NAMES.has(c.localName));
  const runNode = runNodes[op.runIndex];
  if (!runNode) {
    throw new Error(`applyEdit: run XML not found for paragraph ${op.paragraphIndex}.`);
  }
  if (runNode.localName !== "r" && runNode.localName !== "fld") {
    throw new Error(`applyEdit: cannot set text on a <a:${runNode.localName}> (break/tab) run.`);
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

/**
 * The `a:p` children of a text body. A direct-children scan rather than
 * `getElementsByTagNameNS`, whose namespace matching is unreliable in some
 * DOM implementations (happy-dom misses the paragraphs entirely).
 */
function childParagraphs(txBodyElement: Element): Element[] {
  return Array.from(txBodyElement.children).filter((child) => child.localName === "p");
}

function applySetTextBody(pres: PresentationData, op: SetTextBodyOperation): EditResult {
  const sourcePackage = requirePkg(pres);
  const slide = findSlide(pres, op.slideId);
  const { textBody, txBody } = findTextTarget(pres, slide, op.nodeId, op.cell);

  const prevParagraphs = textBody.paragraphs;

  const txBodyElement = requireElement(txBody, "txBody");
  const doc = txBodyElement.ownerDocument;

  const origParagraphElements = childParagraphs(txBodyElement);
  const origParagraphs = prevParagraphs;

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

  for (const paragraphElement of origParagraphElements) {
    removeChild(paragraphElement);
  }

  const newParagraphs: TextParagraph[] = [];

  for (const opPara of op.paragraphs) {
    const paragraphElement = doc.createElementNS(A_NS, "a:p");

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

  textBody.paragraphs = newParagraphs;
  sourcePackage.markDirty(slide.id);

  return {
    affectedSlideIds: [slide.id],
    undo: () => {
      const currentParagraphElements = childParagraphs(txBodyElement);
      for (const paragraphElement of currentParagraphElements) {
        removeChild(paragraphElement);
      }
      for (const paragraphElement of origParagraphElements) {
        insertChild(txBodyElement, paragraphElement, null);
      }
      textBody.paragraphs = prevParagraphs;
      sourcePackage.markDirty(slide.id);
    },
  };
}

interface XfrmHandle {
  xfrmElement: Element;
  offElement: Element;
  extElement: Element;
  created: boolean;
}

/**
 * Finds the node's xfrm element, creating one when the node inherits its
 * transform (e.g. layout placeholders have no xfrm of their own; moving
 * them requires writing a new one into the slide part).
 */
function getOrCreateXfrm(node: SlideNode): XfrmHandle {
  return getOrCreateXfrmOn(node.source, node.nodeType, node.id);
}

function getOrCreateXfrmOn(
  source: SafeXmlNode,
  nodeType: SlideNode["nodeType"],
  label: string,
): XfrmHandle {
  const sourceElement = requireElement(source, "node element");
  const doc = sourceElement.ownerDocument;

  let container: Element;
  let xfrmQualified: string;
  let xfrmNs: string;
  let insertRef: Element | null;

  if (nodeType === "table" || nodeType === "chart") {
    // Graphic frames carry p:xfrm directly, between nvGraphicFramePr and a:graphic.
    container = sourceElement;
    xfrmQualified = "p:xfrm";
    xfrmNs = P_NS;
    insertRef = source.child("graphic").element;
  } else {
    const prName = nodeType === "group" ? "grpSpPr" : "spPr";
    container = requireElement(source.child(prName), `${prName} of node "${label}"`);
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
 * Scales a table's grid to a new frame size, as PowerPoint's resize handle does.
 * Tables lay out from `a:gridCol/@w` and `a:tr/@h`, not the graphicFrame extent,
 * so writing the extent alone would only shrink the selection box. Returns an
 * undo closure, or null when there is no scalable grid.
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
    throw new Error(`applyEdit: setSolidFill only supports shape nodes, got "${node.nodeType}".`);
  }
  const shape = node as ShapeNodeData;

  const color = parseHexColor(op.color);

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

/** rPr fill children (EG_FillProperties). */
const RPR_FILL_LOCAL_NAMES = new Set([
  "noFill",
  "solidFill",
  "gradFill",
  "blipFill",
  "pattFill",
  "grpFill",
]);
/** rPr children that come after `a:latin` in the schema sequence. */
const RPR_AFTER_LATIN_LOCAL_NAMES = new Set([
  "ea",
  "cs",
  "sym",
  "hlinkClick",
  "hlinkMouseOver",
  "rtl",
  "extLst",
]);
/** rPr children that come after the fill in the schema sequence. */
const RPR_AFTER_FILL_LOCAL_NAMES = new Set([
  "effectLst",
  "effectDag",
  "highlight",
  "uLnTx",
  "uLn",
  "uFillTx",
  "uFill",
  "latin",
  ...RPR_AFTER_LATIN_LOCAL_NAMES,
]);

const ALIGN_TO_OOXML: Record<ParagraphAlign, string> = {
  left: "l",
  center: "ctr",
  right: "r",
  justify: "just",
};

function parseHexColor(value: string): string {
  const match = /^#?([0-9a-fA-F]{6})$/.exec(value);
  if (!match) {
    throw new Error(`applyEdit: invalid color "${value}"; expected 6-digit hex.`);
  }
  return match[1].toUpperCase();
}

function firstChildIn(parent: Element, localNames: ReadonlySet<string>): Element | null {
  for (const child of parent.children) {
    if (localNames.has(child.localName)) return child;
  }
  return null;
}

function childNamed(parent: Element, localName: string): Element | null {
  for (const child of parent.children) {
    if (child.localName === localName) return child;
  }
  return null;
}

function runText(run: Element): string {
  return childNamed(run, "t")?.textContent ?? "";
}

/** Character length of a run-like element, matching `TextParagraph.runs` text. */
function runLength(run: Element): number {
  return run.localName === "br" || run.localName === "tab" ? 1 : runText(run).length;
}

/** Splits an `a:r` at `at` characters; the original keeps the left part. Returns the right part. */
function splitRun(run: Element, at: number): Element {
  const text = runText(run);
  const right = run.cloneNode(true) as Element;
  childNamed(run, "t")!.textContent = text.slice(0, at);
  childNamed(right, "t")!.textContent = text.slice(at);
  run.parentNode!.insertBefore(right, run.nextSibling);
  return right;
}

function ensureRunProperties(run: Element): Element {
  const existing = childNamed(run, "rPr");
  if (existing) return existing;
  const rPr = run.ownerDocument.createElementNS(A_NS, "a:rPr");
  // rPr is the first child of a:r, a:br and a:fld.
  run.insertBefore(rPr, run.firstElementChild);
  return rPr;
}

function setBoolAttr(element: Element, name: string, value: boolean | undefined): void {
  if (value !== undefined) element.setAttribute(name, value ? "1" : "0");
}

/** Writes `style` onto character properties (`a:rPr` or `a:endParaRPr`). */
function applyCharacterStyle(rPr: Element, style: TextStyle): void {
  const doc = rPr.ownerDocument;

  setBoolAttr(rPr, "b", style.bold);
  setBoolAttr(rPr, "i", style.italic);
  if (style.underline !== undefined) rPr.setAttribute("u", style.underline ? "sng" : "none");
  if (style.strikethrough !== undefined) {
    rPr.setAttribute("strike", style.strikethrough ? "sngStrike" : "noStrike");
  }
  if (style.fontSize !== undefined) {
    // OOXML allows 1–4000pt (`a:rPr@sz` 100–400000); PowerPoint rejects the rest.
    if (style.fontSize !== null && !(style.fontSize >= 1 && style.fontSize <= 4000)) {
      throw new Error(`applyEdit: invalid font size "${style.fontSize}".`);
    }
    setOrRemoveAttr(
      rPr,
      "sz",
      style.fontSize === null ? undefined : String(Math.round(style.fontSize * 100)),
    );
  }

  if (style.color !== undefined) {
    for (const child of Array.from(rPr.children)) {
      if (RPR_FILL_LOCAL_NAMES.has(child.localName)) child.remove();
    }
    if (style.color !== null) {
      const solidFill = doc.createElementNS(A_NS, "a:solidFill");
      const srgb = doc.createElementNS(A_NS, "a:srgbClr");
      srgb.setAttribute("val", parseHexColor(style.color));
      solidFill.appendChild(srgb);
      rPr.insertBefore(solidFill, firstChildIn(rPr, RPR_AFTER_FILL_LOCAL_NAMES));
    }
  }

  if (style.fontFamily !== undefined) {
    let latin = childNamed(rPr, "latin");
    if (style.fontFamily === null) {
      latin?.remove();
    } else {
      if (!latin) {
        latin = doc.createElementNS(A_NS, "a:latin");
        rPr.insertBefore(latin, firstChildIn(rPr, RPR_AFTER_LATIN_LOCAL_NAMES));
      }
      latin.setAttribute("typeface", style.fontFamily);
    }
  }
}

/**
 * Formats the characters `[from, to)` of a detached `a:p` clone, splitting runs
 * at the range edges. `whole` formats every run and the paragraph mark.
 */
function styleParagraph(
  paragraph: Element,
  from: number,
  to: number,
  style: TextStyle,
  whole: boolean,
): void {
  let position = 0;
  let hasRuns = false;

  for (const child of Array.from(paragraph.children)) {
    if (!RUN_LOCAL_NAMES.has(child.localName)) continue;
    hasRuns = true;
    const runStart = position;
    const runEnd = position + runLength(child);
    position = runEnd;
    if (child.localName === "tab") continue;

    if (whole) {
      applyCharacterStyle(ensureRunProperties(child), style);
      continue;
    }

    const lo = Math.max(from, runStart);
    const hi = Math.min(to, runEnd);
    if (lo >= hi) continue;

    // Fields and breaks can't be split: format them whole when touched.
    if (child.localName !== "r") {
      applyCharacterStyle(ensureRunProperties(child), style);
      continue;
    }

    let target = child;
    let targetStart = runStart;
    if (lo > targetStart) {
      target = splitRun(target, lo - targetStart);
      targetStart = lo;
    }
    if (hi < runEnd) splitRun(target, hi - targetStart);
    applyCharacterStyle(ensureRunProperties(target), style);
  }

  // The paragraph mark carries the formatting for text typed at the end of
  // the paragraph, and for an empty paragraph's first keystroke.
  let endParaRPr = childNamed(paragraph, "endParaRPr");
  if (!endParaRPr && !hasRuns) {
    endParaRPr = paragraph.ownerDocument.createElementNS(A_NS, "a:endParaRPr");
    paragraph.appendChild(endParaRPr);
  }
  if (endParaRPr && (whole || to >= position)) applyCharacterStyle(endParaRPr, style);
}

/** Ordered `[paragraphIndex, from, to]` spans a range covers, or every paragraph. */
function paragraphSpans(
  paragraphCount: number,
  range: TextRange | undefined,
): [index: number, from: number, to: number][] {
  if (!range) {
    return Array.from({ length: paragraphCount }, (_, index) => [index, 0, Infinity]);
  }
  let { start, end } = range;
  if (
    start.paragraphIndex > end.paragraphIndex ||
    (start.paragraphIndex === end.paragraphIndex && start.offset > end.offset)
  ) {
    [start, end] = [end, start];
  }
  if (start.paragraphIndex < 0 || end.paragraphIndex >= paragraphCount) {
    throw new Error(
      `applyEdit: text range paragraphs ${start.paragraphIndex}-${end.paragraphIndex} are out of bounds.`,
    );
  }
  const spans: [number, number, number][] = [];
  for (let index = start.paragraphIndex; index <= end.paragraphIndex; index++) {
    spans.push([
      index,
      index === start.paragraphIndex ? start.offset : 0,
      index === end.paragraphIndex ? end.offset : Infinity,
    ]);
  }
  return spans;
}

/**
 * Rewrites some of a text body's paragraphs. Each `a:p` is cloned, mutated
 * detached, and swapped in, so undo is a swap back and the model paragraph is
 * re-parsed from the clone.
 */
function rewriteParagraphs(
  pres: PresentationData,
  slideId: string,
  nodeId: string,
  cell: TableCellAddress | undefined,
  rewrite: (paragraphCount: number) => [index: number, mutate: (paragraph: Element) => void][],
): EditResult {
  const sourcePackage = requirePkg(pres);
  const slide = findSlide(pres, slideId);
  const { textBody, txBody } = findTextTarget(pres, slide, nodeId, cell);
  const txBodyElement = requireElement(txBody, "txBody");
  const paragraphElements = childParagraphs(txBodyElement);

  const prevParagraphs = textBody.paragraphs;
  const nextParagraphs = [...prevParagraphs];
  const swaps: [original: Element, clone: Element][] = [];

  for (const [index, mutate] of rewrite(paragraphElements.length)) {
    const original = paragraphElements[index];
    if (!original) continue;
    const clone = original.cloneNode(true) as Element;
    mutate(clone);
    insertChild(txBodyElement, clone, original);
    removeChild(original);
    nextParagraphs[index] = parseParagraph(new SafeXmlNode(clone));
    swaps.push([original, clone]);
  }

  textBody.paragraphs = nextParagraphs;
  sourcePackage.markDirty(slide.id);

  return {
    affectedSlideIds: [slide.id],
    undo: () => {
      for (const [original, clone] of swaps.reverse()) {
        insertChild(txBodyElement, original, clone);
        removeChild(clone);
      }
      textBody.paragraphs = prevParagraphs;
      sourcePackage.markDirty(slide.id);
    },
  };
}

function applySetTextStyle(pres: PresentationData, op: SetTextStyleOperation): EditResult {
  // Fails on a bad value before any paragraph is touched.
  if (op.style.color) parseHexColor(op.style.color);
  return rewriteParagraphs(pres, op.slideId, op.nodeId, op.cell, (count) =>
    paragraphSpans(count, op.range).map(([index, from, to]) => [
      index,
      (paragraph) => styleParagraph(paragraph, from, to, op.style, !op.range),
    ]),
  );
}

function applySetParagraphStyle(
  pres: PresentationData,
  op: SetParagraphStyleOperation,
): EditResult {
  return rewriteParagraphs(pres, op.slideId, op.nodeId, op.cell, (count) =>
    paragraphSpans(count, op.range).map(([index]) => [
      index,
      (paragraph) => {
        let pPr = childNamed(paragraph, "pPr");
        if (!pPr) {
          pPr = paragraph.ownerDocument.createElementNS(A_NS, "a:pPr");
          // pPr is the first child of a:p.
          paragraph.insertBefore(pPr, paragraph.firstElementChild);
        }
        if (op.style.align) pPr.setAttribute("algn", ALIGN_TO_OOXML[op.style.align]);
      },
    ]),
  );
}

function applyDeleteNode(pres: PresentationData, op: DeleteNodeOperation): EditResult {
  const sourcePackage = requirePkg(pres);
  const slide = findSlide(pres, op.slideId);
  const node = findNode(pres, slide, op.nodeId);

  const element = topLevelElement(node);
  const parentElement = element.parentElement;
  if (!parentElement)
    throw new Error(`applyEdit: node "${op.nodeId}" is not attached to the slide.`);

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

const DEFAULT_DUPLICATE_OFFSET: NodePosition = { x: 10, y: 10 };

function descendantsNamed(root: Element, localName: string): Element[] {
  const found: Element[] = [];
  const visit = (element: Element) => {
    if (element.localName === localName) found.push(element);
    for (const child of element.children) visit(child);
  };
  visit(root);
  return found;
}

/** The slide's spTree element; new nodes go in it. */
function requireSpTree(pres: PresentationData, slide: SlideData): Element {
  const root = requirePkg(pres).getXmlRoot(slide.id);
  const spTree = root?.child("cSld").child("spTree");
  if (!spTree) throw new Error(`applyEdit: slide part "${slide.id}" is not loaded.`);
  return requireElement(spTree, "spTree");
}

/** The next free shape id on a slide. Ids must be unique across the whole part. */
function nextShapeId(spTree: Element): number {
  let max = 0;
  for (const cNvPr of descendantsNamed(spTree, "cNvPr")) {
    const id = Number(cNvPr.getAttribute("id"));
    if (Number.isFinite(id)) max = Math.max(max, id);
  }
  return max + 1;
}

/**
 * Gives every shape in a copied element a fresh id from `nextId`, a group's
 * children included, and points connectors inside it at the copies.
 * Connections to shapes outside the copy are left alone.
 */
function renumberShapes(copy: Element, nextId: number): number {
  const ids = new Map<string, string>();
  for (const cNvPr of descendantsNamed(copy, "cNvPr")) {
    const id = String(nextId++);
    const previous = cNvPr.getAttribute("id");
    if (previous) ids.set(previous, id);
    cNvPr.setAttribute("id", id);
  }
  for (const connection of [
    ...descendantsNamed(copy, "stCxn"),
    ...descendantsNamed(copy, "endCxn"),
  ]) {
    const id = ids.get(connection.getAttribute("id") ?? "");
    if (id) connection.setAttribute("id", id);
  }
  return nextId;
}

/** Parses a node element added to a slide, with its placeholder inheritance resolved. */
function parseAddedNode(pres: PresentationData, slide: SlideData, element: Element): SlideNode {
  const node = parseRenderableChild(new SafeXmlNode(element), {
    rels: slide.rels,
    partPath: slide.id,
    diagramDrawings: pres.diagramDrawings,
  });
  if (!node) throw new Error(`applyEdit: <${element.localName}> is not a renderable node.`);
  resolveSlideNodePlaceholders(pres, slide, [node]);
  return node;
}

async function applyDuplicateNode(
  pres: PresentationData,
  op: DuplicateNodeOperation,
): Promise<EditResult> {
  const sourcePackage = requirePkg(pres);
  const slide = findSlide(pres, op.slideId);
  const node = findNode(pres, slide, op.nodeId);
  const element = topLevelElement(node);
  const parentElement = element.parentElement;
  if (!parentElement) {
    throw new Error(`applyEdit: node "${op.nodeId}" is not attached to the slide.`);
  }

  const copy = element.cloneNode(true) as Element;
  renumberShapes(copy, nextShapeId(requireSpTree(pres, slide)));
  insertChild(parentElement, copy, element.nextElementSibling);

  const copyNode = parseAddedNode(pres, slide, copy);
  const index = slide.nodes.indexOf(node) + 1;
  slide.nodes.splice(index, 0, copyNode);

  // The copy gets its full transform: an inherited placeholder position has to
  // become its own before it can be offset.
  const offset = op.offset ?? DEFAULT_DUPLICATE_OFFSET;
  const transform = applySetNodeTransform(pres, {
    type: "setNodeTransform",
    slideId: slide.id,
    nodeId: copyNode.id,
    position: { x: node.position.x + offset.x, y: node.position.y + offset.y },
    size: { ...node.size },
  });
  sourcePackage.markDirty(slide.id);

  return {
    affectedSlideIds: [slide.id],
    createdNodeIds: [copyNode.id],
    undo: () => {
      transform.undo();
      removeChild(copy);
      slide.nodes.splice(slide.nodes.indexOf(copyNode), 1);
      sourcePackage.markDirty(slide.id);
    },
  };
}

function applyReorderNode(pres: PresentationData, op: ReorderNodeOperation): EditResult {
  const sourcePackage = requirePkg(pres);
  const slide = findSlide(pres, op.slideId);
  const node = findNode(pres, slide, op.nodeId);
  const element = topLevelElement(node);
  const parentElement = element.parentElement;
  if (!parentElement) {
    throw new Error(`applyEdit: node "${op.nodeId}" is not attached to the slide.`);
  }

  const fromIndex = slide.nodes.indexOf(node);
  const last = slide.nodes.length - 1;
  const toIndex = {
    front: last,
    back: 0,
    forward: Math.min(fromIndex + 1, last),
    backward: Math.max(fromIndex - 1, 0),
  }[op.order];
  if (toIndex === fromIndex) return { affectedSlideIds: [], undo: () => {} };

  // Place the element next to the node it swaps with, so XML children the
  // model doesn't render keep their places.
  const neighbor = topLevelElement(slide.nodes[toIndex]);
  const originalNext = element.nextElementSibling;
  removeChild(element);
  insertChild(parentElement, element, toIndex > fromIndex ? neighbor.nextElementSibling : neighbor);
  slide.nodes.splice(fromIndex, 1);
  slide.nodes.splice(toIndex, 0, node);
  sourcePackage.markDirty(slide.id);

  return {
    affectedSlideIds: [slide.id],
    undo: () => {
      removeChild(element);
      insertChild(parentElement, element, originalNext);
      slide.nodes.splice(slide.nodes.indexOf(node), 1);
      slide.nodes.splice(fromIndex, 0, node);
      sourcePackage.markDirty(slide.id);
    },
  };
}

const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const IMAGE_EXTENSIONS: Record<string, string> = {
  "image/png": "png",
  "image/jpeg": "jpeg",
  "image/gif": "gif",
};

/** Pixel size from a PNG, GIF, or JPEG header, or undefined when it can't be read. */
export function readImageSize(data: Uint8Array): NodeSize | undefined {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  if (data.length >= 24 && view.getUint32(0) === 0x89504e47) {
    return { w: view.getUint32(16), h: view.getUint32(20) };
  }
  if (data.length >= 10 && data[0] === 0x47 && data[1] === 0x49 && data[2] === 0x46) {
    return { w: view.getUint16(6, true), h: view.getUint16(8, true) };
  }
  if (data.length >= 4 && data[0] === 0xff && data[1] === 0xd8) {
    let offset = 2;
    while (offset + 9 < data.length) {
      if (data[offset] !== 0xff) return undefined;
      const marker = data[offset + 1];
      const length = view.getUint16(offset + 2);
      // SOF0-SOF15 carry the frame size; C4, C8, and CC are other segments.
      if (
        marker >= 0xc0 &&
        marker <= 0xcf &&
        marker !== 0xc4 &&
        marker !== 0xc8 &&
        marker !== 0xcc
      ) {
        return { w: view.getUint16(offset + 7), h: view.getUint16(offset + 5) };
      }
      offset += 2 + length;
    }
  }
  return undefined;
}

/**
 * Media paths handed out this session. An undone image frees its path in the
 * package, but a renderer may still cache a blob URL under it, so a new image
 * must not take it over.
 */
const issuedMediaPaths = new WeakMap<PresentationData, Set<string>>();

function nextMediaPath(pres: PresentationData, extension: string): string {
  let issued = issuedMediaPaths.get(pres);
  if (!issued) {
    issued = new Set();
    issuedMediaPaths.set(pres, issued);
  }
  const taken = new Set([...requirePkg(pres).paths(), ...pres.media.keys(), ...issued]);
  let index = 1;
  while (taken.has(`ppt/media/image${index}.${extension}`)) index++;
  const path = `ppt/media/image${index}.${extension}`;
  issued.add(path);
  return path;
}

interface PartChange {
  undo: () => void;
}

/** Adds a media part, declaring its extension's content type when the deck lacks it. */
async function addMediaPart(
  pres: PresentationData,
  extension: string,
  contentType: string,
  data: Uint8Array,
): Promise<PartChange & { path: string; target: string }> {
  const sourcePackage = requirePkg(pres);
  const path = nextMediaPath(pres, extension);
  sourcePackage.setEntry(path, data);
  pres.media.set(path, data);

  const ctText = (await sourcePackage.readText(CONTENT_TYPES_PATH)) ?? "";
  const ctRoot = parseXml(ctText);
  const ctRootElement = requireElement(ctRoot, "[Content_Types].xml root");
  const hasDefault = ctRoot
    .children("Default")
    .some((entry) => entry.attr("Extension")?.toLowerCase() === extension);
  if (!hasDefault) {
    const defaultElement = ctRootElement.ownerDocument.createElementNS(CT_NS, "Default");
    defaultElement.setAttribute("Extension", extension);
    defaultElement.setAttribute("ContentType", contentType);
    insertChild(ctRootElement, defaultElement, ctRootElement.firstElementChild);
    sourcePackage.setEntry(CONTENT_TYPES_PATH, serializePartText(ctRootElement, ctText));
  }

  return {
    path,
    // Slides live in ppt/slides/, so media is one level up.
    target: `../media/${path.slice("ppt/media/".length)}`,
    undo: () => {
      sourcePackage.deleteEntry(path);
      pres.media.delete(path);
      if (!hasDefault) sourcePackage.setEntry(CONTENT_TYPES_PATH, ctText);
    },
  };
}

/** Adds relationships to a slide's rels part, creating the part when missing. */
async function addSlideRels(
  pres: PresentationData,
  slide: SlideData,
  entries: RelEntry[],
): Promise<PartChange & { ids: string[] }> {
  const sourcePackage = requirePkg(pres);
  const relsPath = relsPathFor(slide.id);
  const existingText = await sourcePackage.readText(relsPath);
  const relsText =
    existingText ??
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="${RELS_NS}"/>`;
  const relsRootElement = requireElement(parseXml(relsText), "slide rels root");
  const taken = new Map([...slide.rels, ...parseRels(relsText)]);

  const ids = entries.map((entry) => {
    const id = nextRelId(taken);
    taken.set(id, entry);
    const relElement = relsRootElement.ownerDocument.createElementNS(RELS_NS, "Relationship");
    relElement.setAttribute("Id", id);
    relElement.setAttribute("Type", entry.type);
    relElement.setAttribute("Target", entry.target);
    if (entry.targetMode) relElement.setAttribute("TargetMode", entry.targetMode);
    insertChild(relsRootElement, relElement, null);
    slide.rels.set(id, entry);
    return id;
  });
  sourcePackage.setEntry(relsPath, serializePartText(relsRootElement, relsText));

  return {
    ids,
    undo: () => {
      for (const id of ids) slide.rels.delete(id);
      if (existingText === undefined) sourcePackage.deleteEntry(relsPath);
      else sourcePackage.setEntry(relsPath, existingText);
    },
  };
}

/** A relationship copied nodes refer to, with the media part's bytes when it's internal. */
export interface ClipboardRelationship {
  type: string;
  target: string;
  targetMode?: string;
  media?: { data: Uint8Array; contentType: string; extension: string };
}

/** Nodes copied off a slide, self-contained so they paste into any deck. */
export interface NodeClipboard {
  /** Each copied node's XML, in stacking order. */
  nodes: string[];
  /** Relationships the XML refers to, keyed by the ids it uses. */
  relationships: Record<string, ClipboardRelationship>;
  /** Nodes left out because they refer to parts other than images (charts, diagrams, video). */
  omittedNodeIds: string[];
}

function relationshipIdsIn(element: Element): string[] {
  const ids: string[] = [];
  const visit = (current: Element) => {
    for (const attr of Array.from(current.attributes)) {
      if (attr.namespaceURI === R_NS && attr.value) ids.push(attr.value);
    }
    for (const child of current.children) visit(child);
  };
  visit(element);
  return ids;
}

function contentTypeOf(ctRoot: SafeXmlNode, partPath: string): string | undefined {
  const override = ctRoot.children("Override").find((o) => o.attr("PartName") === `/${partPath}`);
  if (override) return override.attr("ContentType");
  const extension = partPath.slice(partPath.lastIndexOf(".") + 1).toLowerCase();
  return ctRoot
    .children("Default")
    .find((entry) => entry.attr("Extension")?.toLowerCase() === extension)
    ?.attr("ContentType");
}

/**
 * Copies top-level nodes for `pasteNodes`, in their stacking order. Nodes that
 * inherit their position (placeholders) get it written out, and the image
 * parts they use travel with them.
 */
export async function copyNodes(
  pres: PresentationData,
  slideId: string,
  nodeIds: string[],
): Promise<NodeClipboard> {
  const sourcePackage = requirePkg(pres);
  const slide = findSlide(pres, slideId);
  materializeSlide(pres, slide);
  const serializer = new XMLSerializer();

  // Everything is captured before the first await, since a cut deletes the
  // nodes right after copying them.
  const captured = slide.nodes
    .filter((node) => nodeIds.includes(node.id))
    .map((node) => {
      const element = topLevelElement(node);
      const copy = element.cloneNode(true) as Element;
      const xfrm = getOrCreateXfrmOn(new SafeXmlNode(copy), node.nodeType, node.id);
      if (xfrm.created) {
        xfrm.offElement.setAttribute("x", String(pxToEmu(node.position.x)));
        xfrm.offElement.setAttribute("y", String(pxToEmu(node.position.y)));
        xfrm.extElement.setAttribute("cx", String(pxToEmu(node.size.w)));
        xfrm.extElement.setAttribute("cy", String(pxToEmu(node.size.h)));
      }
      const rels = relationshipIdsIn(element).flatMap((id) => {
        const rel = slide.rels.get(id);
        return rel ? [[id, rel] as const] : [];
      });
      return { id: node.id, xml: serializer.serializeToString(copy), rels };
    });

  const ctRoot = parseXml((await sourcePackage.readText(CONTENT_TYPES_PATH)) ?? "");
  const clipboard: NodeClipboard = { nodes: [], relationships: {}, omittedNodeIds: [] };

  for (const node of captured) {
    const relationships: Record<string, ClipboardRelationship> = {};
    let supported = true;

    for (const [id, rel] of node.rels) {
      if (isExternalTargetMode(rel.targetMode)) {
        relationships[id] = { type: rel.type, target: rel.target, targetMode: rel.targetMode };
        continue;
      }
      if (rel.type !== IMAGE_REL_TYPE) {
        supported = false;
        break;
      }
      const partPath = resolveRelTarget("ppt/slides", rel.target);
      const data = pres.media.get(partPath) ?? (await sourcePackage.readBytes(partPath));
      const extension = partPath.slice(partPath.lastIndexOf(".") + 1).toLowerCase();
      const contentType = contentTypeOf(ctRoot, partPath);
      if (!data || !contentType) {
        supported = false;
        break;
      }
      relationships[id] = {
        type: rel.type,
        target: rel.target,
        media: { data, contentType, extension },
      };
    }
    if (!supported) {
      clipboard.omittedNodeIds.push(node.id);
      continue;
    }

    clipboard.nodes.push(node.xml);
    Object.assign(clipboard.relationships, relationships);
  }
  return clipboard;
}

/** Paste nodes from `copyNodes`, from this deck or another, on top of a slide. */
export interface PasteNodesOperation {
  type: "pasteNodes";
  slideId: string;
  clipboard: NodeClipboard;
  /**
   * How far the pasted nodes sit from where they were copied, in px.
   *
   * @default { x: 0, y: 0 }
   */
  offset?: NodePosition;
}

/** A media part in the deck holding exactly these bytes, if there is one. */
async function findMediaPart(
  pres: PresentationData,
  data: Uint8Array,
): Promise<string | undefined> {
  const sourcePackage = requirePkg(pres);
  for (const path of sourcePackage.paths()) {
    if (!path.startsWith("ppt/media/")) continue;
    const existing = pres.media.get(path) ?? (await sourcePackage.readBytes(path));
    if (existing?.length === data.length && existing.every((byte, i) => byte === data[i])) {
      return path;
    }
  }
  return undefined;
}

async function applyPasteNodes(
  pres: PresentationData,
  op: PasteNodesOperation,
): Promise<EditResult> {
  const sourcePackage = requirePkg(pres);
  const slide = findSlide(pres, op.slideId);
  materializeSlide(pres, slide);
  const spTree = requireSpTree(pres, slide);
  const partChanges: PartChange[] = [];
  const transformChanges: PartChange[] = [];

  // Media the deck already has (always, when pasting within it) is reused.
  const oldIds = Object.keys(op.clipboard.relationships);
  const entries: RelEntry[] = [];
  for (const id of oldIds) {
    const rel = op.clipboard.relationships[id];
    if (!rel.media) {
      entries.push({ type: rel.type, target: rel.target, targetMode: rel.targetMode });
      continue;
    }
    const existing = await findMediaPart(pres, rel.media.data);
    if (existing) {
      entries.push({ type: rel.type, target: `../media/${existing.slice("ppt/media/".length)}` });
      continue;
    }
    const media = await addMediaPart(
      pres,
      rel.media.extension,
      rel.media.contentType,
      rel.media.data,
    );
    partChanges.push(media);
    entries.push({ type: rel.type, target: media.target });
  }
  const rels = await addSlideRels(pres, slide, entries);
  partChanges.push(rels);
  const newIds = new Map(oldIds.map((id, index) => [id, rels.ids[index]]));

  const extLst = Array.from(spTree.children).find((child) => child.localName === "extLst") ?? null;
  let nextId = nextShapeId(spTree);
  const pasted: { element: Element; node: SlideNode }[] = [];
  for (const xml of op.clipboard.nodes) {
    const element = spTree.ownerDocument.importNode(
      requireElement(parseXml(xml), "node"),
      true,
    ) as Element;
    const visit = (current: Element) => {
      for (const attr of Array.from(current.attributes)) {
        if (attr.namespaceURI === R_NS) attr.value = newIds.get(attr.value) ?? attr.value;
      }
      for (const child of current.children) visit(child);
    };
    visit(element);
    nextId = renumberShapes(element, nextId);
    insertChild(spTree, element, extLst);
    const node = parseAddedNode(pres, slide, element);
    slide.nodes.push(node);
    pasted.push({ element, node });

    if (op.offset && (op.offset.x !== 0 || op.offset.y !== 0)) {
      transformChanges.push(
        applySetNodeTransform(pres, {
          type: "setNodeTransform",
          slideId: slide.id,
          nodeId: node.id,
          position: { x: node.position.x + op.offset.x, y: node.position.y + op.offset.y },
        }),
      );
    }
  }
  sourcePackage.markDirty(slide.id);

  return {
    affectedSlideIds: [slide.id],
    createdNodeIds: pasted.map(({ node }) => node.id),
    undo: () => {
      for (const change of transformChanges.reverse()) change.undo();
      for (const { element, node } of pasted) {
        removeChild(element);
        slide.nodes.splice(slide.nodes.indexOf(node), 1);
      }
      for (const change of partChanges.reverse()) change.undo();
      sourcePackage.markDirty(slide.id);
    },
  };
}

/** One line of 18pt text, plus the default insets. */
const TEXT_BOX_HEIGHT = 40;

type ElementBuilder = (
  qualifiedName: string,
  attrs?: Record<string, string>,
  ...children: Node[]
) => Element;

function elementBuilder(doc: Document): ElementBuilder {
  return (qualifiedName, attrs = {}, ...children) => {
    const element = doc.createElementNS(
      qualifiedName.startsWith("a:") ? A_NS : P_NS,
      qualifiedName,
    );
    for (const [attr, value] of Object.entries(attrs)) element.setAttribute(attr, value);
    for (const child of children) element.appendChild(child);
    return element;
  };
}

function xfrm(el: ElementBuilder, position: NodePosition, size: NodeSize): Element {
  return el(
    "a:xfrm",
    {},
    el("a:off", { x: String(pxToEmu(position.x)), y: String(pxToEmu(position.y)) }),
    el("a:ext", { cx: String(pxToEmu(size.w)), cy: String(pxToEmu(size.h)) }),
  );
}

/** Appends a new top-level shape built by `build`, given its id, to a slide. */
function insertNewShape(
  pres: PresentationData,
  slideId: string,
  build: (el: ElementBuilder, id: number) => Element,
): EditResult {
  const sourcePackage = requirePkg(pres);
  const slide = findSlide(pres, slideId);
  materializeSlide(pres, slide);
  const spTree = requireSpTree(pres, slide);

  const sp = build(elementBuilder(spTree.ownerDocument), nextShapeId(spTree));
  const extLst = Array.from(spTree.children).find((child) => child.localName === "extLst");
  insertChild(spTree, sp, extLst ?? null);
  const node = parseAddedNode(pres, slide, sp);
  slide.nodes.push(node);
  sourcePackage.markDirty(slide.id);

  return {
    affectedSlideIds: [slide.id],
    createdNodeIds: [node.id],
    undo: () => {
      removeChild(sp);
      slide.nodes.splice(slide.nodes.indexOf(node), 1);
      sourcePackage.markDirty(slide.id);
    },
  };
}

function centered(pres: PresentationData, size: NodeSize): NodePosition {
  return { x: (pres.width - size.w) / 2, y: (pres.height - size.h) / 2 };
}

function applyAddTextBox(pres: PresentationData, op: AddTextBoxOperation): EditResult {
  const size = op.size ?? { w: pres.width / 3, h: TEXT_BOX_HEIGHT };
  const position = op.position ?? centered(pres, size);
  return insertNewShape(pres, op.slideId, (el, id) => {
    const paragraph = op.text
      ? el(
          "a:p",
          {},
          el("a:r", {}, el("a:rPr", { lang: "en-US" }), el("a:t")),
          el("a:endParaRPr", { lang: "en-US" }),
        )
      : el("a:p", {}, el("a:endParaRPr", { lang: "en-US" }));
    const text = paragraph.getElementsByTagNameNS(A_NS, "t")[0];
    if (text) text.textContent = op.text ?? "";
    return el(
      "p:sp",
      {},
      el(
        "p:nvSpPr",
        {},
        el("p:cNvPr", { id: String(id), name: `TextBox ${id - 1}` }),
        el("p:cNvSpPr", { txBox: "1" }),
        el("p:nvPr"),
      ),
      el(
        "p:spPr",
        {},
        xfrm(el, position, size),
        el("a:prstGeom", { prst: "rect" }, el("a:avLst")),
        el("a:noFill"),
      ),
      el(
        "p:txBody",
        {},
        el("a:bodyPr", { wrap: "square", rtlCol: "0" }, el("a:spAutoFit")),
        el("a:lstStyle"),
        paragraph,
      ),
    );
  });
}

const PRESET_NAME = /^[a-zA-Z0-9]+$/;

function applyAddShape(pres: PresentationData, op: AddShapeOperation): EditResult {
  if (!PRESET_NAME.test(op.geometry)) {
    throw new Error(`applyEdit: "${op.geometry}" is not a preset geometry name.`);
  }
  const side = pres.width / 4;
  const size = op.size ?? { w: side, h: side };
  const position = op.position ?? centered(pres, size);
  const name = op.geometry.charAt(0).toUpperCase() + op.geometry.slice(1);
  return insertNewShape(pres, op.slideId, (el, id) =>
    el(
      "p:sp",
      {},
      el(
        "p:nvSpPr",
        {},
        el("p:cNvPr", { id: String(id), name: `${name} ${id - 1}` }),
        el("p:cNvSpPr"),
        el("p:nvPr"),
      ),
      el(
        "p:spPr",
        {},
        xfrm(el, position, size),
        el("a:prstGeom", { prst: op.geometry }, el("a:avLst")),
      ),
      el(
        "p:style",
        {},
        el(
          "a:lnRef",
          { idx: "2" },
          el("a:schemeClr", { val: "accent1" }, el("a:shade", { val: "50000" })),
        ),
        el("a:fillRef", { idx: "1" }, el("a:schemeClr", { val: "accent1" })),
        el("a:effectRef", { idx: "0" }, el("a:schemeClr", { val: "accent1" })),
        el("a:fontRef", { idx: "minor" }, el("a:schemeClr", { val: "lt1" })),
      ),
      el(
        "p:txBody",
        {},
        el("a:bodyPr", { rtlCol: "0", anchor: "ctr" }),
        el("a:lstStyle"),
        el("a:p", {}, el("a:pPr", { algn: "ctr" }), el("a:endParaRPr", { lang: "en-US" })),
      ),
    ),
  );
}

const TABLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/table";
const DEFAULT_TABLE_STYLE_ID = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
const TABLE_ROW_HEIGHT = 40;
const MAX_TABLE_TRACKS = 75;

function applyAddTable(pres: PresentationData, op: AddTableOperation): EditResult {
  for (const count of [op.rows, op.columns]) {
    if (!Number.isInteger(count) || count < 1 || count > MAX_TABLE_TRACKS) {
      throw new Error(`applyEdit: invalid table size ${op.rows}x${op.columns}.`);
    }
  }
  const size = op.size ?? { w: (pres.width * 2) / 3, h: op.rows * TABLE_ROW_HEIGHT };
  const position = op.position ?? centered(pres, size);
  const styleId = pres.tableStyles?.attr("def") || DEFAULT_TABLE_STYLE_ID;
  const columnWidth = String(pxToEmu(size.w / op.columns));
  const rowHeight = String(pxToEmu(size.h / op.rows));

  return insertNewShape(pres, op.slideId, (el, id) => {
    const cell = () =>
      el(
        "a:tc",
        {},
        el(
          "a:txBody",
          {},
          el("a:bodyPr"),
          el("a:lstStyle"),
          el("a:p", {}, el("a:endParaRPr", { lang: "en-US" })),
        ),
        el("a:tcPr"),
      );
    const tableStyleId = el("a:tableStyleId");
    tableStyleId.textContent = styleId;
    return el(
      "p:graphicFrame",
      {},
      el(
        "p:nvGraphicFramePr",
        {},
        el("p:cNvPr", { id: String(id), name: `Table ${id - 1}` }),
        el("p:cNvGraphicFramePr", {}, el("a:graphicFrameLocks", { noGrp: "1" })),
        el("p:nvPr"),
      ),
      el(
        "p:xfrm",
        {},
        el("a:off", { x: String(pxToEmu(position.x)), y: String(pxToEmu(position.y)) }),
        el("a:ext", { cx: String(pxToEmu(size.w)), cy: String(pxToEmu(size.h)) }),
      ),
      el(
        "a:graphic",
        {},
        el(
          "a:graphicData",
          { uri: TABLE_URI },
          el(
            "a:tbl",
            {},
            el("a:tblPr", { firstRow: "1", bandRow: "1" }, tableStyleId),
            el(
              "a:tblGrid",
              {},
              ...Array.from({ length: op.columns }, () => el("a:gridCol", { w: columnWidth })),
            ),
            ...Array.from({ length: op.rows }, () =>
              el("a:tr", { h: rowHeight }, ...Array.from({ length: op.columns }, cell)),
            ),
          ),
        ),
      ),
    );
  });
}

function rewriteTable(
  pres: PresentationData,
  slideId: string,
  nodeId: string,
  mutate: (tbl: Element) => void,
): EditResult {
  const sourcePackage = requirePkg(pres);
  const slide = findSlide(pres, slideId);
  const node = findNode(pres, slide, nodeId);
  const original = requireElement(node.source, "table frame");
  if (node.nodeType !== "table" || original.localName !== "graphicFrame") {
    throw new Error(`applyEdit: node "${nodeId}" is not a table.`);
  }
  const parent = original.parentElement;
  if (!parent) throw new Error(`applyEdit: node "${nodeId}" is not attached to the slide.`);

  const clone = original.cloneNode(true) as Element;
  const tbl = descendantsNamed(clone, "tbl")[0];
  if (!tbl) throw new Error(`applyEdit: table "${nodeId}" has no a:tbl.`);
  mutate(tbl);

  const ext = childNamed(childNamed(clone, "xfrm") ?? clone, "ext");
  if (ext) {
    const sum = (elements: Element[], name: string) =>
      elements.reduce((total, element) => total + Number(element.getAttribute(name) ?? 0), 0);
    ext.setAttribute("cx", String(sum(gridColumns(tbl), "w")));
    ext.setAttribute("cy", String(sum(tableRows(tbl), "h")));
  }

  insertChild(parent, clone, original);
  removeChild(original);
  const next = parseAddedNode(pres, slide, clone);
  const index = slide.nodes.indexOf(node);
  slide.nodes[index] = next;
  sourcePackage.markDirty(slide.id);

  return {
    affectedSlideIds: [slide.id],
    undo: () => {
      insertChild(parent, original, clone);
      removeChild(clone);
      slide.nodes[index] = node;
      sourcePackage.markDirty(slide.id);
    },
  };
}

function gridColumns(tbl: Element): Element[] {
  const grid = childNamed(tbl, "tblGrid");
  return grid ? Array.from(grid.children).filter((child) => child.localName === "gridCol") : [];
}

function tableRows(tbl: Element): Element[] {
  return Array.from(tbl.children).filter((child) => child.localName === "tr");
}

function rowCells(tr: Element): Element[] {
  return Array.from(tr.children).filter((child) => child.localName === "tc");
}

function spanOf(tc: Element, name: "gridSpan" | "rowSpan"): number {
  return Number(tc.getAttribute(name) ?? 1);
}

function setSpan(tc: Element, name: "gridSpan" | "rowSpan", span: number): void {
  setOrRemoveAttr(tc, name, span > 1 ? String(span) : undefined);
}

function clearCellText(tc: Element): void {
  const doc = tc.ownerDocument;
  let txBody = childNamed(tc, "txBody");
  if (!txBody) {
    txBody = doc.createElementNS(A_NS, "a:txBody");
    txBody.append(doc.createElementNS(A_NS, "a:bodyPr"), doc.createElementNS(A_NS, "a:lstStyle"));
    tc.insertBefore(txBody, tc.firstElementChild);
  }
  const paragraphs = childParagraphs(txBody);
  const first = paragraphs[0];
  const paragraph = doc.createElementNS(A_NS, "a:p");
  const pPr = first ? childNamed(first, "pPr") : null;
  if (pPr) paragraph.appendChild(pPr.cloneNode(true));
  const firstRun = first ? firstChildIn(first, RUN_LOCAL_NAMES) : null;
  const marks = first ? childNamed(first, "endParaRPr") : null;
  const source = marks ?? (firstRun ? childNamed(firstRun, "rPr") : null);
  const endParaRPr = doc.createElementNS(A_NS, "a:endParaRPr");
  if (source) {
    for (const attr of Array.from(source.attributes)) {
      if (!attr.name.startsWith("xmlns")) endParaRPr.setAttribute(attr.name, attr.value);
    }
    for (const child of Array.from(source.childNodes))
      endParaRPr.appendChild(child.cloneNode(true));
  }
  paragraph.appendChild(endParaRPr);
  for (const old of paragraphs) old.remove();
  txBody.appendChild(paragraph);
}

const MERGE_ATTRS = ["gridSpan", "rowSpan", "hMerge", "vMerge"] as const;

function getIsMergeFlag(tc: Element, name: "hMerge" | "vMerge"): boolean {
  return parseOoxmlBool(tc.getAttribute(name) ?? undefined);
}

function blankCellFrom(source: Element): Element {
  const tc = source.cloneNode(true) as Element;
  for (const name of MERGE_ATTRS) tc.removeAttribute(name);
  clearCellText(tc);
  return tc;
}

function joinMerge(tc: Element, continuation: Element): void {
  for (const name of ["gridSpan", "hMerge", "vMerge"] as const) {
    const value = continuation.getAttribute(name);
    if (value !== null) tc.setAttribute(name, value);
  }
}

function checkIndex(index: number, max: number, what: string): void {
  if (!Number.isInteger(index) || index < 0 || index > max) {
    throw new Error(`applyEdit: ${what} index ${index} is out of bounds.`);
  }
}

function applySetTableGrid(pres: PresentationData, op: SetTableGridOperation): EditResult {
  return rewriteTable(pres, op.slideId, op.nodeId, (tbl) => {
    const tracks: [Element[], number[] | undefined, string][] = [
      [gridColumns(tbl), op.columnWidths, "w"],
      [tableRows(tbl), op.rowHeights, "h"],
    ];
    for (const [elements, sizes, attr] of tracks) {
      if (!sizes) continue;
      if (sizes.length !== elements.length || sizes.some((size) => !(size > 0))) {
        throw new Error(`applyEdit: table grid sizes don't match the table.`);
      }
      elements.forEach((element, index) =>
        element.setAttribute(attr, String(pxToEmu(sizes[index]))),
      );
    }
  });
}

/** The cell whose text and spans cover `cell`: itself, unless a merge hides it. */
export function getTableCellOrigin(table: TableNodeData, cell: TableCellAddress): TableCellAddress {
  let { row, column } = cell;
  while (column > 0 && table.rows[row]?.cells[column]?.hMerge) column--;
  while (row > 0 && table.rows[row]?.cells[column]?.vMerge) row--;
  return { row, column };
}

export interface TableCellRange {
  start: TableCellAddress;
  end: TableCellAddress;
}

/**
 * The rectangle between two cells, as top-left and bottom-right corners,
 * grown until it cuts through no merged cell.
 */
export function getTableCellRange(
  table: TableNodeData,
  from: TableCellAddress,
  to: TableCellAddress,
): TableCellRange {
  let top = Math.min(from.row, to.row);
  let bottom = Math.max(from.row, to.row);
  let left = Math.min(from.column, to.column);
  let right = Math.max(from.column, to.column);
  let grown = true;
  while (grown) {
    grown = false;
    for (let row = top; row <= bottom; row++) {
      for (let column = left; column <= right; column++) {
        const origin = getTableCellOrigin(table, { row, column });
        const cell = table.rows[origin.row]?.cells[origin.column];
        if (!cell) continue;
        const next = {
          top: Math.min(top, origin.row),
          left: Math.min(left, origin.column),
          bottom: Math.max(bottom, origin.row + cell.rowSpan - 1),
          right: Math.max(right, origin.column + cell.gridSpan - 1),
        };
        if (next.top < top || next.left < left || next.bottom > bottom || next.right > right) {
          ({ top, left, bottom, right } = next);
          grown = true;
        }
      }
    }
  }
  return { start: { row: top, column: left }, end: { row: bottom, column: right } };
}

function findTable(pres: PresentationData, slideId: string, nodeId: string): TableNodeData {
  const node = findNode(pres, findSlide(pres, slideId), nodeId);
  if (node.nodeType !== "table") throw new Error(`applyEdit: node "${nodeId}" is not a table.`);
  return node as TableNodeData;
}

const TCPR_LINES = ["lnL", "lnR", "lnT", "lnB"] as const;
const TCPR_FILLS = new Set(["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"]);
const TCPR_AFTER_FILL = new Set(["headers", "extLst"]);
const VERTICAL_ALIGN_TO_OOXML: Record<TableCellVerticalAlign, string> = {
  top: "t",
  middle: "ctr",
  bottom: "b",
};
const OPPOSITE_LINE = { lnL: "lnR", lnR: "lnL", lnT: "lnB", lnB: "lnT" } as const;

function ensureCellProperties(tc: Element): Element {
  const existing = childNamed(tc, "tcPr");
  if (existing) return existing;
  const tcPr = tc.ownerDocument.createElementNS(A_NS, "a:tcPr");
  tc.insertBefore(tcPr, firstChildIn(tc, new Set(["extLst"])));
  return tcPr;
}

function solidFillElement(doc: Document, color: string): Element {
  const solidFill = doc.createElementNS(A_NS, "a:solidFill");
  const srgb = doc.createElementNS(A_NS, "a:srgbClr");
  srgb.setAttribute("val", parseHexColor(color));
  solidFill.appendChild(srgb);
  return solidFill;
}

function setCellLine(
  tc: Element,
  name: (typeof TCPR_LINES)[number],
  borders: TableCellStyle["borders"],
): void {
  const tcPr = ensureCellProperties(tc);
  const doc = tc.ownerDocument;
  childNamed(tcPr, name)?.remove();
  const line = doc.createElementNS(A_NS, `a:${name}`);
  line.setAttribute("w", String(Math.round((borders?.width ?? 1) * 12700)));
  line.appendChild(
    borders ? solidFillElement(doc, borders.color) : doc.createElementNS(A_NS, "a:noFill"),
  );
  const later = TCPR_LINES.slice(TCPR_LINES.indexOf(name) + 1) as string[];
  const before = Array.from(tcPr.children).find(
    (child) => later.includes(child.localName) || !TCPR_LINES.includes(child.localName as never),
  );
  tcPr.insertBefore(line, before ?? null);
}

function applySetTableCellStyle(
  pres: PresentationData,
  op: SetTableCellStyleOperation,
): EditResult {
  const table = findTable(pres, op.slideId, op.nodeId);
  const { fill, borders, verticalAlign } = op.style;
  if (fill) parseHexColor(fill);
  if (borders) {
    parseHexColor(borders.color);
    if (borders.width !== undefined && !(borders.width > 0)) {
      throw new Error(`applyEdit: invalid border width "${borders.width}".`);
    }
  }
  const lastRow = table.rows.length - 1;
  const lastColumn = table.columns.length - 1;
  const { start, end } =
    op.from && op.to
      ? getTableCellRange(table, op.from, op.to)
      : { start: { row: 0, column: 0 }, end: { row: lastRow, column: lastColumn } };

  return rewriteTable(pres, op.slideId, op.nodeId, (tbl) => {
    const cells = tableRows(tbl).map(rowCells);
    const doc = tbl.ownerDocument;
    for (let row = start.row; row <= end.row; row++) {
      for (let column = start.column; column <= end.column; column++) {
        const tc = cells[row]?.[column];
        if (!tc) continue;
        const tcPr = ensureCellProperties(tc);
        if (fill !== undefined) {
          for (const child of Array.from(tcPr.children)) {
            if (TCPR_FILLS.has(child.localName)) child.remove();
          }
          tcPr.insertBefore(
            fill === null ? doc.createElementNS(A_NS, "a:noFill") : solidFillElement(doc, fill),
            firstChildIn(tcPr, TCPR_AFTER_FILL),
          );
        }
        if (verticalAlign) tcPr.setAttribute("anchor", VERTICAL_ALIGN_TO_OOXML[verticalAlign]);
        if (borders === undefined) continue;
        for (const name of TCPR_LINES) setCellLine(tc, name, borders);
        // A shared edge is drawn by one of its two cells, so the neighbor gets it too.
        const neighbors = {
          lnL: column === start.column ? cells[row]?.[column - 1] : undefined,
          lnR: column === end.column ? cells[row]?.[column + 1] : undefined,
          lnT: row === start.row ? cells[row - 1]?.[column] : undefined,
          lnB: row === end.row ? cells[row + 1]?.[column] : undefined,
        };
        for (const name of TCPR_LINES) {
          const neighbor = neighbors[name];
          if (neighbor) setCellLine(neighbor, OPPOSITE_LINE[name], borders);
        }
      }
    }
  });
}

const TABLE_OPTION_ATTRS: Record<keyof TableOptions, string> = {
  headerRow: "firstRow",
  totalRow: "lastRow",
  bandedRows: "bandRow",
  firstColumn: "firstCol",
  lastColumn: "lastCol",
  bandedColumns: "bandCol",
};

function applySetTableOptions(pres: PresentationData, op: SetTableOptionsOperation): EditResult {
  return rewriteTable(pres, op.slideId, op.nodeId, (tbl) => {
    let tblPr = childNamed(tbl, "tblPr");
    if (!tblPr) {
      tblPr = tbl.ownerDocument.createElementNS(A_NS, "a:tblPr");
      tbl.insertBefore(tblPr, tbl.firstElementChild);
    }
    for (const [key, attr] of Object.entries(TABLE_OPTION_ATTRS)) {
      const value = op.options[key as keyof TableOptions];
      if (value !== undefined) setOrRemoveAttr(tblPr, attr, value ? "1" : undefined);
    }
  });
}

function applyMergeTableCells(pres: PresentationData, op: MergeTableCellsOperation): EditResult {
  const table = findTable(pres, op.slideId, op.nodeId);
  for (const cell of [op.from, op.to]) {
    if (!table.rows[cell.row]?.cells[cell.column]) {
      throw new Error(`applyEdit: no cell at row ${cell.row}, column ${cell.column}.`);
    }
  }
  const { start, end } = getTableCellRange(table, op.from, op.to);
  if (start.row === end.row && start.column === end.column) {
    throw new Error("applyEdit: merging needs more than one cell.");
  }
  return rewriteTable(pres, op.slideId, op.nodeId, (tbl) => {
    const rows = tableRows(tbl);
    const origin = rowCells(rows[start.row])[start.column];
    if (!childNamed(origin, "txBody")) clearCellText(origin);
    const originText = childNamed(origin, "txBody")!;
    for (let row = start.row; row <= end.row; row++) {
      const cells = rowCells(rows[row]);
      for (let column = start.column; column <= end.column; column++) {
        const tc = cells[column];
        const isOrigin = row === start.row && column === start.column;
        const isHidden = getIsMergeFlag(tc, "hMerge") || getIsMergeFlag(tc, "vMerge");
        if (!isOrigin && !isHidden) {
          const text = childNamed(tc, "txBody");
          const paragraphs = text ? childParagraphs(text) : [];
          if (paragraphs.some((paragraph) => (paragraph.textContent ?? "").length > 0)) {
            for (const paragraph of paragraphs) originText.appendChild(paragraph);
          }
        }
        for (const name of MERGE_ATTRS) tc.removeAttribute(name);
        if (!isOrigin) {
          clearCellText(tc);
          if (column > start.column) tc.setAttribute("hMerge", "1");
          if (row > start.row) tc.setAttribute("vMerge", "1");
        }
      }
    }
    setSpan(origin, "gridSpan", end.column - start.column + 1);
    setSpan(origin, "rowSpan", end.row - start.row + 1);
    // An empty origin keeps one paragraph ahead of the text that joined it.
    const paragraphs = childParagraphs(originText);
    if (paragraphs.length > 1 && (paragraphs[0].textContent ?? "").length === 0) {
      paragraphs[0].remove();
    }
  });
}

function applySplitTableCell(pres: PresentationData, op: SplitTableCellOperation): EditResult {
  const table = findTable(pres, op.slideId, op.nodeId);
  const origin = getTableCellOrigin(table, op.cell);
  const cell = table.rows[origin.row]?.cells[origin.column];
  if (!cell || (cell.gridSpan === 1 && cell.rowSpan === 1)) {
    throw new Error(`applyEdit: no merged cell at row ${op.cell.row}, column ${op.cell.column}.`);
  }
  return rewriteTable(pres, op.slideId, op.nodeId, (tbl) => {
    const rows = tableRows(tbl);
    for (let row = origin.row; row < origin.row + cell.rowSpan; row++) {
      const cells = rowCells(rows[row]);
      for (let column = origin.column; column < origin.column + cell.gridSpan; column++) {
        const tc = cells[column];
        for (const name of MERGE_ATTRS) tc.removeAttribute(name);
        if (row !== origin.row || column !== origin.column) clearCellText(tc);
      }
    }
  });
}

function applyInsertTableRow(pres: PresentationData, op: InsertTableRowOperation): EditResult {
  return rewriteTable(pres, op.slideId, op.nodeId, (tbl) => {
    const rows = tableRows(tbl);
    checkIndex(op.index, rows.length, "row");
    const source = rows[op.index - 1] ?? rows[0];
    const below = rows[op.index];
    const tr = source.cloneNode(true) as Element;
    rowCells(tr).forEach((tc, column) => {
      const blank = blankCellFrom(tc);
      const displaced = below ? rowCells(below)[column] : undefined;
      if (displaced !== undefined && getIsMergeFlag(displaced, "vMerge")) {
        joinMerge(blank, displaced);
        const origin = rows
          .slice(0, op.index)
          .reverse()
          .map((row) => rowCells(row)[column])
          .find((cell) => cell && !getIsMergeFlag(cell, "vMerge"));
        if (origin?.hasAttribute("rowSpan"))
          setSpan(origin, "rowSpan", spanOf(origin, "rowSpan") + 1);
      } else {
        for (const name of ["gridSpan", "hMerge"] as const) {
          const value = tc.getAttribute(name);
          if (value !== null) blank.setAttribute(name, value);
        }
      }
      tr.replaceChild(blank, tc);
    });
    tbl.insertBefore(tr, below ?? null);
  });
}

function applyDeleteTableRow(pres: PresentationData, op: DeleteTableRowOperation): EditResult {
  return rewriteTable(pres, op.slideId, op.nodeId, (tbl) => {
    const rows = tableRows(tbl);
    checkIndex(op.index, rows.length - 1, "row");
    if (rows.length === 1) throw new Error("applyEdit: cannot delete a table's last row.");
    const tr = rows[op.index];
    rowCells(tr).forEach((tc, column) => {
      const rowSpan = spanOf(tc, "rowSpan");
      if (rowSpan > 1) {
        const next = rows[op.index + 1] ? rowCells(rows[op.index + 1])[column] : undefined;
        if (!next) return;
        const text = childNamed(tc, "txBody");
        const nextText = childNamed(next, "txBody");
        if (text && nextText) next.replaceChild(text, nextText);
        next.removeAttribute("vMerge");
        setSpan(next, "rowSpan", rowSpan - 1);
      } else if (getIsMergeFlag(tc, "vMerge")) {
        const origin = rows
          .slice(0, op.index)
          .reverse()
          .map((row) => rowCells(row)[column])
          .find((cell) => cell && !getIsMergeFlag(cell, "vMerge"));
        if (origin?.hasAttribute("rowSpan"))
          setSpan(origin, "rowSpan", spanOf(origin, "rowSpan") - 1);
      }
    });
    tr.remove();
  });
}

function applyInsertTableColumn(
  pres: PresentationData,
  op: InsertTableColumnOperation,
): EditResult {
  return rewriteTable(pres, op.slideId, op.nodeId, (tbl) => {
    const columns = gridColumns(tbl);
    checkIndex(op.index, columns.length, "column");
    const sourceColumn = columns[op.index - 1] ?? columns[0];
    sourceColumn.parentElement!.insertBefore(
      sourceColumn.cloneNode(true),
      columns[op.index] ?? null,
    );
    for (const tr of tableRows(tbl)) {
      const cells = rowCells(tr);
      const blank = blankCellFrom(cells[op.index - 1] ?? cells[0]);
      const displaced = cells[op.index];
      if (displaced !== undefined && getIsMergeFlag(displaced, "hMerge")) {
        joinMerge(blank, displaced);
        blank.removeAttribute("gridSpan");
        const origin = cells
          .slice(0, op.index)
          .reverse()
          .find((cell) => !getIsMergeFlag(cell, "hMerge"));
        if (origin?.hasAttribute("gridSpan"))
          setSpan(origin, "gridSpan", spanOf(origin, "gridSpan") + 1);
      }
      tr.insertBefore(blank, displaced ?? firstChildIn(tr, new Set(["extLst"])));
    }
  });
}

function applyDeleteTableColumn(
  pres: PresentationData,
  op: DeleteTableColumnOperation,
): EditResult {
  return rewriteTable(pres, op.slideId, op.nodeId, (tbl) => {
    const columns = gridColumns(tbl);
    checkIndex(op.index, columns.length - 1, "column");
    if (columns.length === 1) throw new Error("applyEdit: cannot delete a table's last column.");
    for (const tr of tableRows(tbl)) {
      const cells = rowCells(tr);
      const tc = cells[op.index];
      if (!tc) continue;
      const gridSpan = spanOf(tc, "gridSpan");
      if (gridSpan > 1) {
        const next = cells[op.index + 1];
        if (next) {
          const text = childNamed(tc, "txBody");
          const nextText = childNamed(next, "txBody");
          if (text && nextText) next.replaceChild(text, nextText);
          next.removeAttribute("hMerge");
          setSpan(next, "gridSpan", gridSpan - 1);
          const rowSpan = tc.getAttribute("rowSpan");
          if (rowSpan !== null) next.setAttribute("rowSpan", rowSpan);
        }
      } else if (getIsMergeFlag(tc, "hMerge")) {
        const origin = cells
          .slice(0, op.index)
          .reverse()
          .find((cell) => !getIsMergeFlag(cell, "hMerge"));
        if (origin?.hasAttribute("gridSpan"))
          setSpan(origin, "gridSpan", spanOf(origin, "gridSpan") - 1);
      }
      tc.remove();
    }
    columns[op.index].remove();
  });
}

/** spPr children that come after `a:ln` in the schema sequence. */
const AFTER_LINE_LOCAL_NAMES = new Set(["effectLst", "effectDag", "scene3d", "sp3d", "extLst"]);

function applySetOutline(pres: PresentationData, op: SetOutlineOperation): EditResult {
  const sourcePackage = requirePkg(pres);
  const slide = findSlide(pres, op.slideId);
  const node = findNode(pres, slide, op.nodeId);
  if (node.nodeType !== "shape" && node.nodeType !== "picture") {
    throw new Error(
      `applyEdit: setOutline only supports shapes and pictures, got "${node.nodeType}".`,
    );
  }
  if (op.width !== undefined && !(op.width >= 0)) {
    throw new Error(`applyEdit: invalid outline width "${op.width}".`);
  }
  const color = op.color ? parseHexColor(op.color) : op.color;

  const spPr = requireElement(node.source.child("spPr"), `spPr of node "${node.id}"`);
  const original = childNamed(spPr, "ln");
  const ref = original?.nextElementSibling ?? firstChildIn(spPr, AFTER_LINE_LOCAL_NAMES);
  const line = original
    ? (original.cloneNode(true) as Element)
    : spPr.ownerDocument.createElementNS(A_NS, "a:ln");

  if (op.width !== undefined) line.setAttribute("w", String(Math.round(op.width * 12700)));
  if (color !== undefined) {
    // The fill is the first child of a:ln.
    for (const child of Array.from(line.children)) {
      if (RPR_FILL_LOCAL_NAMES.has(child.localName)) child.remove();
    }
    const fill = spPr.ownerDocument.createElementNS(
      A_NS,
      color === null ? "a:noFill" : "a:solidFill",
    );
    if (color !== null) {
      const srgb = spPr.ownerDocument.createElementNS(A_NS, "a:srgbClr");
      srgb.setAttribute("val", color);
      fill.appendChild(srgb);
    }
    line.insertBefore(fill, line.firstElementChild);
  }

  if (original) removeChild(original);
  insertChild(spPr, line, ref);
  const lined = node as ShapeNodeData | PictureNodeData;
  const previousLine = lined.line;
  lined.line = new SafeXmlNode(line);
  sourcePackage.markDirty(slide.id);

  return {
    affectedSlideIds: [slide.id],
    undo: () => {
      const restoreRef = line.nextElementSibling;
      removeChild(line);
      if (original) insertChild(spPr, original, restoreRef);
      lined.line = previousLine;
      sourcePackage.markDirty(slide.id);
    },
  };
}

async function applyAddImage(pres: PresentationData, op: AddImageOperation): Promise<EditResult> {
  const sourcePackage = requirePkg(pres);
  const extension = IMAGE_EXTENSIONS[op.contentType];
  if (!extension) {
    throw new Error(`applyEdit: unsupported image type "${op.contentType}".`);
  }
  const slide = findSlide(pres, op.slideId);
  materializeSlide(pres, slide);
  const spTree = requireSpTree(pres, slide);
  const doc = spTree.ownerDocument;

  let size = op.size;
  if (!size) {
    const natural = readImageSize(op.data) ?? { w: pres.width / 2, h: pres.height / 2 };
    const scale = Math.min(1, (pres.width * 0.8) / natural.w, (pres.height * 0.8) / natural.h);
    size = { w: natural.w * scale, h: natural.h * scale };
  }
  const position = op.position ?? {
    x: (pres.width - size.w) / 2,
    y: (pres.height - size.h) / 2,
  };

  const media = await addMediaPart(pres, extension, op.contentType, op.data);
  const rels = await addSlideRels(pres, slide, [{ type: IMAGE_REL_TYPE, target: media.target }]);
  const [rId] = rels.ids;

  const id = nextShapeId(spTree);
  const name = op.name ?? `Picture ${id - 1}`;
  const el = (qualifiedName: string, attrs: Record<string, string> = {}, ...children: Node[]) => {
    const element = doc.createElementNS(
      qualifiedName.startsWith("a:") ? A_NS : P_NS,
      qualifiedName,
    );
    for (const [attr, value] of Object.entries(attrs)) element.setAttribute(attr, value);
    for (const child of children) element.appendChild(child);
    return element;
  };
  const blip = el("a:blip");
  blip.setAttributeNS(R_NS, "r:embed", rId);
  const pic = el(
    "p:pic",
    {},
    el(
      "p:nvPicPr",
      {},
      el("p:cNvPr", { id: String(id), name, descr: name }),
      el("p:cNvPicPr", {}, el("a:picLocks", { noChangeAspect: "1" })),
      el("p:nvPr"),
    ),
    el("p:blipFill", {}, blip, el("a:stretch", {}, el("a:fillRect"))),
    el(
      "p:spPr",
      {},
      el(
        "a:xfrm",
        {},
        el("a:off", { x: String(pxToEmu(position.x)), y: String(pxToEmu(position.y)) }),
        el("a:ext", { cx: String(pxToEmu(size.w)), cy: String(pxToEmu(size.h)) }),
      ),
      el("a:prstGeom", { prst: "rect" }, el("a:avLst")),
    ),
  );
  // extLst, when present, has to stay the last child of spTree.
  const extLst = Array.from(spTree.children).find((child) => child.localName === "extLst");
  insertChild(spTree, pic, extLst ?? null);

  const picNode = parseAddedNode(pres, slide, pic);
  slide.nodes.push(picNode);
  sourcePackage.markDirty(slide.id);

  return {
    affectedSlideIds: [slide.id],
    createdNodeIds: [picNode.id],
    undo: () => {
      removeChild(pic);
      slide.nodes.splice(slide.nodes.indexOf(picNode), 1);
      rels.undo();
      media.undo();
      sourcePackage.markDirty(slide.id);
    },
  };
}

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
    throw new Error("applyEdit: presentation.xml is not registered on the package.");
  }
  const lstNode = presRoot.child("sldIdLst");
  const lstElement = lstNode.element;
  if (!lstElement) {
    throw new Error("applyEdit: unsupported package: presentation.xml has no sldIdLst.");
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
  throw new Error(`applyEdit: slide "${slideId}" not found in presentation.xml sldIdLst.`);
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

  const slideText = await sourcePackage.readText(source.id);
  if (slideText === undefined) {
    throw new Error(`applyEdit: slide part "${source.id}" is missing from the package.`);
  }
  const slideRelsText = await sourcePackage.readText(relsPathFor(source.id));

  return insertSlidePart(pres, ctx, {
    slideText,
    slideRelsText,
    contentTypeFrom: source.id,
    afterSldIdElement: sourceSldIdElement,
    index: sourceIndex + 1,
  });
}

interface InsertSlidePartOptions {
  slideText: string;
  slideRelsText: string | undefined;
  /** Existing slide part whose content type override to copy. */
  contentTypeFrom?: string;
  /** Insert the `p:sldId` after this entry; `null` appends. */
  afterSldIdElement: Element | null;
  /** Position of the new slide in the deck. */
  index: number;
}

/**
 * Registers a new slide part on the package (part, rels, content type,
 * presentation relationship, and `p:sldId`) and parses it into the model.
 */
async function insertSlidePart(
  pres: PresentationData,
  ctx: SlideListContext,
  options: InsertSlidePartOptions,
): Promise<EditResult> {
  const { sourcePackage } = ctx;
  const { slideText, slideRelsText, contentTypeFrom, afterSldIdElement, index } = options;

  const newPath = nextSlidePartPath(sourcePackage);
  const newRelsPath = relsPathFor(newPath);
  const ctText = (await sourcePackage.readText(CONTENT_TYPES_PATH)) ?? "";

  sourcePackage.setEntry(newPath, slideText);
  if (slideRelsText !== undefined) {
    sourcePackage.setEntry(newRelsPath, slideRelsText);
  }

  const ctRoot = parseXml(ctText);
  const ctRootElement = requireElement(ctRoot, "[Content_Types].xml root");
  const sourceOverride = contentTypeFrom
    ? ctRoot.children("Override").find((o) => o.attr("PartName") === `/${contentTypeFrom}`)
    : undefined;
  const overrideElement = ctRootElement.ownerDocument.createElementNS(CT_NS, "Override");
  overrideElement.setAttribute("PartName", `/${newPath}`);
  overrideElement.setAttribute(
    "ContentType",
    sourceOverride?.attr("ContentType") ?? SLIDE_CONTENT_TYPE,
  );
  insertChild(ctRootElement, overrideElement, null);
  sourcePackage.setEntry(CONTENT_TYPES_PATH, serializePartText(ctRootElement, ctText));

  const newRId = nextRelId(ctx.rels);
  const relsRoot = parseXml(ctx.relsText);
  const relsRootElement = requireElement(relsRoot, "presentation rels root");
  const relElement = relsRootElement.ownerDocument.createElementNS(RELS_NS, "Relationship");
  relElement.setAttribute("Id", newRId);
  relElement.setAttribute("Type", SLIDE_REL_TYPE);
  relElement.setAttribute("Target", newPath.replace(/^ppt\//, ""));
  insertChild(relsRootElement, relElement, null);
  sourcePackage.setEntry(PRESENTATION_RELS_PATH, serializePartText(relsRootElement, ctx.relsText));

  const sldIdElement = ctx.lstElement.ownerDocument.createElementNS(P_NS, "p:sldId");
  sldIdElement.setAttribute("id", String(nextSldIdNumber(ctx)));
  sldIdElement.setAttributeNS(R_NS, "r:id", newRId);
  insertChild(ctx.lstElement, sldIdElement, afterSldIdElement?.nextElementSibling ?? null);
  sourcePackage.markDirty(PRESENTATION_PATH);

  const newRoot = parseXml(slideText);
  sourcePackage.registerXmlRoot(newPath, newRoot);
  const newRels = parseRels(slideRelsText ?? "");
  const newSlide = parseSlide(newRoot, index, newRels, newPath, pres.diagramDrawings);
  if (newSlide.layoutIndex) {
    newSlide.layoutIndex = resolveRelTarget("ppt/slides", newSlide.layoutIndex);
  }
  pres.slides.splice(index, 0, newSlide);
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

const SLIDE_LAYOUT_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
/** Layout placeholders PowerPoint leaves off a new slide (header/footer furniture). */
const SKIPPED_PLACEHOLDER_TYPES = new Set(["dt", "ftr", "sldNum", "hdr"]);

const EMPTY_SLIDE_XML =
  `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">` +
  "<p:cSld><p:spTree>" +
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
  "</p:spTree></p:cSld>" +
  "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>" +
  "</p:sld>";

/**
 * Builds slide XML holding an empty copy of each content placeholder on the
 * layout. The copies carry only the `p:ph` reference, so position, size and
 * text styles keep inheriting from the layout.
 */
function buildSlideFromLayout(layoutSpTree: SafeXmlNode): string {
  const root = parseXml(EMPTY_SLIDE_XML);
  const rootElement = requireElement(root, "slide root");
  const doc = rootElement.ownerDocument;
  const spTree = requireElement(root.child("cSld").child("spTree"), "spTree");

  const el = (qualifiedName: string, attrs: Record<string, string> = {}, ...children: Node[]) => {
    const element = doc.createElementNS(
      qualifiedName.startsWith("a:") ? A_NS : P_NS,
      qualifiedName,
    );
    for (const [name, value] of Object.entries(attrs)) element.setAttribute(name, value);
    for (const child of children) element.appendChild(child);
    return element;
  };

  let nextId = 2;
  for (const sp of layoutSpTree.children("sp")) {
    const nvSpPr = sp.child("nvSpPr");
    const ph = nvSpPr.child("nvPr").child("ph");
    const phElement = ph.element;
    if (!phElement || SKIPPED_PLACEHOLDER_TYPES.has(ph.attr("type") ?? "")) continue;

    const phCopy = doc.importNode(phElement, true) as Element;
    phCopy.removeAttribute("hasCustomPrompt");
    const id = nextId++;
    const name = nvSpPr.child("cNvPr").attr("name") ?? `Placeholder ${id - 1}`;

    spTree.appendChild(
      el(
        "p:sp",
        {},
        el(
          "p:nvSpPr",
          {},
          el("p:cNvPr", { id: String(id), name }),
          el("p:cNvSpPr", {}, el("a:spLocks", { noGrp: "1" })),
          el("p:nvPr", {}, phCopy),
        ),
        el("p:spPr"),
        el(
          "p:txBody",
          {},
          el("a:bodyPr"),
          el("a:lstStyle"),
          el("a:p", {}, el("a:endParaRPr", { lang: "en-US" })),
        ),
      ),
    );
  }

  return serializePartText(rootElement, "");
}

async function applyAddSlide(pres: PresentationData, op: AddSlideOperation): Promise<EditResult> {
  const ctx = await loadSlideListContext(pres);
  const anchor = op.afterSlideId
    ? findSlide(pres, op.afterSlideId)
    : pres.slides[pres.slides.length - 1];

  const layoutId = op.layoutId ?? anchor?.layoutIndex ?? pres.layouts.keys().next().value;
  const layout = layoutId ? pres.layouts.get(layoutId) : undefined;
  if (!layoutId || !layout) {
    throw new Error(`applyEdit: unknown slide layout "${layoutId ?? ""}".`);
  }

  // Slides live in ppt/slides/, so the layout target is relative to it.
  const layoutTarget = `../${layoutId.replace(/^ppt\//, "")}`;
  const slideRelsText =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    `<Relationships xmlns="${RELS_NS}">` +
    `<Relationship Id="rId1" Type="${SLIDE_LAYOUT_REL_TYPE}" Target="${layoutTarget}"/>` +
    "</Relationships>";

  return insertSlidePart(pres, ctx, {
    slideText: buildSlideFromLayout(layout.spTree),
    slideRelsText,
    contentTypeFrom: anchor?.id,
    afterSldIdElement: anchor
      ? findSldIdEntry(ctx, anchor.id).element
      : (sldIdElements(ctx).at(-1) ?? null),
    index: anchor ? pres.slides.indexOf(anchor) + 1 : pres.slides.length,
  });
}

function applySetSlideBackground(
  pres: PresentationData,
  op: SetSlideBackgroundOperation,
): EditResult {
  const sourcePackage = requirePkg(pres);
  const slide = findSlide(pres, op.slideId);
  materializeSlide(pres, slide);
  const color = op.color === null ? null : parseHexColor(op.color);

  const cSld = sourcePackage.getXmlRoot(slide.id)?.child("cSld");
  const cSldElement = cSld ? requireElement(cSld, "cSld") : null;
  if (!cSldElement) throw new Error(`applyEdit: slide part "${slide.id}" is not loaded.`);

  const prevBackground = slide.background;
  const oldBg = childNamed(cSldElement, "bg");
  const ref = oldBg?.nextElementSibling ?? cSldElement.firstElementChild;

  let newBg: Element | null = null;
  if (color !== null) {
    const doc = cSldElement.ownerDocument;
    newBg = doc.createElementNS(P_NS, "p:bg");
    const bgPr = doc.createElementNS(P_NS, "p:bgPr");
    const solidFill = doc.createElementNS(A_NS, "a:solidFill");
    const srgb = doc.createElementNS(A_NS, "a:srgbClr");
    srgb.setAttribute("val", color);
    solidFill.appendChild(srgb);
    bgPr.appendChild(solidFill);
    bgPr.appendChild(doc.createElementNS(A_NS, "a:effectLst"));
    newBg.appendChild(bgPr);
  }

  if (oldBg) removeChild(oldBg);
  // p:bg is the first child of p:cSld.
  if (newBg) insertChild(cSldElement, newBg, ref);
  slide.background = newBg ? new SafeXmlNode(newBg) : undefined;
  sourcePackage.markDirty(slide.id);

  return {
    affectedSlideIds: [slide.id],
    undo: () => {
      const restoreRef = newBg?.nextElementSibling ?? cSldElement.firstElementChild;
      if (newBg) removeChild(newBg);
      if (oldBg) insertChild(cSldElement, oldBg, restoreRef);
      slide.background = prevBackground;
      sourcePackage.markDirty(slide.id);
    },
  };
}

async function applyDeleteSlide(
  pres: PresentationData,
  op: DeleteSlideOperation,
): Promise<EditResult> {
  if (pres.slides.length <= 1) {
    throw new Error("applyEdit: cannot delete the last slide in the deck.");
  }

  const ctx = await loadSlideListContext(pres);
  const { sourcePackage } = ctx;
  const slide = findSlide(pres, op.slideId);
  const slideIndex = pres.slides.indexOf(slide);
  const { element: sldIdElement, rId } = findSldIdEntry(ctx, op.slideId);
  const slideRelsPath = relsPathFor(slide.id);

  const slideBytes = await sourcePackage.readBytes(slide.id);
  const slideRelsBytes = await sourcePackage.readBytes(slideRelsPath);
  const slideXmlRoot = sourcePackage.getXmlRoot(slide.id);
  const ctText = (await sourcePackage.readText(CONTENT_TYPES_PATH)) ?? "";

  const sldIdRef = sldIdElement.nextElementSibling;
  removeChild(sldIdElement);
  sourcePackage.markDirty(PRESENTATION_PATH);

  const relsRoot = parseXml(ctx.relsText);
  const relsRootElement = requireElement(relsRoot, "presentation rels root");
  const relNode = relsRoot.children("Relationship").find((r) => r.attr("Id") === rId);
  if (relNode?.element) {
    removeChild(relNode.element);
  }
  sourcePackage.setEntry(PRESENTATION_RELS_PATH, serializePartText(relsRootElement, ctx.relsText));

  const ctRoot = parseXml(ctText);
  const ctRootElement = requireElement(ctRoot, "[Content_Types].xml root");
  const overrideNode = ctRoot
    .children("Override")
    .find((o) => o.attr("PartName") === `/${slide.id}`);
  if (overrideNode?.element) {
    removeChild(overrideNode.element);
    sourcePackage.setEntry(CONTENT_TYPES_PATH, serializePartText(ctRootElement, ctText));
  }

  // Media parts stay: they may be shared, and orphans are harmless.
  sourcePackage.deleteEntry(slide.id);
  sourcePackage.deleteEntry(slideRelsPath);

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
