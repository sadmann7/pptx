import type {
  ParagraphAlign,
  SetTextBodyParagraph,
  ShapeNodeData,
  SlideNode,
  TableCellAddress,
  TableNodeData,
  TextBody,
  TextPosition,
  TextRange,
} from "@diceui/pptx-core";
import { PPTX_ATTRS, PPTX_DATASET } from "@diceui/pptx-core";

/** Formatting of the current text, as a toolbar shows it. */
export interface ActiveTextStyle {
  /** `true` when every character in the selection is bold. */
  bold: boolean;
  /** `true` when every character in the selection is italic. */
  italic: boolean;
  /** `true` when every character in the selection is underlined. */
  underline: boolean;
  /** `true` when every character in the selection is struck through. */
  strikethrough: boolean;
  /** Hex color (`"#RRGGBB"`), or `null` when the selection mixes colors. */
  color: string | null;
  /** Font family, or `null` when the selection mixes fonts. */
  fontFamily: string | null;
  /** Font size in points as rendered, or `null` when the selection mixes sizes. */
  fontSize: number | null;
  /** Paragraph alignment, or `null` when the selected paragraphs differ. */
  align: ParagraphAlign | null;
}

// The renderer's zero-width line-height spacers must not reach the model.
export function cleanText(raw: string | null | undefined): string {
  return (raw ?? "").replace(/​/g, "");
}

// A container whose paragraphs were flattened by editing counts as one paragraph.
export function getParagraphElements(container: HTMLElement): HTMLElement[] {
  const children = Array.from(container.children).filter(
    (element): element is HTMLElement => element instanceof HTMLElement,
  );
  const paragraphDivs = children.filter(
    (element) => element.dataset[PPTX_DATASET.run] === undefined,
  );
  return paragraphDivs.length > 0 && paragraphDivs.length === children.length
    ? paragraphDivs
    : [container];
}

// The `<br>` a browser adds to keep an empty line open has nothing after it, unlike an `a:br` one.
function isRenderedBreak(element: HTMLBRElement): boolean {
  return (
    element.nextSibling !== null && !element.hasAttribute(PPTX_ATTRS.emptyLine)
  );
}

function getTextLeaves(paragraph: HTMLElement): (Text | HTMLBRElement)[] {
  const leaves: (Text | HTMLBRElement)[] = [];
  const walker = document.createTreeWalker(
    paragraph,
    NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT,
    {
      acceptNode: (node) =>
        node instanceof HTMLElement &&
        node.dataset[PPTX_DATASET.bullet] !== undefined
          ? NodeFilter.FILTER_REJECT
          : NodeFilter.FILTER_ACCEPT,
    },
  );
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (node.nodeType === Node.TEXT_NODE) {
      leaves.push(node as Text);
    } else if (node instanceof HTMLBRElement && isRenderedBreak(node)) {
      leaves.push(node);
    }
  }
  return leaves;
}

function getLeafLength(leaf: Text | HTMLBRElement): number {
  return leaf instanceof HTMLBRElement ? 1 : cleanText(leaf.data).length;
}

export function getParagraphText(paragraph: HTMLElement): string {
  return getTextLeaves(paragraph)
    .map((leaf) =>
      leaf instanceof HTMLBRElement ? "\n" : cleanText(leaf.data),
    )
    .join("");
}

export function getTextPosition(
  container: HTMLElement,
  node: Node,
  offset: number,
): TextPosition | null {
  const paragraphs = getParagraphElements(container);

  // A point between paragraphs (on the container itself) snaps to the start
  // of the paragraph after it, or the end of the last one.
  if (node === container && paragraphs[0] !== container) {
    if (offset < paragraphs.length)
      return { paragraphIndex: offset, offset: 0 };
    const last = paragraphs.length - 1;
    return {
      paragraphIndex: last,
      offset: getParagraphText(paragraphs[last]).length,
    };
  }

  const paragraphIndex = paragraphs.findIndex((paragraph) =>
    paragraph.contains(node),
  );
  if (paragraphIndex < 0) return null;
  const paragraph = paragraphs[paragraphIndex];

  const point = document.createRange();
  point.setStart(node, offset);

  let count = 0;
  for (const leaf of getTextLeaves(paragraph)) {
    if (leaf === node) {
      count += cleanText((leaf as Text).data.slice(0, offset)).length;
      break;
    }
    if (point.comparePoint(leaf, 0) >= 0) break;
    count += getLeafLength(leaf);
  }
  return { paragraphIndex, offset: count };
}

export function getTextRange(
  container: HTMLElement,
  range: AbstractRange,
): TextRange | null {
  const start = getTextPosition(
    container,
    range.startContainer,
    range.startOffset,
  );
  const end = getTextPosition(container, range.endContainer, range.endOffset);
  return start && end ? { start, end } : null;
}

export function getDomPoint(
  container: HTMLElement,
  position: TextPosition,
): [node: Node, offset: number] | null {
  const paragraph = getParagraphElements(container)[position.paragraphIndex];
  if (!paragraph) return null;

  let remaining = position.offset;
  for (const leaf of getTextLeaves(paragraph)) {
    const length = getLeafLength(leaf);
    if (remaining <= length && leaf instanceof Text) return [leaf, remaining];
    if (remaining === 0) {
      return [
        leaf.parentNode!,
        Array.prototype.indexOf.call(leaf.parentNode!.childNodes, leaf),
      ];
    }
    remaining -= length;
  }
  return [paragraph, paragraph.childNodes.length];
}

const WORD_CHARACTER = /[\p{L}\p{N}_'’-]/u;

// PowerPoint formats the word under a collapsed caret.
export function expandToWord(
  text: string,
  offset: number,
): [start: number, end: number] | null {
  let start = offset;
  let end = offset;
  while (start > 0 && WORD_CHARACTER.test(text[start - 1])) start--;
  while (end < text.length && WORD_CHARACTER.test(text[end])) end++;
  return start === end ? null : [start, end];
}

function toHexColor(cssColor: string): string | null {
  const match = /rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)/.exec(cssColor);
  if (!match) return null;
  return `#${match
    .slice(1, 4)
    .map((channel) => Number(channel).toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase()}`;
}

// The inline `pt` size, which CSS `zoom` leaves alone, divided by the container's `normAutofit` scale.
function getFontSize(element: HTMLElement): number | null {
  const scale = Number(
    element.closest<HTMLElement>(`[${PPTX_ATTRS.fontScale}]`)?.dataset[
      PPTX_DATASET.fontScale
    ] ?? 1,
  );
  for (
    let current: HTMLElement | null = element;
    current;
    current = current.parentElement
  ) {
    const size = current.style.fontSize;
    if (size.endsWith("pt"))
      return Math.round((Number.parseFloat(size) / scale) * 10) / 10;
    if (current.dataset[PPTX_DATASET.nodeId] !== undefined) break;
  }
  return null;
}

function toAlign(textAlign: string): ParagraphAlign {
  if (textAlign === "center") return "center";
  if (textAlign === "right" || textAlign === "end") return "right";
  if (textAlign === "justify") return "justify";
  return "left";
}

function getShared<T>(values: T[]): T | null {
  return values.every((value) => value === values[0])
    ? (values[0] ?? null)
    : null;
}

export function getActiveTextStyle(
  runs: HTMLElement[],
  paragraphs: HTMLElement[],
): ActiveTextStyle | null {
  if (runs.length === 0) return null;
  const styles = runs.map((run) => {
    const computed = getComputedStyle(run);
    const decoration = `${computed.textDecorationLine} ${run.style.textDecoration}`;
    return {
      bold:
        Number.parseInt(computed.fontWeight, 10) >= 600 ||
        computed.fontWeight === "bold",
      italic:
        computed.fontStyle === "italic" || computed.fontStyle === "oblique",
      underline: decoration.includes("underline"),
      strikethrough: decoration.includes("line-through"),
      color: toHexColor(computed.color),
      fontFamily:
        computed.fontFamily
          .split(",")[0]
          ?.trim()
          .replace(/^["']|["']$/g, "") || null,
      fontSize: getFontSize(run),
    };
  });

  return {
    bold: styles.every((style) => style.bold),
    italic: styles.every((style) => style.italic),
    underline: styles.every((style) => style.underline),
    strikethrough: styles.every((style) => style.strikethrough),
    color: getShared(styles.map((style) => style.color)),
    fontFamily: getShared(styles.map((style) => style.fontFamily)),
    fontSize: getShared(styles.map((style) => style.fontSize)),
    align: getShared(
      paragraphs.map((paragraph) =>
        toAlign(getComputedStyle(paragraph).textAlign),
      ),
    ),
  };
}

function getSelectedText(range: Range, node: Node): string {
  const overlap = document.createRange();
  overlap.selectNodeContents(node);
  if (overlap.compareBoundaryPoints(Range.START_TO_START, range) < 0) {
    overlap.setStart(range.startContainer, range.startOffset);
  }
  if (overlap.compareBoundaryPoints(Range.END_TO_END, range) > 0) {
    overlap.setEnd(range.endContainer, range.endOffset);
  }
  return cleanText(overlap.toString());
}

// An empty paragraph returns itself, since it carries the paragraph mark's size.
export function getRangeRuns(
  container: HTMLElement,
  range: AbstractRange,
): { runs: HTMLElement[]; paragraphs: HTMLElement[] } {
  const live = document.createRange();
  live.setStart(range.startContainer, range.startOffset);
  live.setEnd(range.endContainer, range.endOffset);

  const paragraphs = getParagraphElements(container).filter((paragraph) =>
    live.intersectsNode(paragraph),
  );
  const allRuns = Array.from(
    container.querySelectorAll<HTMLElement>(`[${PPTX_ATTRS.run}]`),
  );

  if (!live.collapsed) {
    const runs = allRuns.filter(
      (run) =>
        live.intersectsNode(run) && getSelectedText(live, run).length > 0,
    );
    if (runs.length > 0) return { runs, paragraphs };
  }

  const start = range.startContainer;
  const startElement =
    start instanceof HTMLElement ? start : start.parentElement;
  const run = startElement?.closest<HTMLElement>(`[${PPTX_ATTRS.run}]`);
  if (run && container.contains(run)) return { runs: [run], paragraphs };
  const paragraph = paragraphs[0] ?? getParagraphElements(container)[0];
  if (!paragraph) return { runs: [], paragraphs };
  const firstRun = paragraph.querySelector<HTMLElement>(`[${PPTX_ATTRS.run}]`);
  return { runs: [firstRun ?? paragraph], paragraphs: [paragraph] };
}

export function textContainerOf(shapeElement: HTMLElement): HTMLElement | null {
  // An empty placeholder's prompt overlay also carries the paragraph attribute but is not editable.
  for (const paragraphElement of Array.from(
    shapeElement.querySelectorAll<HTMLElement>(`[${PPTX_ATTRS.paragraph}]`),
  )) {
    if (paragraphElement.closest(`[${PPTX_ATTRS.placeholderPrompt}]`)) continue;
    return paragraphElement.parentElement;
  }
  return null;
}

export function placeholderPromptOf(
  shapeElement: HTMLElement,
): HTMLElement | null {
  return shapeElement.querySelector<HTMLElement>(
    `[${PPTX_ATTRS.placeholderPrompt}]`,
  );
}

export function readBackTextBody(
  container: HTMLElement,
): SetTextBodyParagraph[] {
  const paragraphs: SetTextBodyParagraph[] = [];
  let lastSourceParagraph = 0;

  for (const paragraphDiv of getParagraphElements(container)) {
    const paragraphIndexAttr = paragraphDiv.dataset?.[PPTX_DATASET.paragraph];
    const sourceParagraphIndex =
      paragraphIndexAttr !== undefined
        ? Number(paragraphIndexAttr)
        : lastSourceParagraph;
    lastSourceParagraph = sourceParagraphIndex;

    const runs = readRunsFromParagraphElement(
      paragraphDiv,
      sourceParagraphIndex,
    );
    paragraphs.push({ sourceParagraphIndex, runs });
  }

  return paragraphs;
}

function readRunsFromParagraphElement(
  paragraphDiv: HTMLElement,
  defaultSourceParagraph: number,
): SetTextBodyParagraph["runs"] {
  const runs: SetTextBodyParagraph["runs"] = [];
  let lastSourceRun: [number, number] | undefined;

  for (const child of Array.from(paragraphDiv.childNodes)) {
    if (
      child instanceof HTMLElement &&
      child.dataset[PPTX_DATASET.bullet] !== undefined
    ) {
      continue;
    }

    if (
      child instanceof HTMLElement &&
      child.dataset[PPTX_DATASET.run] !== undefined
    ) {
      const runIndex = Number(child.dataset[PPTX_DATASET.run]);
      const sourceRun: [number, number] = [defaultSourceParagraph, runIndex];
      const text = cleanText(child.textContent);
      if (text.length > 0) {
        runs.push({ text, sourceRun });
        lastSourceRun = sourceRun;
      }
    } else if (child instanceof HTMLBRElement) {
      if (isRenderedBreak(child))
        runs.push({ text: "\n", sourceRun: lastSourceRun });
    } else if (child.nodeType === Node.TEXT_NODE) {
      const text = cleanText(child.textContent);
      if (text.length > 0) {
        runs.push({ text, sourceRun: lastSourceRun });
      }
    } else if (child instanceof HTMLElement) {
      const text = cleanText(child.textContent);
      if (text.length > 0) {
        runs.push({ text, sourceRun: lastSourceRun });
      }
    }
  }

  if (runs.length === 0) {
    runs.push({ text: "", sourceRun: lastSourceRun });
  }

  return runs;
}

export function getTableCells(node: SlideNode): TableCellAddress[] {
  if (node.nodeType !== "table") return [];
  const cells: TableCellAddress[] = [];
  for (const [row, { cells: rowCells }] of (
    node as TableNodeData
  ).rows.entries()) {
    for (const [column, cell] of rowCells.entries()) {
      if (!cell.hMerge && !cell.vMerge && cell.textBody)
        cells.push({ row, column });
    }
  }
  return cells;
}

export function getTextBody(
  node: SlideNode,
  cell?: TableCellAddress,
): TextBody | undefined {
  if (node.nodeType === "table") {
    return cell
      ? (node as TableNodeData).rows[cell.row]?.cells[cell.column]?.textBody
      : undefined;
  }
  return node.nodeType === "shape"
    ? (node as ShapeNodeData).textBody
    : undefined;
}

export function getCellAddress(element: Element): TableCellAddress | undefined {
  const value = element
    .closest(`[${PPTX_ATTRS.cell}]`)
    ?.getAttribute(PPTX_ATTRS.cell);
  if (!value) return undefined;
  const [row, column] = value.split(",").map(Number);
  return { row, column };
}

export function getIsSameCell(
  a?: TableCellAddress,
  b?: TableCellAddress,
): boolean {
  return a?.row === b?.row && a?.column === b?.column;
}

export function getPlainText(paragraphs: SetTextBodyParagraph[]): string {
  return paragraphs
    .map((paragraph) => paragraph.runs.map((run) => run.text).join(""))
    .join("\n");
}

export function textBodyChanged(
  node: SlideNode,
  readBack: SetTextBodyParagraph[],
  cell?: TableCellAddress,
): boolean {
  const paragraphs = getTextBody(node, cell)?.paragraphs;
  if (!paragraphs) return false;
  if (paragraphs.length !== readBack.length) return true;
  for (let i = 0; i < paragraphs.length; i++) {
    const origText = paragraphs[i].runs.map((run) => run.text ?? "").join("");
    const newText = readBack[i].runs.map((run) => run.text).join("");
    if (origText !== newText) return true;
  }
  return false;
}
