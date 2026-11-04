/**
 * Mapping between the rendered text DOM and model text positions, plus the
 * formatting a toolbar shows for the current text.
 *
 * Positions follow `TextParagraph.runs`: a paragraph's offset counts the
 * characters of its runs, with a line break (`a:br`) as one character.
 */

import type { ParagraphAlign, TextPosition, TextRange } from "@diceui/pptx-core";
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

/**
 * Strip zero-width spaces: line-height spacer spans from the renderer must
 * not reach the model.
 */
export function cleanText(raw: string | null | undefined): string {
  return (raw ?? "").replace(/​/g, "");
}

/**
 * Paragraph elements of a text container in model order. A container whose
 * paragraphs were flattened by editing is treated as one paragraph.
 */
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

/**
 * A `<br>` the renderer emitted for an `a:br` run, rather than the one a
 * browser adds to keep an empty line open: it has content after it.
 */
function isRenderedBreak(element: HTMLBRElement): boolean {
  return element.nextSibling !== null;
}

/**
 * Text-bearing leaves of a paragraph in order, skipping bullets. Line breaks
 * count only when `countBreaks` is set: a paragraph rendered straight from
 * the model has them, while text read back after editing drops them.
 */
function getTextLeaves(paragraph: HTMLElement, countBreaks: boolean): (Text | HTMLBRElement)[] {
  const leaves: (Text | HTMLBRElement)[] = [];
  const walker = document.createTreeWalker(
    paragraph,
    NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT,
    {
      acceptNode: (node) =>
        node instanceof HTMLElement && node.dataset[PPTX_DATASET.bullet] !== undefined
          ? NodeFilter.FILTER_REJECT
          : NodeFilter.FILTER_ACCEPT,
    },
  );
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (node.nodeType === Node.TEXT_NODE) {
      leaves.push(node as Text);
    } else if (countBreaks && node instanceof HTMLBRElement && isRenderedBreak(node)) {
      leaves.push(node);
    }
  }
  return leaves;
}

function getLeafLength(leaf: Text | HTMLBRElement): number {
  return leaf instanceof HTMLBRElement ? 1 : cleanText(leaf.data).length;
}

/** The text of a paragraph element as the model counts it. */
export function getParagraphText(paragraph: HTMLElement, countBreaks: boolean): string {
  return getTextLeaves(paragraph, countBreaks)
    .map((leaf) => (leaf instanceof HTMLBRElement ? "\n" : cleanText(leaf.data)))
    .join("");
}

/** Maps a DOM point inside a text container to a model text position. */
export function getTextPosition(
  container: HTMLElement,
  node: Node,
  offset: number,
  countBreaks: boolean,
): TextPosition | null {
  const paragraphs = getParagraphElements(container);

  // A point between paragraphs (on the container itself) snaps to the start
  // of the paragraph after it, or the end of the last one.
  if (node === container && paragraphs[0] !== container) {
    if (offset < paragraphs.length) return { paragraphIndex: offset, offset: 0 };
    const last = paragraphs.length - 1;
    return {
      paragraphIndex: last,
      offset: getParagraphText(paragraphs[last], countBreaks).length,
    };
  }

  const paragraphIndex = paragraphs.findIndex((paragraph) => paragraph.contains(node));
  if (paragraphIndex < 0) return null;
  const paragraph = paragraphs[paragraphIndex];

  const point = document.createRange();
  point.setStart(node, offset);

  let count = 0;
  for (const leaf of getTextLeaves(paragraph, countBreaks)) {
    if (leaf === node) {
      count += cleanText((leaf as Text).data.slice(0, offset)).length;
      break;
    }
    // Stop at the first leaf that starts at or after the point.
    if (point.comparePoint(leaf, 0) >= 0) break;
    count += getLeafLength(leaf);
  }
  return { paragraphIndex, offset: count };
}

/** Maps a DOM range inside a text container to a model text range. */
export function getTextRange(
  container: HTMLElement,
  range: AbstractRange,
  countBreaks: boolean,
): TextRange | null {
  const start = getTextPosition(container, range.startContainer, range.startOffset, countBreaks);
  const end = getTextPosition(container, range.endContainer, range.endOffset, countBreaks);
  return start && end ? { start, end } : null;
}

/**
 * Maps a model text position back to a DOM point in freshly rendered text.
 * The inverse of `getTextPosition` with breaks counted.
 */
export function getDomPoint(
  container: HTMLElement,
  position: TextPosition,
): [node: Node, offset: number] | null {
  const paragraph = getParagraphElements(container)[position.paragraphIndex];
  if (!paragraph) return null;

  let remaining = position.offset;
  for (const leaf of getTextLeaves(paragraph, true)) {
    const length = getLeafLength(leaf);
    if (remaining <= length && leaf instanceof Text) return [leaf, remaining];
    if (remaining === 0) {
      return [leaf.parentNode!, Array.prototype.indexOf.call(leaf.parentNode!.childNodes, leaf)];
    }
    remaining -= length;
  }
  return [paragraph, paragraph.childNodes.length];
}

const WORD_CHARACTER = /[\p{L}\p{N}_'’-]/u;

/**
 * The word around a collapsed caret, the way PowerPoint formats the word
 * under the caret. Returns `null` when the caret is not touching a word.
 */
export function expandToWord(text: string, offset: number): [start: number, end: number] | null {
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

/**
 * The rendered font size in points. The renderer writes sizes inline in `pt`,
 * which CSS `zoom` leaves alone, unlike the computed pixel size.
 */
function getFontSize(element: HTMLElement): number | null {
  for (let current: HTMLElement | null = element; current; current = current.parentElement) {
    const size = current.style.fontSize;
    if (size.endsWith("pt")) return Math.round(Number.parseFloat(size) * 10) / 10;
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

/** One value when all agree, else `null`. */
function getShared<T>(values: T[]): T | null {
  return values.every((value) => value === values[0]) ? (values[0] ?? null) : null;
}

/**
 * Reads the formatting of rendered runs and paragraphs. Toggles are on only
 * when every run has them; values are `null` when the runs disagree.
 */
export function getActiveTextStyle(
  runs: HTMLElement[],
  paragraphs: HTMLElement[],
): ActiveTextStyle | null {
  if (runs.length === 0) return null;
  const styles = runs.map((run) => {
    const computed = getComputedStyle(run);
    const decoration = `${computed.textDecorationLine} ${run.style.textDecoration}`;
    return {
      bold: Number.parseInt(computed.fontWeight, 10) >= 600 || computed.fontWeight === "bold",
      italic: computed.fontStyle === "italic" || computed.fontStyle === "oblique",
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
    align: getShared(paragraphs.map((paragraph) => toAlign(getComputedStyle(paragraph).textAlign))),
  };
}

/**
 * Run spans a DOM range touches inside a text container. A collapsed caret
 * reads the run it sits in, falling back to its paragraph's first run (or the
 * paragraph itself, which carries the paragraph mark's size, when empty).
 */
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
  const allRuns = Array.from(container.querySelectorAll<HTMLElement>(`[${PPTX_ATTRS.run}]`));

  if (!live.collapsed) {
    const runs = allRuns.filter((run) => {
      if (!live.intersectsNode(run)) return false;
      // A range that ends at a run's very start (or starts at its very end)
      // touches it without selecting any of its text.
      return live.comparePoint(run, run.childNodes.length) > -1 && live.comparePoint(run, 0) < 1;
    });
    if (runs.length > 0) return { runs, paragraphs };
  }

  const start = range.startContainer;
  const startElement = start instanceof HTMLElement ? start : start.parentElement;
  const run = startElement?.closest<HTMLElement>(`[${PPTX_ATTRS.run}]`);
  if (run && container.contains(run)) return { runs: [run], paragraphs };
  const paragraph = paragraphs[0] ?? getParagraphElements(container)[0];
  if (!paragraph) return { runs: [], paragraphs };
  const firstRun = paragraph.querySelector<HTMLElement>(`[${PPTX_ATTRS.run}]`);
  return { runs: [firstRun ?? paragraph], paragraphs: [paragraph] };
}
