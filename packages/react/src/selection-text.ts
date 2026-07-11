/**
 * DOM read-back for inline text editing: converts an edited contentEditable
 * container back into a `setTextBody` paragraphs payload, and detects
 * whether the text actually changed against the model.
 */

import type { SetTextBodyParagraph, ShapeNodeData, SlideNode } from "@diceui/pptx-core";
import { PPTX_DATASET } from "@diceui/pptx-core";

/**
 * Strip zero-width spaces: line-height spacer spans from the renderer must
 * not reach the model.
 */
export function cleanText(raw: string | null | undefined): string {
  return (raw ?? "").replace(/\u200B/g, "");
}

/**
 * Walk the edited contentEditable container and produce a `setTextBody`
 * paragraphs payload. Uses the paragraph/run data attributes (see
 * `PPTX_ATTRS`) to map back to source indices for style inheritance.
 */
export function readBackTextBody(container: HTMLElement): SetTextBodyParagraph[] {
  const paragraphs: SetTextBodyParagraph[] = [];

  const children = Array.from(container.children).filter(
    (el) => el instanceof HTMLElement,
  ) as HTMLElement[];

  // After destructive edits the paragraph divs may be gone, leaving run
  // spans (or bare text) directly in the container; treat the container
  // itself as a single implicit paragraph in that case.
  const paraDivs = children.filter((el) => el.dataset[PPTX_DATASET.run] === undefined);
  const effectiveDivs =
    paraDivs.length > 0 && paraDivs.length === children.length ? paraDivs : [container];
  let lastSourceP = 0;

  for (const paraDiv of effectiveDivs) {
    const srcPStr = paraDiv.dataset?.[PPTX_DATASET.paragraph];
    const sourceParagraphIndex = srcPStr !== undefined ? Number(srcPStr) : lastSourceP;
    lastSourceP = sourceParagraphIndex;

    const runs = readRunsFromParagraphDiv(paraDiv, sourceParagraphIndex);
    paragraphs.push({ sourceParagraphIndex, runs });
  }

  return paragraphs;
}

function readRunsFromParagraphDiv(
  paraDiv: HTMLElement,
  defaultSourceP: number,
): SetTextBodyParagraph["runs"] {
  const runs: SetTextBodyParagraph["runs"] = [];
  let lastSourceR: [number, number] | undefined;

  for (const child of Array.from(paraDiv.childNodes)) {
    if (child instanceof HTMLElement && child.dataset[PPTX_DATASET.bullet] !== undefined) {
      continue;
    }

    if (child instanceof HTMLElement && child.dataset[PPTX_DATASET.run] !== undefined) {
      const runIdx = Number(child.dataset[PPTX_DATASET.run]);
      const sourceRun: [number, number] = [defaultSourceP, runIdx];
      const text = cleanText(child.textContent);
      if (text.length > 0) {
        runs.push({ text, sourceRun });
        lastSourceR = sourceRun;
      }
    } else if (child instanceof HTMLBRElement) {
      // Browsers insert <br> for empty paragraphs; skip.
    } else if (child.nodeType === Node.TEXT_NODE) {
      const text = cleanText(child.textContent);
      if (text.length > 0) {
        runs.push({ text, sourceRun: lastSourceR });
      }
    } else if (child instanceof HTMLElement) {
      const text = cleanText(child.textContent);
      if (text.length > 0) {
        runs.push({ text, sourceRun: lastSourceR });
      }
    }
  }

  if (runs.length === 0) {
    runs.push({ text: "", sourceRun: lastSourceR });
  }

  return runs;
}

/** Compare the read-back paragraphs to the model to detect changes. */
export function textBodyChanged(node: SlideNode, readBack: SetTextBodyParagraph[]): boolean {
  if (node.nodeType !== "shape") return false;
  const shape = node as ShapeNodeData;
  const paragraphs = shape.textBody?.paragraphs;
  if (!paragraphs) return false;
  if (paragraphs.length !== readBack.length) return true;
  for (let i = 0; i < paragraphs.length; i++) {
    // Compare concatenated text, not run-by-run: a paragraph with no runs
    // reads back as a single empty run, which is not a real change (plain
    // shapes render an empty editable paragraph before any typing happens).
    const origText = paragraphs[i].runs.map((r) => r.text ?? "").join("");
    const newText = readBack[i].runs.map((r) => r.text).join("");
    if (origText !== newText) return true;
  }
  return false;
}
