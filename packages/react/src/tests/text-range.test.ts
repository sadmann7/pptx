import { describe, expect, it } from "vitest";

import {
  expandToWord,
  getActiveTextStyle,
  getDomPoint,
  getParagraphText,
  getRangeRuns,
  getTextPosition,
  getTextRange,
  readBackTextBody,
} from "../selection/text-range";

function textContainer(html: string): HTMLElement {
  const container = document.createElement("div");
  container.innerHTML = html;
  document.body.appendChild(container);
  return container;
}

const BULLETED = `<div data-pptx-p="0"><span data-pptx-bullet="">•</span><span data-pptx-r="0">Hello </span><span data-pptx-r="1">world</span></div><div data-pptx-p="1"><span data-pptx-r="0">Two</span><br><span data-pptx-r="2">lines</span></div>`;

describe("getTextPosition", () => {
  it("counts characters across runs and skips bullets", () => {
    const container = textContainer(BULLETED);
    const world = container.querySelector('[data-pptx-r="1"]')!.firstChild!;

    expect(getTextPosition(container, world, 2)).toEqual({
      paragraphIndex: 0,
      offset: 8,
    });
    container.remove();
  });

  it("counts a rendered line break as one character", () => {
    const container = textContainer(BULLETED);
    const lines = container.querySelector('[data-pptx-r="2"]')!.firstChild!;

    expect(getTextPosition(container, lines, 1)).toEqual({
      paragraphIndex: 1,
      offset: 5,
    });
    expect(getParagraphText(container.children[1] as HTMLElement)).toBe(
      "Two\nlines",
    );
    container.remove();
  });

  it("maps element-level points and points between paragraphs", () => {
    const container = textContainer(BULLETED);
    const firstParagraph = container.children[0];

    // After the bullet and the first run.
    expect(getTextPosition(container, firstParagraph, 2)).toEqual({
      paragraphIndex: 0,
      offset: 6,
    });
    expect(getTextPosition(container, container, 1)).toEqual({
      paragraphIndex: 1,
      offset: 0,
    });
    expect(getTextPosition(container, container, 2)).toEqual({
      paragraphIndex: 1,
      offset: 9,
    });
    expect(getTextPosition(container, document.body, 0)).toBeNull();
    container.remove();
  });

  it("maps a DOM range spanning paragraphs", () => {
    const container = textContainer(BULLETED);
    const range = document.createRange();
    range.setStart(
      container.querySelector('[data-pptx-r="1"]')!.firstChild!,
      1,
    );
    range.setEnd(
      container.children[1].querySelector('[data-pptx-r="0"]')!.firstChild!,
      2,
    );

    expect(getTextRange(container, range)).toEqual({
      start: { paragraphIndex: 0, offset: 7 },
      end: { paragraphIndex: 1, offset: 2 },
    });
    container.remove();
  });
});

describe("getDomPoint", () => {
  it("is the inverse of getTextPosition", () => {
    const container = textContainer(BULLETED);
    for (const position of [
      { paragraphIndex: 0, offset: 0 },
      { paragraphIndex: 0, offset: 8 },
      { paragraphIndex: 0, offset: 11 },
      { paragraphIndex: 1, offset: 2 },
      { paragraphIndex: 1, offset: 6 },
    ]) {
      const point = getDomPoint(container, position)!;
      expect(getTextPosition(container, point[0], point[1])).toEqual(position);
    }
    expect(getDomPoint(container, { paragraphIndex: 5, offset: 0 })).toBeNull();
    container.remove();
  });
});

describe("expandToWord", () => {
  it("expands a caret to the word around it", () => {
    expect(expandToWord("Hello world", 2)).toEqual([0, 5]);
    expect(expandToWord("Hello world", 5)).toEqual([0, 5]);
    expect(expandToWord("Hello world", 11)).toEqual([6, 11]);
    expect(expandToWord("don't stop", 1)).toEqual([0, 5]);
    expect(expandToWord("Café olé", 2)).toEqual([0, 4]);
  });

  it("returns null between words", () => {
    expect(expandToWord("a  b", 2)).toBeNull();
    expect(expandToWord("", 0)).toBeNull();
  });
});

describe("getActiveTextStyle", () => {
  it("turns toggles on only when every run has them, and nulls mixed values", () => {
    const container = textContainer(
      `<div data-pptx-p="0" style="text-align: center">` +
        `<span data-pptx-r="0" style="font-weight: bold; font-size: 24pt; color: rgb(255, 0, 0); font-family: Arial; text-decoration: underline">A</span>` +
        `<span data-pptx-r="1" style="font-weight: bold; font-size: 18pt; color: rgb(255, 0, 0); font-family: Arial">B</span>` +
        `</div>`,
    );
    const runs = Array.from(
      container.querySelectorAll<HTMLElement>("[data-pptx-r]"),
    );
    const paragraphs = [container.firstElementChild as HTMLElement];

    expect(getActiveTextStyle(runs, paragraphs)).toEqual({
      bold: true,
      italic: false,
      underline: false,
      strikethrough: false,
      color: "#FF0000",
      fontFamily: "Arial",
      fontSize: null,
      align: "center",
    });
    expect(getActiveTextStyle(runs.slice(0, 1), paragraphs)).toMatchObject({
      underline: true,
      fontSize: 24,
    });
    expect(getActiveTextStyle([], paragraphs)).toBeNull();
    container.remove();
  });
});

describe("getActiveTextStyle font size", () => {
  it("reports the size the file stores, before autofit shrinks it", () => {
    const container = textContainer(
      `<div data-pptx-p="0"><span data-pptx-r="0" style="font-size: 15pt">A</span></div>`,
    );
    container.setAttribute("data-pptx-font-scale", "0.625");
    const runs = Array.from(
      container.querySelectorAll<HTMLElement>("[data-pptx-r]"),
    );

    expect(getActiveTextStyle(runs, [])?.fontSize).toBe(24);
    container.remove();
  });
});

describe("getRangeRuns", () => {
  it("skips a run the range only touches at its edge", () => {
    const container = textContainer(BULLETED);
    const [hello, world] = Array.from(
      container.querySelectorAll<HTMLElement>("[data-pptx-r]"),
    );
    const range = document.createRange();
    range.setStart(hello.firstChild!, 0);
    range.setEnd(world.firstChild!, 0);

    expect(getRangeRuns(container, range).runs).toEqual([hello]);
    container.remove();
  });
});

describe("readBackTextBody", () => {
  it("keeps rendered line breaks and drops the trailing one", () => {
    const container = textContainer(
      `${BULLETED}<div data-pptx-p="2"><span data-pptx-r="0">End</span><br></div>`,
    );

    expect(readBackTextBody(container)).toEqual([
      {
        sourceParagraphIndex: 0,
        runs: [
          { text: "Hello ", sourceRun: [0, 0] },
          { text: "world", sourceRun: [0, 1] },
        ],
      },
      {
        sourceParagraphIndex: 1,
        runs: [
          { text: "Two", sourceRun: [1, 0] },
          { text: "\n", sourceRun: [1, 0] },
          { text: "lines", sourceRun: [1, 2] },
        ],
      },
      { sourceParagraphIndex: 2, runs: [{ text: "End", sourceRun: [2, 0] }] },
    ]);
    container.remove();
  });
});
