import { act } from "react";

import type { SlideNode } from "@diceui/pptx-core";
import { PPTX_DATASET } from "@diceui/pptx-core";
import { fireEvent, render, waitFor } from "@testing-library/react";
import { beforeAll, describe, expect, it, vi } from "vitest";

import { Presentation } from "../index";
import type { GripDirection, Rect } from "../selection";
import {
  MIN_SIZE,
  cleanText,
  constrainMove,
  getNodeRect,
  mergeSelection,
  readBackTextBody,
  resizeRect,
  textBodyChanged,
  toggleSelection,
} from "../selection";
import { createStore, type Store } from "../store";
import { loadFixture } from "./test-utils";

// ===========================================================================
// Unit tests: selection-utils (geometry + text read-back)
// ===========================================================================

describe("resizeRect", () => {
  const origin: Rect = { x: 100, y: 100, w: 200, h: 100 };

  describe("cardinal handles", () => {
    it("e: grows width rightward", () => {
      expect(resizeRect(origin, "e", 50, 0)).toEqual({ x: 100, y: 100, w: 250, h: 100 });
    });

    it("w: grows width leftward (shifts x)", () => {
      expect(resizeRect(origin, "w", -50, 0)).toEqual({ x: 50, y: 100, w: 250, h: 100 });
    });

    it("s: grows height downward", () => {
      expect(resizeRect(origin, "s", 0, 30)).toEqual({ x: 100, y: 100, w: 200, h: 130 });
    });

    it("n: grows height upward (shifts y)", () => {
      expect(resizeRect(origin, "n", 0, -40)).toEqual({ x: 100, y: 60, w: 200, h: 140 });
    });
  });

  describe("corner handles", () => {
    it("se: grows both dimensions", () => {
      expect(resizeRect(origin, "se", 20, 10)).toEqual({ x: 100, y: 100, w: 220, h: 110 });
    });

    it("nw: grows both dimensions (shifts x and y)", () => {
      expect(resizeRect(origin, "nw", -30, -20)).toEqual({ x: 70, y: 80, w: 230, h: 120 });
    });

    it("ne: grows width right, height up", () => {
      expect(resizeRect(origin, "ne", 10, -15)).toEqual({ x: 100, y: 85, w: 210, h: 115 });
    });

    it("sw: grows width left, height down", () => {
      expect(resizeRect(origin, "sw", -10, 15)).toEqual({ x: 90, y: 100, w: 210, h: 115 });
    });
  });

  describe("min-size clamping", () => {
    it("clamps width to MIN_SIZE when e shrinks past it", () => {
      const result = resizeRect(origin, "e", -300, 0);
      expect(result.w).toBe(MIN_SIZE);
      expect(result.x).toBe(origin.x);
    });

    it("clamps width and re-anchors x when w shrinks past MIN_SIZE", () => {
      const result = resizeRect(origin, "w", 300, 0);
      expect(result.w).toBe(MIN_SIZE);
      expect(result.x).toBe(origin.x + origin.w - MIN_SIZE);
    });

    it("clamps height to MIN_SIZE when s shrinks past it", () => {
      const result = resizeRect(origin, "s", 0, -200);
      expect(result.h).toBe(MIN_SIZE);
      expect(result.y).toBe(origin.y);
    });

    it("clamps height and re-anchors y when n shrinks past MIN_SIZE", () => {
      const result = resizeRect(origin, "n", 0, 200);
      expect(result.h).toBe(MIN_SIZE);
      expect(result.y).toBe(origin.y + origin.h - MIN_SIZE);
    });
  });

  describe("aspect lock", () => {
    const square: Rect = { x: 50, y: 50, w: 100, h: 100 };

    it("scales proportionally for corner handles", () => {
      const result = resizeRect(square, "se", 50, 0, true);
      expect(result.w).toBe(result.h);
      expect(result.w).toBe(150);
    });

    it("re-anchors opposite corner for nw handle", () => {
      const result = resizeRect(square, "nw", -20, 0, true);
      expect(result.w).toBeCloseTo(120);
      expect(result.h).toBeCloseTo(120);
      expect(result.x).toBeCloseTo(50 + 100 - 120);
      expect(result.y).toBeCloseTo(50 + 100 - 120);
    });

    it("does not apply aspect lock for edge handles", () => {
      const result = resizeRect(square, "e", 50, 0, true);
      expect(result.w).toBe(150);
      expect(result.h).toBe(100);
    });

    it("enforces MIN_SIZE during aspect lock", () => {
      const result = resizeRect(square, "se", -200, -200, true);
      expect(result.w).toBeGreaterThanOrEqual(MIN_SIZE);
      expect(result.h).toBeGreaterThanOrEqual(MIN_SIZE);
    });
  });

  describe("zero deltas", () => {
    it("returns same rect when dx=0 dy=0 for any handle", () => {
      const grips: GripDirection[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];
      for (const h of grips) {
        expect(resizeRect(origin, h, 0, 0)).toEqual(origin);
      }
    });
  });
});

describe("constrainMove", () => {
  it("passes both deltas through when unlocked", () => {
    expect(constrainMove(30, -12)).toEqual({ x: 30, y: -12 });
  });

  it("drops the vertical delta when the drag is mostly horizontal", () => {
    expect(constrainMove(40, 9, true)).toEqual({ x: 40, y: 0 });
    expect(constrainMove(-40, -9, true)).toEqual({ x: -40, y: 0 });
  });

  it("drops the horizontal delta when the drag is mostly vertical", () => {
    expect(constrainMove(9, 40, true)).toEqual({ x: 0, y: 40 });
    expect(constrainMove(-9, -40, true)).toEqual({ x: 0, y: -40 });
  });

  it("keeps the horizontal delta on an exact diagonal", () => {
    expect(constrainMove(25, -25, true)).toEqual({ x: 25, y: 0 });
  });

  it("stays put when locked with no movement", () => {
    expect(constrainMove(0, 0, true)).toEqual({ x: 0, y: 0 });
  });
});

describe("mergeSelection", () => {
  it("replaces the selection when there is no base", () => {
    expect(mergeSelection([], ["b", "c"])).toEqual(["b", "c"]);
  });

  it("appends newly enclosed ids after the base selection", () => {
    expect(mergeSelection(["a"], ["b", "c"])).toEqual(["a", "b", "c"]);
  });

  it("keeps already selected shapes instead of toggling them out", () => {
    expect(mergeSelection(["a", "b"], ["b"])).toEqual(["a", "b"]);
  });

  it("keeps the base selection when the band enclosed nothing", () => {
    expect(mergeSelection(["a", "b"], [])).toEqual(["a", "b"]);
  });
});

describe("toggleSelection", () => {
  it("adds a shape that is not selected yet", () => {
    expect(toggleSelection(["a"], "b")).toEqual(["a", "b"]);
  });

  it("takes an already selected shape back out", () => {
    expect(toggleSelection(["a", "b"], "a")).toEqual(["b"]);
  });

  it("empties the selection when the last shape is toggled out", () => {
    expect(toggleSelection(["a"], "a")).toEqual([]);
  });

  it("selects the shape when nothing was selected", () => {
    expect(toggleSelection([], "a")).toEqual(["a"]);
  });
});

describe("getNodeRect", () => {
  it("maps node position and size to Rect", () => {
    const node = {
      id: "n1",
      name: "Rect",
      nodeType: "shape",
      position: { x: 10, y: 20 },
      size: { w: 300, h: 150 },
    } as unknown as SlideNode;
    expect(getNodeRect(node)).toEqual({ x: 10, y: 20, w: 300, h: 150 });
  });
});

describe("cleanText", () => {
  it("strips zero-width spaces", () => {
    expect(cleanText("hello\u200Bworld")).toBe("helloworld");
  });

  it("passes ordinary text unchanged", () => {
    expect(cleanText("hello world")).toBe("hello world");
  });

  it("handles null", () => {
    expect(cleanText(null)).toBe("");
  });

  it("handles undefined", () => {
    expect(cleanText(undefined)).toBe("");
  });
});

describe("readBackTextBody", () => {
  function makePara(pIndex: number, runs: { rIndex: number; text: string }[]): HTMLElement {
    const div = document.createElement("div");
    div.dataset[PPTX_DATASET.paragraph] = String(pIndex);
    for (const r of runs) {
      const span = document.createElement("span");
      span.dataset[PPTX_DATASET.run] = String(r.rIndex);
      span.textContent = r.text;
      div.appendChild(span);
    }
    return div;
  }

  it("reads typical paragraph → run structure", () => {
    const container = document.createElement("div");
    container.appendChild(
      makePara(0, [
        { rIndex: 0, text: "Hello " },
        { rIndex: 1, text: "World" },
      ]),
    );
    container.appendChild(makePara(1, [{ rIndex: 0, text: "Second" }]));

    const result = readBackTextBody(container);
    expect(result).toHaveLength(2);
    expect(result[0].sourceParagraphIndex).toBe(0);
    expect(result[0].runs).toEqual([
      { text: "Hello ", sourceRun: [0, 0] },
      { text: "World", sourceRun: [0, 1] },
    ]);
    expect(result[1].runs).toEqual([{ text: "Second", sourceRun: [1, 0] }]);
  });

  it("skips bullet spans", () => {
    const container = document.createElement("div");
    const para = document.createElement("div");
    para.dataset[PPTX_DATASET.paragraph] = "0";
    const bullet = document.createElement("span");
    bullet.dataset[PPTX_DATASET.bullet] = "";
    bullet.textContent = "•";
    para.appendChild(bullet);
    const run = document.createElement("span");
    run.dataset[PPTX_DATASET.run] = "0";
    run.textContent = "Item";
    para.appendChild(run);
    container.appendChild(para);

    const result = readBackTextBody(container);
    expect(result[0].runs).toEqual([{ text: "Item", sourceRun: [0, 0] }]);
  });

  it("skips <br> elements", () => {
    const container = document.createElement("div");
    const para = document.createElement("div");
    para.dataset[PPTX_DATASET.paragraph] = "0";
    para.appendChild(document.createElement("br"));
    container.appendChild(para);

    const result = readBackTextBody(container);
    expect(result[0].runs).toEqual([{ text: "", sourceRun: undefined }]);
  });

  it("bare text nodes inherit lastSourceR", () => {
    const container = document.createElement("div");
    const para = document.createElement("div");
    para.dataset[PPTX_DATASET.paragraph] = "0";
    const run = document.createElement("span");
    run.dataset[PPTX_DATASET.run] = "2";
    run.textContent = "styled";
    para.appendChild(run);
    para.appendChild(document.createTextNode("bare"));
    container.appendChild(para);

    const result = readBackTextBody(container);
    expect(result[0].runs).toEqual([
      { text: "styled", sourceRun: [0, 2] },
      { text: "bare", sourceRun: [0, 2] },
    ]);
  });

  it("falls back to container as single paragraph when para divs are gone", () => {
    const container = document.createElement("div");
    const run1 = document.createElement("span");
    run1.dataset[PPTX_DATASET.run] = "0";
    run1.textContent = "orphan";
    container.appendChild(run1);

    const result = readBackTextBody(container);
    expect(result).toHaveLength(1);
    expect(result[0].runs).toEqual([{ text: "orphan", sourceRun: [0, 0] }]);
  });

  it("empty paragraph produces empty-text run", () => {
    const container = document.createElement("div");
    const para = document.createElement("div");
    para.dataset[PPTX_DATASET.paragraph] = "0";
    container.appendChild(para);

    const result = readBackTextBody(container);
    expect(result[0].runs).toEqual([{ text: "", sourceRun: undefined }]);
  });
});

describe("textBodyChanged", () => {
  function makeShapeNode(paragraphs: { runs: { text: string }[] }[]): SlideNode {
    return {
      id: "s1",
      name: "Shape",
      nodeType: "shape",
      position: { x: 0, y: 0 },
      size: { w: 100, h: 50 },
      textBody: { paragraphs },
    } as unknown as SlideNode;
  }

  it("returns false for non-shape node", () => {
    const node = { nodeType: "picture" } as unknown as SlideNode;
    expect(textBodyChanged(node, [])).toBe(false);
  });

  it("returns false when shape has no textBody", () => {
    const node = { nodeType: "shape" } as unknown as SlideNode;
    expect(textBodyChanged(node, [])).toBe(false);
  });

  it("returns false when text matches", () => {
    const node = makeShapeNode([{ runs: [{ text: "Hello" }] }]);
    const readBack = [{ sourceParagraphIndex: 0, runs: [{ text: "Hello" }] }];
    expect(textBodyChanged(node, readBack)).toBe(false);
  });

  it("returns false when run split differs but concatenated text matches", () => {
    const node = makeShapeNode([{ runs: [{ text: "He" }, { text: "llo" }] }]);
    const readBack = [{ sourceParagraphIndex: 0, runs: [{ text: "Hello" }] }];
    expect(textBodyChanged(node, readBack)).toBe(false);
  });

  it("returns true when text differs", () => {
    const node = makeShapeNode([{ runs: [{ text: "Hello" }] }]);
    const readBack = [{ sourceParagraphIndex: 0, runs: [{ text: "Goodbye" }] }];
    expect(textBodyChanged(node, readBack)).toBe(true);
  });

  it("returns true when paragraph count differs", () => {
    const node = makeShapeNode([{ runs: [{ text: "A" }] }, { runs: [{ text: "B" }] }]);
    const readBack = [{ sourceParagraphIndex: 0, runs: [{ text: "AB" }] }];
    expect(textBodyChanged(node, readBack)).toBe(true);
  });
});

// ===========================================================================
// Integration tests: Selection component
// ===========================================================================

let fixture: ArrayBuffer;

beforeAll(async () => {
  fixture = await loadFixture();
});

async function editableStore(): Promise<Store> {
  const store = createStore();
  await store.load(fixture, { readOnly: false, embedFonts: false });
  return store;
}

/**
 * Renders a full Presentation tree with Slide + Selection, waits for ready,
 * and returns helpers for interacting with the overlay.
 */
async function renderSelection(props: Record<string, unknown> = {}) {
  const store = await editableStore();

  const { container } = render(
    <Presentation.Provider store={store}>
      <Presentation.Slide data-testid="slide">
        <Presentation.Selection data-testid="selection" {...props} />
      </Presentation.Slide>
    </Presentation.Provider>,
  );

  await waitFor(() => {
    expect(container.querySelector("[data-pptx-selection]")).not.toBeNull();
  });

  const overlay = container.querySelector<HTMLDivElement>("[data-pptx-selection]")!;
  const slideWrapper = overlay.parentElement!.parentElement!;
  const shapeElements = slideWrapper.querySelectorAll<HTMLElement>("[data-pptx-node-id]");

  return { store, container, overlay, slideWrapper, shapeElements };
}

// ---------------------------------------------------------------------------
// Mount guards
// ---------------------------------------------------------------------------

describe("Selection mount guards", () => {
  it("renders nothing when no presentation is loaded", () => {
    const store = createStore();
    const { container } = render(
      <Presentation.Provider store={store}>
        <Presentation.Slide>
          <Presentation.Selection data-testid="selection" />
        </Presentation.Slide>
      </Presentation.Provider>,
    );
    expect(container.querySelector("[data-pptx-selection]")).toBeNull();
  });

  it("shows overlay in idle mode when presentation is ready", async () => {
    const { overlay } = await renderSelection();
    expect(overlay).not.toBeNull();
    expect(overlay.getAttribute("data-mode")).toBe("idle");
  });
});

// ---------------------------------------------------------------------------
// Ctrl+A select all
// ---------------------------------------------------------------------------

describe("Ctrl+A select all", () => {
  it("selects all shapes on the slide", async () => {
    const { overlay, shapeElements } = await renderSelection();
    expect(shapeElements.length).toBeGreaterThan(0);

    act(() => {
      overlay.focus();
      fireEvent.keyDown(overlay, { key: "a", ctrlKey: true });
    });

    expect(overlay.getAttribute("data-mode")).toBe("selected");
  });
});

// ---------------------------------------------------------------------------
// Marquee selection
// ---------------------------------------------------------------------------

/**
 * happy-dom ships no `document.elementsFromPoint`, which `hitTest` needs, so
 * stub it as empty: every press then lands on empty canvas, which is the
 * precondition for a band. The flip side is that pressing a *shape* cannot be
 * simulated here, so Shift+drag on a shape stays manual-only.
 *
 * Containment still works, because node rects come from the model while only
 * the band corners come from the pointer. With layout zeroed, `clientToSlide`
 * reduces to `client / zoom`, so a band over the origin encloses the whole
 * slide and one far outside it encloses nothing.
 */
const OVER_SLIDE = 100_000;
const PAST_SLIDE = 90_000;

beforeAll(() => {
  document.elementsFromPoint ??= () => [];
});

function band(
  overlay: HTMLElement,
  from: number,
  to: number,
  modifier: { shiftKey?: boolean } = {},
): void {
  act(() => {
    fireEvent.pointerDown(overlay, {
      button: 0,
      clientX: from,
      clientY: from,
      pointerId: 1,
      ...modifier,
    });
  });
  act(() => {
    fireEvent.pointerMove(overlay, {
      buttons: 1,
      clientX: to,
      clientY: to,
      pointerId: 1,
      ...modifier,
    });
  });
  act(() => {
    fireEvent.pointerUp(overlay, { pointerId: 1, ...modifier });
  });
}

describe("marquee selection", () => {
  it("enters marquee mode on a press over empty canvas", async () => {
    const { overlay } = await renderSelection();

    act(() => {
      fireEvent.pointerDown(overlay, { button: 0, clientX: 0, clientY: 0, pointerId: 1 });
    });

    expect(overlay.getAttribute("data-mode")).toBe("marquee");
  });

  it("selects the shapes a band encloses", async () => {
    const { overlay } = await renderSelection();

    band(overlay, 0, OVER_SLIDE);

    expect(overlay.getAttribute("data-mode")).toBe("selected");
  });

  it("clears the selection when a plain band encloses nothing", async () => {
    const { overlay } = await renderSelection();

    act(() => {
      overlay.focus();
      fireEvent.keyDown(overlay, { key: "a", ctrlKey: true });
    });
    expect(overlay.getAttribute("data-mode")).toBe("selected");

    band(overlay, PAST_SLIDE, OVER_SLIDE);

    expect(overlay.getAttribute("data-mode")).toBe("idle");
  });

  it("keeps the selection when a Shift band encloses nothing", async () => {
    const { overlay } = await renderSelection();

    act(() => {
      overlay.focus();
      fireEvent.keyDown(overlay, { key: "a", ctrlKey: true });
    });

    band(overlay, PAST_SLIDE, OVER_SLIDE, { shiftKey: true });

    expect(overlay.getAttribute("data-mode")).toBe("selected");
  });
});

// ---------------------------------------------------------------------------
// Keyboard shortcuts (Escape, Delete, Arrow, Undo/Redo)
// ---------------------------------------------------------------------------

describe("keyboard shortcuts", () => {
  it("Escape from selected returns to idle", async () => {
    const { overlay } = await renderSelection();

    // Ctrl+A to select all first
    act(() => {
      overlay.focus();
      fireEvent.keyDown(overlay, { key: "a", ctrlKey: true });
    });
    expect(overlay.getAttribute("data-mode")).toBe("selected");

    act(() => {
      fireEvent.keyDown(overlay, { key: "Escape" });
    });
    expect(overlay.getAttribute("data-mode")).toBe("idle");
  });

  it("Delete key while selected triggers node deletion", async () => {
    const onNodeDelete = vi.fn();
    const { overlay } = await renderSelection({ onNodeDelete });

    act(() => {
      overlay.focus();
      fireEvent.keyDown(overlay, { key: "a", ctrlKey: true });
    });
    expect(overlay.getAttribute("data-mode")).toBe("selected");

    act(() => {
      fireEvent.keyDown(overlay, { key: "Delete" });
    });

    await waitFor(() => {
      expect(onNodeDelete).toHaveBeenCalled();
    });
  });

  it("Arrow key nudges selected node", async () => {
    const onNodeTransform = vi.fn();
    const { overlay } = await renderSelection({ onNodeTransform });

    act(() => {
      overlay.focus();
      fireEvent.keyDown(overlay, { key: "a", ctrlKey: true });
    });

    act(() => {
      fireEvent.keyDown(overlay, { key: "ArrowRight" });
    });

    await waitFor(() => {
      expect(onNodeTransform).toHaveBeenCalled();
    });
  });

  it("Ctrl+Z triggers undo callback", async () => {
    const onUndo = vi.fn();
    const store = await editableStore();

    const { container } = render(
      <Presentation.Provider store={store}>
        <Presentation.Slide>
          <Presentation.Selection onUndo={onUndo} />
        </Presentation.Slide>
      </Presentation.Provider>,
    );

    await waitFor(() => {
      expect(container.querySelector("[data-pptx-selection]")).not.toBeNull();
    });

    act(() => {
      fireEvent.keyDown(document, { key: "z", ctrlKey: true });
    });

    expect(onUndo).toHaveBeenCalledWith("empty");
  });

  it("Ctrl+Shift+Z triggers redo callback", async () => {
    const onRedo = vi.fn();
    const store = await editableStore();

    const { container } = render(
      <Presentation.Provider store={store}>
        <Presentation.Slide>
          <Presentation.Selection onRedo={onRedo} />
        </Presentation.Slide>
      </Presentation.Provider>,
    );

    await waitFor(() => {
      expect(container.querySelector("[data-pptx-selection]")).not.toBeNull();
    });

    act(() => {
      fireEvent.keyDown(document, { key: "z", ctrlKey: true, shiftKey: true });
    });

    await waitFor(() => {
      expect(onRedo).toHaveBeenCalled();
    });
  });
});

// ---------------------------------------------------------------------------
// render prop
// ---------------------------------------------------------------------------

describe("render prop", () => {
  it("passes SelectionState to render function", async () => {
    let capturedState: unknown = null;

    const store = await editableStore();
    const { container } = render(
      <Presentation.Provider store={store}>
        <Presentation.Slide>
          <Presentation.Selection
            render={(props, state) => {
              capturedState = state;
              return <div {...props} data-custom="yes" />;
            }}
          />
        </Presentation.Slide>
      </Presentation.Provider>,
    );

    await waitFor(() => {
      expect(container.querySelector("[data-custom]")).not.toBeNull();
    });

    expect(capturedState).toEqual(
      expect.objectContaining({
        mode: "idle",
        selectedNode: null,
        selectedNodes: [],
      }),
    );
  });
});
