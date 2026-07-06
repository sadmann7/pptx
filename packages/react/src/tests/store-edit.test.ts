// @vitest-environment jsdom
/**
 * Store-level editing tests: `edit()`, `undo()`/`redo()`, `save()`, and
 * per-slide revision tracking.
 *
 * Runs under jsdom (not the default happy-dom) because happy-dom's XML
 * parser drops namespaced attributes (e.g. `r:id` on `p:sldId`), which the
 * slide-level edit operations depend on.
 */
import type { ShapeNodeData } from "@diceui/pptx-parser";
import { beforeAll, describe, expect, it } from "vitest";

import { createPresentationStore, PresentationStore } from "../store";
import { buildMinimalPptx, FIXTURE_SLIDE_COUNT } from "./minimal-pptx";

let fixture: ArrayBuffer;

beforeAll(async () => {
  fixture = await buildMinimalPptx();
});

async function editableStore(): Promise<PresentationStore> {
  const store = createPresentationStore();
  // embedFonts off: the fixture has none, and the font pipeline needs workers.
  await store.load(fixture, { readOnly: false, embedFonts: false });
  return store;
}

function slideText(store: PresentationStore, index: number): string | undefined {
  const slide = store.getState().presentation?.slides[index];
  const shape = slide?.nodes.find((n) => n.nodeType === "shape") as ShapeNodeData | undefined;
  return shape?.textBody?.paragraphs[0]?.runs[0]?.text;
}

describe("store.edit", () => {
  it("rejects when no presentation is loaded", async () => {
    const store = createPresentationStore();
    await expect(
      store.edit({ type: "deleteSlide", slideId: "ppt/slides/slide1.xml" }),
    ).rejects.toThrow(/no presentation/);
  });

  it("rejects when the deck was loaded without readOnly: false", async () => {
    const store = createPresentationStore();
    await store.load(fixture, { embedFonts: false });
    const slideId = store.getState().presentation!.slides[0].id;
    await expect(
      store.edit({
        type: "setTextRun",
        slideId,
        nodeId: "2",
        paragraphIndex: 0,
        runIndex: 0,
        text: "x",
      }),
    ).rejects.toThrow(/keepPackage|readOnly/);
  });

  it("applies an edit, bumps revisions, and notifies subscribers", async () => {
    const store = await editableStore();
    const slideId = store.getState().presentation!.slides[0].id;

    let notified = 0;
    store.subscribe(() => {
      notified += 1;
    });

    expect(store.getSlideRevision(slideId)).toBe(0);
    await store.edit({
      type: "setTextRun",
      slideId,
      nodeId: "2",
      paragraphIndex: 0,
      runIndex: 0,
      text: "Edited",
    });

    expect(slideText(store, 0)).toBe("Edited");
    expect(store.getSlideRevision(slideId)).toBe(1);
    expect(store.getState().revision).toBe(1);
    expect(notified).toBeGreaterThan(0);
    // Other slides' revisions are untouched.
    expect(store.getSlideRevision(store.getState().presentation!.slides[1].id)).toBe(0);
  });

  it("save() round-trips edits through a real .pptx", async () => {
    const store = await editableStore();
    const slideId = store.getState().presentation!.slides[0].id;
    await store.edit({
      type: "setTextRun",
      slideId,
      nodeId: "2",
      paragraphIndex: 0,
      runIndex: 0,
      text: "Persisted",
    });

    const bytes = await store.save();
    const reopened = createPresentationStore();
    await reopened.load(bytes.slice().buffer as ArrayBuffer, { embedFonts: false });
    expect(slideText(reopened, 0)).toBe("Persisted");
  });
});

describe("undo / redo", () => {
  it("walks the history in both directions", async () => {
    const store = await editableStore();
    const slideId = store.getState().presentation!.slides[0].id;
    const editOf = (text: string) =>
      store.edit({
        type: "setTextRun",
        slideId,
        nodeId: "2",
        paragraphIndex: 0,
        runIndex: 0,
        text,
      });

    expect(store.canUndo()).toBe(false);
    expect(store.canRedo()).toBe(false);

    await editOf("First");
    await editOf("Second");
    expect(slideText(store, 0)).toBe("Second");
    expect(store.canUndo()).toBe(true);

    expect(store.undo()).toBe(true);
    expect(slideText(store, 0)).toBe("First");
    expect(store.canRedo()).toBe(true);

    expect(store.undo()).toBe(true);
    expect(slideText(store, 0)).toBe("Slide 1"); // fixture original
    expect(store.undo()).toBe(false);

    await expect(store.redo()).resolves.toBe(true);
    expect(slideText(store, 0)).toBe("First");
    await expect(store.redo()).resolves.toBe(true);
    expect(slideText(store, 0)).toBe("Second");
    await expect(store.redo()).resolves.toBe(false);
  });

  it("a new edit clears the redo stack", async () => {
    const store = await editableStore();
    const slideId = store.getState().presentation!.slides[0].id;
    const editOf = (text: string) =>
      store.edit({
        type: "setTextRun",
        slideId,
        nodeId: "2",
        paragraphIndex: 0,
        runIndex: 0,
        text,
      });

    await editOf("A");
    store.undo();
    await editOf("B");
    expect(store.canRedo()).toBe(false);
    expect(slideText(store, 0)).toBe("B");
  });

  it("bumps the slide revision on undo so views re-render", async () => {
    const store = await editableStore();
    const slideId = store.getState().presentation!.slides[0].id;
    await store.edit({
      type: "setTextRun",
      slideId,
      nodeId: "2",
      paragraphIndex: 0,
      runIndex: 0,
      text: "x",
    });
    const before = store.getSlideRevision(slideId);
    store.undo();
    expect(store.getSlideRevision(slideId)).toBe(before + 1);
  });
});

describe("structural edits and navigation", () => {
  it("keeps the active slide id stable across reorders", async () => {
    const store = await editableStore();
    const slides = store.getState().presentation!.slides;
    const firstId = slides[0].id;

    await store.edit({ type: "moveSlide", slideId: firstId, toIndex: 2 });

    expect(store.getState().activeSlideId).toBe(firstId);
    expect(store.getActiveSlideIndex()).toBe(2);
    expect(store.getSlideIndex(firstId)).toBe(2);
  });

  it("moves the active slide to a neighbor when it is deleted", async () => {
    const store = await editableStore();
    const slides = store.getState().presentation!.slides;
    store.goTo(slides[1].id);

    await store.edit({ type: "deleteSlide", slideId: slides[1].id });

    const state = store.getState();
    expect(state.presentation!.slides).toHaveLength(FIXTURE_SLIDE_COUNT - 1);
    // Falls back to the slide now occupying the old index.
    expect(state.activeSlideId).toBe(state.presentation!.slides[1].id);
    // The replacement slide is materialized and renderable.
    expect(store.getActiveSlide()?.nodesMaterialized).toBe(true);
  });

  it("duplicateSlide inserts a navigable copy and undo removes it", async () => {
    const store = await editableStore();
    const sourceId = store.getState().presentation!.slides[0].id;

    const result = await store.edit({ type: "duplicateSlide", slideId: sourceId });
    expect(store.getState().presentation!.slides).toHaveLength(FIXTURE_SLIDE_COUNT + 1);
    expect(store.getSlideIndex(result.createdSlideId!)).toBe(1);

    store.goTo(result.createdSlideId!);
    expect(store.getActiveSlideIndex()).toBe(1);

    store.undo();
    expect(store.getState().presentation!.slides).toHaveLength(FIXTURE_SLIDE_COUNT);
    // Active moved off the removed copy onto a valid neighbor.
    expect(store.getState().activeSlideId).not.toBe(result.createdSlideId);
    expect(store.getActiveSlideIndex()).toBeGreaterThanOrEqual(0);
  });
});

describe("history lifecycle", () => {
  it("clears edit history and revisions on a new load", async () => {
    const store = await editableStore();
    const slideId = store.getState().presentation!.slides[0].id;
    await store.edit({
      type: "setTextRun",
      slideId,
      nodeId: "2",
      paragraphIndex: 0,
      runIndex: 0,
      text: "x",
    });
    expect(store.canUndo()).toBe(true);

    await store.load(fixture, { readOnly: false, embedFonts: false });
    expect(store.canUndo()).toBe(false);
    expect(store.canRedo()).toBe(false);
    expect(store.getState().revision).toBe(0);
    expect(store.getSlideRevision(slideId)).toBe(0);
  });
});
