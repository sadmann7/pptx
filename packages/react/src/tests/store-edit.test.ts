// @vitest-environment jsdom
/**
 * Store-level editing tests: `edit()`, `undo()`/`redo()`, `save()`, and
 * per-slide revision tracking.
 *
 * Runs under jsdom (not the default happy-dom) because happy-dom's XML
 * parser drops namespaced attributes (e.g. `r:id` on `p:sldId`), which the
 * slide-level edit operations depend on.
 */
import type { EditOperation, ShapeNodeData } from "@diceui/pptx-core";
import { beforeAll, describe, expect, it } from "vitest";

import type { EditEvent, HistoryChangeEvent, SlideChangeEvent } from "../store";
import { createStore, Store } from "../store";
import { FIXTURE_SLIDE_COUNT } from "./minimal-pptx";
import { editableStore, loadFixture } from "./test-utils";

let fixture: ArrayBuffer;

beforeAll(async () => {
  fixture = await loadFixture();
});

function slideText(store: Store, index: number): string | undefined {
  const slide = store.getState().presentation?.slides[index];
  const shape = slide?.nodes.find((n) => n.nodeType === "shape") as ShapeNodeData | undefined;
  return shape?.textBody?.paragraphs[0]?.runs[0]?.text;
}

describe("store.edit", () => {
  it("rejects when no presentation is loaded", async () => {
    const store = createStore();
    await expect(
      store.edit({ type: "deleteSlide", slideId: "ppt/slides/slide1.xml" }),
    ).rejects.toThrow(/no presentation/);
  });

  it("rejects when the deck was loaded without readOnly: false", async () => {
    const store = createStore();
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
    ).rejects.toThrow(/keepSourcePackage|readOnly/);
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
    const reopened = createStore();
    await reopened.load(bytes.slice().buffer as ArrayBuffer, { embedFonts: false });
    expect(slideText(reopened, 0)).toBe("Persisted");
  });

  it("batch-deletes a multi-selection as one undoable edit", async () => {
    const store = await editableStore();
    const slide = store.getState().presentation!.slides[0];
    const slideId = slide.id;
    const nodeIds = slide.nodes.map((n) => n.id);
    expect(nodeIds.length).toBeGreaterThan(0);

    await store.edit({
      type: "batch",
      operations: nodeIds.map((nodeId) => ({ type: "deleteNode" as const, slideId, nodeId })),
    });
    expect(store.getState().presentation!.slides[0].nodes).toHaveLength(0);

    // A single undo restores the whole selection.
    expect(store.undo()).toBe(true);
    expect(store.getState().presentation!.slides[0].nodes).toHaveLength(nodeIds.length);
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

  it("navigates to the edited slide on cross-slide undo", async () => {
    const store = await editableStore();
    const slides = store.getState().presentation!.slides;
    const slide0Id = slides[0].id;
    const slide1Id = slides[1].id;

    await store.edit({
      type: "setTextRun",
      slideId: slide0Id,
      nodeId: "2",
      paragraphIndex: 0,
      runIndex: 0,
      text: "edited",
    });
    store.goTo(slide1Id);
    expect(store.getState().activeSlideId).toBe(slide1Id);

    // Undo should navigate back to slide 0 (where the edit was).
    store.undo();
    expect(store.getState().activeSlideId).toBe(slide0Id);
    expect(slideText(store, 0)).toBe("Slide 1");

    // Redo should navigate to slide 0 again.
    await store.redo();
    expect(store.getState().activeSlideId).toBe(slide0Id);
    expect(slideText(store, 0)).toBe("edited");
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

describe("edit and historyChange events", () => {
  const textEdit = (slideId: string, text: string): EditOperation => ({
    type: "setTextRun",
    slideId,
    nodeId: "2",
    paragraphIndex: 0,
    runIndex: 0,
    text,
  });

  it("reports the operation and where it came from", async () => {
    const store = await editableStore();
    const slideId = store.getState().presentation!.slides[0].id;
    const events: EditEvent[] = [];
    store.on("edit", (event) => events.push(event));

    const operation = textEdit(slideId, "Edited");
    await store.edit(operation);
    store.undo();
    await store.redo();

    expect(events.map((event) => event.source)).toEqual(["edit", "undo", "redo"]);
    expect(events.every((event) => event.operation === operation)).toBe(true);
    expect(events[0].result.affectedSlideIds).toContain(slideId);
  });

  it("tracks undo/redo availability and dirtiness together", async () => {
    const store = await editableStore();
    const slideId = store.getState().presentation!.slides[0].id;
    const events: HistoryChangeEvent[] = [];
    store.on("historyChange", (event) => events.push(event));

    await store.edit(textEdit(slideId, "Edited"));
    expect(events.at(-1)).toEqual({ canUndo: true, canRedo: false, isDirty: true });

    store.undo();
    expect(events.at(-1)).toEqual({ canUndo: false, canRedo: true, isDirty: false });

    await store.redo();
    expect(events.at(-1)).toEqual({ canUndo: true, canRedo: false, isDirty: true });
  });

  it("stays quiet when undo and redo find nothing to do", async () => {
    const store = await editableStore();
    const events: HistoryChangeEvent[] = [];
    store.on("historyChange", (event) => events.push(event));

    expect(store.undo()).toBe(false);
    expect(await store.redo()).toBe(false);
    expect(events).toHaveLength(0);
  });
});

describe("isDirty", () => {
  const textEdit = (slideId: string, text: string): EditOperation => ({
    type: "setTextRun",
    slideId,
    nodeId: "2",
    paragraphIndex: 0,
    runIndex: 0,
    text,
  });

  it("starts clean and tracks edits against the last save", async () => {
    const store = await editableStore();
    const slideId = store.getState().presentation!.slides[0].id;
    expect(store.isDirty()).toBe(false);

    await store.edit(textEdit(slideId, "One"));
    expect(store.isDirty()).toBe(true);

    await store.save();
    expect(store.isDirty()).toBe(false);

    await store.edit(textEdit(slideId, "Two"));
    expect(store.isDirty()).toBe(true);
  });

  it("clears when undoing back to the saved state and re-dirties on redo", async () => {
    const store = await editableStore();
    const slideId = store.getState().presentation!.slides[0].id;

    await store.edit(textEdit(slideId, "One"));
    expect(store.isDirty()).toBe(true);

    store.undo();
    expect(store.isDirty()).toBe(false);

    await store.redo();
    expect(store.isDirty()).toBe(true);
  });

  it("stays dirty when a new edit discards the path back to the save point", async () => {
    const store = await editableStore();
    const slideId = store.getState().presentation!.slides[0].id;

    await store.edit(textEdit(slideId, "One"));
    await store.edit(textEdit(slideId, "Two"));
    await store.save();

    // Back to one edit deep, then branch: the deck now holds different content
    // at the same stack depth as the save, so it must not read as clean.
    store.undo();
    await store.edit(textEdit(slideId, "Three"));
    expect(store.isDirty()).toBe(true);
  });

  it("resets to clean on a fresh load", async () => {
    const store = await editableStore();
    const slideId = store.getState().presentation!.slides[0].id;
    await store.edit(textEdit(slideId, "One"));
    expect(store.isDirty()).toBe(true);

    await store.load(fixture, { readOnly: false, embedFonts: false });
    expect(store.isDirty()).toBe(false);
  });
});

describe("slideChange events from edits", () => {
  it('reports reason "edit" when deleting the active slide moves it', async () => {
    const store = await editableStore();
    // The slides array is mutated in place by edits, so hold the id, not the index.
    const targetId = store.getState().presentation!.slides[1].id;
    store.goTo(targetId);

    const events: SlideChangeEvent[] = [];
    store.on("slideChange", (event) => events.push(event));

    await store.edit({ type: "deleteSlide", slideId: targetId });

    expect(events).toHaveLength(1);
    expect(events[0].reason).toBe("edit");
    expect(events[0].previousSlideId).toBe(targetId);
    expect(events[0].slideId).toBe(store.getState().activeSlideId);
  });

  it("stays quiet for edits that leave the active slide alone", async () => {
    const store = await editableStore();
    const slideId = store.getState().presentation!.slides[0].id;

    const events: SlideChangeEvent[] = [];
    store.on("slideChange", (event) => events.push(event));

    await store.edit({
      type: "setTextRun",
      slideId,
      nodeId: "2",
      paragraphIndex: 0,
      runIndex: 0,
      text: "Edited",
    });

    expect(events).toHaveLength(0);
  });

  it("reports the jump when undo navigates to the slide it touched", async () => {
    const store = await editableStore();
    const restoredId = store.getState().presentation!.slides[2].id;
    await store.edit({ type: "deleteSlide", slideId: restoredId });

    const events: SlideChangeEvent[] = [];
    store.on("slideChange", (event) => events.push(event));

    store.undo();

    expect(events).toHaveLength(1);
    expect(events[0].reason).toBe("edit");
    expect(events[0].slideId).toBe(restoredId);
  });
});

describe("setTextBody via store", () => {
  it("replaces the text body and bumps revision; undo restores", async () => {
    const store = await editableStore();
    const slideId = store.getState().presentation!.slides[0].id;

    await store.edit({
      type: "setTextBody",
      slideId,
      nodeId: "2",
      paragraphs: [
        { sourceParagraphIndex: 0, runs: [{ text: "First line" }] },
        { sourceParagraphIndex: 0, runs: [{ text: "Second line" }] },
      ],
    });

    const shape = store
      .getState()
      .presentation!.slides[0].nodes.find((n) => n.id === "2") as ShapeNodeData;
    expect(shape.textBody?.paragraphs).toHaveLength(2);
    expect(shape.textBody?.paragraphs[0].runs[0].text).toBe("First line");
    expect(shape.textBody?.paragraphs[1].runs[0].text).toBe("Second line");
    expect(store.getSlideRevision(slideId)).toBe(1);

    store.undo();
    expect(slideText(store, 0)).toBe("Slide 1");
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
