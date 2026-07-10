import { beforeAll, describe, expect, it } from "vitest";

import { createPresentationStore } from "../store";
import { buildMinimalPptx, FIXTURE_SLIDE_COUNT } from "./minimal-pptx";

let fixture: ArrayBuffer;

beforeAll(async () => {
  fixture = await buildMinimalPptx();
});

async function loadedStore() {
  const store = createPresentationStore();
  await store.load(fixture);
  return store;
}

describe("initial state", () => {
  it("starts idle with no presentation", () => {
    const store = createPresentationStore();
    const state = store.getState();
    expect(state.status).toBe("idle");
    expect(state.presentation).toBeNull();
    expect(state.activeSlideId).toBeNull();
    expect(state.zoom).toBe(1);
    expect(store.getActiveSlideIndex()).toBe(-1);
    expect(store.canGoNext()).toBe(false);
    expect(store.canGoPrev()).toBe(false);
  });
});

describe("load", () => {
  it("parses a real pptx and lands on the first slide", async () => {
    const store = await loadedStore();
    const state = store.getState();
    expect(state.status).toBe("ready");
    expect(state.progress).toBe(100);
    expect(state.presentation?.slides).toHaveLength(FIXTURE_SLIDE_COUNT);
    expect(store.getActiveSlideIndex()).toBe(0);
    // 12192000x6858000 EMU = 1280x720 px at 96 DPI.
    expect(state.presentation?.width).toBeCloseTo(1280, 6);
    expect(state.presentation?.height).toBeCloseTo(720, 6);
  });

  it("honors a numeric defaultSlideIndex", async () => {
    const store = createPresentationStore();
    await store.load(fixture, { defaultSlideIndex: 2 });
    expect(store.getActiveSlideIndex()).toBe(2);
  });

  it("honors a defaultSlideIndex resolver and clamps out-of-range values", async () => {
    const store = createPresentationStore();
    await store.load(fixture, { defaultSlideIndex: (slides) => slides.length - 1 });
    expect(store.getActiveSlideIndex()).toBe(FIXTURE_SLIDE_COUNT - 1);

    await store.load(fixture, { defaultSlideIndex: 99 });
    expect(store.getActiveSlideIndex()).toBe(FIXTURE_SLIDE_COUNT - 1);
  });

  it("sets error state when the input is not a valid pptx", async () => {
    const store = createPresentationStore();
    await expect(store.load(new ArrayBuffer(16))).rejects.toThrow();
    const state = store.getState();
    expect(state.status).toBe("error");
    expect(state.error).toBeInstanceOf(Error);
  });

  it("rejects a superseded load with AbortError and keeps the newest result", async () => {
    const store = createPresentationStore();
    const first = store.load(fixture);
    const second = store.load(fixture, { defaultSlideIndex: 1 });

    await expect(first).rejects.toMatchObject({ name: "AbortError" });
    await second;
    expect(store.getState().status).toBe("ready");
    expect(store.getActiveSlideIndex()).toBe(1);
  });
});

describe("lazy slides", () => {
  it("parses only the active slide during load by default", async () => {
    const store = createPresentationStore();
    await store.load(fixture);
    const slides = store.getState().presentation!.slides;

    expect(slides[0].nodesMaterialized).toBe(true);
    expect(slides[0].nodes.length).toBeGreaterThan(0);
    expect(slides[1].nodesMaterialized).toBe(false);
    expect(slides[1].nodes).toHaveLength(0);
  });

  it("materializes the start slide when defaultSlideIndex targets it", async () => {
    const store = createPresentationStore();
    await store.load(fixture, { defaultSlideIndex: 2 });
    const slides = store.getState().presentation!.slides;

    expect(slides[2].nodesMaterialized).toBe(true);
    expect(slides[0].nodesMaterialized).toBe(false);
  });

  it("materializes the target slide before navigation is observable", async () => {
    const store = createPresentationStore();
    await store.load(fixture);

    let nodesAtNotify = -1;
    store.subscribe(() => {
      nodesAtNotify = store.getActiveSlide()?.nodes.length ?? -1;
    });

    store.next();
    const slides = store.getState().presentation!.slides;
    expect(slides[1].nodesMaterialized).toBe(true);
    // Subscribers never see an active slide with unparsed nodes.
    expect(nodesAtNotify).toBeGreaterThan(0);
  });

  it("parses all slides eagerly with lazy: false", async () => {
    const store = createPresentationStore();
    await store.load(fixture, { lazy: false });
    for (const slide of store.getState().presentation!.slides) {
      expect(slide.nodesMaterialized).toBe(true);
      expect(slide.nodes.length).toBeGreaterThan(0);
    }
  });
});

describe("navigation", () => {
  it("navigates with next/prev and reports canGoNext/canGoPrev", async () => {
    const store = await loadedStore();
    expect(store.canGoPrev()).toBe(false);
    expect(store.canGoNext()).toBe(true);

    store.next();
    expect(store.getActiveSlideIndex()).toBe(1);
    expect(store.canGoPrev()).toBe(true);

    store.next();
    expect(store.getActiveSlideIndex()).toBe(2);
    expect(store.canGoNext()).toBe(false);

    store.next(); // no-op on last slide
    expect(store.getActiveSlideIndex()).toBe(2);

    store.prev();
    expect(store.getActiveSlideIndex()).toBe(1);
  });

  it("goToIndex clamps to the valid range", async () => {
    const store = await loadedStore();
    store.goToIndex(99);
    expect(store.getActiveSlideIndex()).toBe(FIXTURE_SLIDE_COUNT - 1);
    store.goToIndex(-5);
    expect(store.getActiveSlideIndex()).toBe(0);
  });

  it("goTo navigates by stable id and ignores unknown ids", async () => {
    const store = await loadedStore();
    const slides = store.getState().presentation!.slides;
    store.goTo(slides[2].id);
    expect(store.getActiveSlideIndex()).toBe(2);
    expect(store.getActiveSlide()?.id).toBe(slides[2].id);

    store.goTo("ppt/slides/nope.xml");
    expect(store.getActiveSlideIndex()).toBe(2);
  });
});

describe("zoom", () => {
  it("clamps setZoom to [0.1, 4] and ignores non-finite values", async () => {
    const store = await loadedStore();
    store.setZoom(10);
    expect(store.getState().zoom).toBe(4);
    store.setZoom(0.0001);
    expect(store.getState().zoom).toBe(0.1);
    store.setZoom(Number.NaN);
    expect(store.getState().zoom).toBe(0.1);
  });

  it("zoomIn/zoomOut step from the current zoom", async () => {
    const store = await loadedStore();
    store.zoomIn();
    expect(store.getState().zoom).toBe(1.25);
    store.zoomOut(0.5);
    expect(store.getState().zoom).toBe(0.75);
  });

  it("fitTo picks the zoom that fits the slide into the container", async () => {
    const store = await loadedStore();
    // Presentation is 1280x720; a 640x360 container fits at exactly 0.5.
    store.fitTo(640, 360);
    expect(store.getState().zoom).toBeCloseTo(0.5, 9);
    // Padding shrinks the available box.
    store.fitTo(660, 380, 10);
    expect(store.getState().zoom).toBeCloseTo(0.5, 9);
  });

  it("fitTo supports per-side padding objects", async () => {
    const store = await loadedStore();
    // 1280x720 slide; horizontal padding is the limiting axis.
    store.fitTo(680, 400, { left: 80, right: 80 });
    expect(store.getState().zoom).toBeCloseTo(0.40625, 9);
    // Vertical padding is the limiting axis.
    store.fitTo(1280, 400, { top: 40, bottom: 40 });
    expect(store.getState().zoom).toBeCloseTo(0.4444444444, 9);
  });
});

describe("reset and subscriptions", () => {
  it("reset returns to the initial idle state", async () => {
    const store = await loadedStore();
    store.reset();
    const state = store.getState();
    expect(state.status).toBe("idle");
    expect(state.presentation).toBeNull();
    expect(state.activeSlideId).toBeNull();
  });

  it("notifies subscribers on change and stops after unsubscribe", async () => {
    const store = await loadedStore();
    let notifications = 0;
    const unsubscribe = store.subscribe(() => {
      notifications++;
    });

    store.next();
    expect(notifications).toBe(1);

    // No state change → no notification.
    store.setZoom(store.getState().zoom);
    expect(notifications).toBe(1);

    unsubscribe();
    store.next();
    expect(notifications).toBe(1);
  });
});
