import { beforeAll, describe, expect, it } from "vitest";

import type { SlideChangeEvent, ZoomChangeEvent } from "../store";
import { createStore } from "../store";
import { FIXTURE_SLIDE_COUNT } from "./minimal-pptx";
import { loadFixture } from "./test-utils";

let fixture: ArrayBuffer;

beforeAll(async () => {
  fixture = await loadFixture();
});

async function loadedStore() {
  const store = createStore();
  await store.load(fixture);
  return store;
}

describe("initial state", () => {
  it("starts idle with no presentation", () => {
    const store = createStore();
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
    const store = createStore();
    await store.load(fixture, { defaultSlideIndex: 2 });
    expect(store.getActiveSlideIndex()).toBe(2);
  });

  it("honors a defaultSlideIndex resolver and clamps out-of-range values", async () => {
    const store = createStore();
    await store.load(fixture, { defaultSlideIndex: (slides) => slides.length - 1 });
    expect(store.getActiveSlideIndex()).toBe(FIXTURE_SLIDE_COUNT - 1);

    await store.load(fixture, { defaultSlideIndex: 99 });
    expect(store.getActiveSlideIndex()).toBe(FIXTURE_SLIDE_COUNT - 1);
  });

  it("sets error state when the input is not a valid pptx", async () => {
    const store = createStore();
    await expect(store.load(new ArrayBuffer(16))).rejects.toThrow();
    const state = store.getState();
    expect(state.status).toBe("error");
    expect(state.error).toBeInstanceOf(Error);
  });

  it("rejects a superseded load with AbortError and keeps the newest result", async () => {
    const store = createStore();
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
    const store = createStore();
    await store.load(fixture);
    const slides = store.getState().presentation!.slides;

    expect(slides[0].nodesMaterialized).toBe(true);
    expect(slides[0].nodes.length).toBeGreaterThan(0);
    expect(slides[1].nodesMaterialized).toBe(false);
    expect(slides[1].nodes).toHaveLength(0);
  });

  it("materializes the start slide when defaultSlideIndex targets it", async () => {
    const store = createStore();
    await store.load(fixture, { defaultSlideIndex: 2 });
    const slides = store.getState().presentation!.slides;

    expect(slides[2].nodesMaterialized).toBe(true);
    expect(slides[0].nodesMaterialized).toBe(false);
  });

  it("materializes the target slide before navigation is observable", async () => {
    const store = createStore();
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

  it("parses all slides eagerly with lazySlides: false", async () => {
    const store = createStore();
    await store.load(fixture, { lazySlides: false });
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

    // Same zoom, but the requested level moves from "fit" to a number.
    store.setZoom(store.getState().zoom);
    expect(notifications).toBe(2);

    // Now nothing moves at all → no notification.
    store.setZoom(store.getState().zoom);
    expect(notifications).toBe(2);

    unsubscribe();
    store.next();
    expect(notifications).toBe(2);
  });
});

describe("statusChange and zoomChange events", () => {
  it("reports each status transition through a load", async () => {
    const store = createStore();
    const transitions: string[] = [];
    store.on("statusChange", ({ status, previousStatus }) =>
      transitions.push(`${previousStatus}->${status}`),
    );

    await store.load(fixture);
    expect(transitions).toEqual(["idle->loading", "loading->ready"]);

    store.reset();
    expect(transitions.at(-1)).toBe("ready->idle");
  });

  it("reports a failed load as a transition to error", async () => {
    const store = createStore();
    const transitions: string[] = [];
    store.on("statusChange", ({ status }) => transitions.push(status));

    await expect(store.load(new ArrayBuffer(16))).rejects.toThrow();
    expect(transitions).toEqual(["loading", "error"]);
  });

  it("reports zoom changes once per effective change", async () => {
    const store = await loadedStore();
    const zooms: number[] = [];
    store.on("zoomChange", ({ zoom }) => zooms.push(zoom));

    store.zoomIn();
    store.setZoom(1.25); // No-op: already there.
    store.fitTo(640, 360);
    expect(zooms).toEqual([1.25, 0.5]);
  });

  it("reports the previous zoom alongside the new one", async () => {
    const store = await loadedStore();
    let event: ZoomChangeEvent | null = null;
    store.on("zoomChange", (payload) => {
      event = payload;
    });

    store.setZoom(2);
    expect(event).toEqual({ zoom: 2, previousZoom: 1, reason: "zoom" });
  });

  it("reports what produced each zoom change", async () => {
    const store = await loadedStore();
    const reasons: string[] = [];
    store.on("zoomChange", ({ reason }) => reasons.push(reason));

    store.setZoom(2);
    store.zoomOut(0.5);
    store.fitTo(640, 360);
    store.reset();

    expect(reasons).toEqual(["zoom", "zoom", "fit", "reset"]);
  });
});

describe("zoomLevel", () => {
  it("starts fitted and pins the level on an explicit zoom", async () => {
    const store = await loadedStore();
    expect(store.getState().zoomLevel).toBe("fit");

    store.setZoom(2);
    expect(store.getState().zoomLevel).toBe(2);

    store.fitTo(640, 360);
    expect(store.getState().zoomLevel).toBe("fit");
  });

  it("pins the level even when the zoom does not move", async () => {
    const store = await loadedStore();
    store.fitTo(640, 360);
    const fittedZoom = store.getState().zoom;

    // Picking the percentage the fit had already produced still pins it,
    // otherwise a resize would silently take it back.
    store.setZoom(fittedZoom);
    expect(store.getState().zoom).toBe(fittedZoom);
    expect(store.getState().zoomLevel).toBe(fittedZoom);
  });

  it("records a fit request without a container to resolve it", async () => {
    const store = await loadedStore();
    store.setZoom(3);

    // Nothing knows the container size here, so the zoom holds until a
    // viewport (or a `fitTo` call) resolves the request.
    store.setZoom("fit");
    expect(store.getState().zoom).toBe(3);
    expect(store.getState().zoomLevel).toBe("fit");
  });

  it("clamps and pins defaultZoom", async () => {
    const store = createStore();
    await store.load(await loadFixture(), { defaultZoom: 0.5, embedFonts: false });

    expect(store.getState().zoom).toBe(0.5);
    expect(store.getState().zoomLevel).toBe(0.5);
  });

  it("stays fitted when no defaultZoom is given", async () => {
    const store = await loadedStore();

    expect(store.getState().zoom).toBe(1);
    expect(store.getState().zoomLevel).toBe("fit");
  });
});

describe("slideChange events", () => {
  it("reports the new slide, the previous one, and why it changed", async () => {
    const store = await loadedStore();
    const slides = store.getState().presentation!.slides;
    const events: SlideChangeEvent[] = [];
    store.on("slideChange", (event) => events.push(event));

    store.next();
    expect(events).toEqual([
      {
        slideId: slides[1].id,
        index: 1,
        previousSlideId: slides[0].id,
        reason: "navigate",
      },
    ]);
  });

  it("fires once per navigation regardless of which action triggered it", async () => {
    const store = await loadedStore();
    const events: SlideChangeEvent[] = [];
    store.on("slideChange", (event) => events.push(event));

    store.next();
    store.prev();
    store.goToIndex(2);
    store.goTo(store.getState().presentation!.slides[0].id);

    expect(events.map((event) => event.index)).toEqual([1, 0, 2, 0]);
    expect(events.every((event) => event.reason === "navigate")).toBe(true);
  });

  it("stays quiet when navigation is a no-op", async () => {
    const store = await loadedStore();
    const events: SlideChangeEvent[] = [];
    store.on("slideChange", (event) => events.push(event));

    // Already on the first slide.
    store.prev();
    store.goToIndex(0);
    store.goTo(store.getState().activeSlideId!);
    store.goTo("ppt/slides/nope.xml");

    expect(events).toHaveLength(0);
  });

  it("reports the start slide on load and the cleared slide on reset", async () => {
    const store = createStore();
    const events: SlideChangeEvent[] = [];
    store.on("slideChange", (event) => events.push(event));

    await store.load(fixture, { defaultSlideIndex: 1 });
    expect(events).toHaveLength(1);
    expect(events[0].reason).toBe("load");
    expect(events[0].index).toBe(1);
    expect(events[0].previousSlideId).toBeNull();

    const loadedSlideId = events[0].slideId;
    store.reset();
    expect(events).toHaveLength(2);
    expect(events[1]).toEqual({
      slideId: null,
      index: -1,
      previousSlideId: loadedSlideId,
      reason: "reset",
    });
  });

  it("exposes state that is already current when the handler runs", async () => {
    const store = await loadedStore();
    let seen: { active: string | null; index: number } | null = null;
    store.on("slideChange", (event) => {
      seen = { active: store.getState().activeSlideId, index: store.getActiveSlideIndex() };
      expect(event.slideId).toBe(store.getState().activeSlideId);
    });

    store.goToIndex(2);
    expect(seen).toEqual({ active: store.getState().activeSlideId, index: 2 });
  });

  it("stops delivering after the listener is removed", async () => {
    const store = await loadedStore();
    let count = 0;
    const off = store.on("slideChange", () => {
      count++;
    });

    store.next();
    expect(count).toBe(1);

    off();
    store.next();
    expect(count).toBe(1);
  });
});
