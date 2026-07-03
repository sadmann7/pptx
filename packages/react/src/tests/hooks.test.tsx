import * as React from "react";

import { act, render } from "@testing-library/react";
import { beforeAll, describe, expect, it } from "vitest";

import type { UsePresentationResult, UseSlideResult, UseZoomResult } from "../context";
import { usePresentation, useSlide, useZoom } from "../context";
import { Presentation } from "../index";
import type { PresentationStore } from "../store";
import { createPresentationStore } from "../store";
import { buildMinimalPptx, FIXTURE_SLIDE_COUNT } from "./minimal-pptx";

let fixture: ArrayBuffer;

beforeAll(async () => {
  fixture = await buildMinimalPptx();
});

/** Renders probes for all three hooks inside a Provider with a loaded store. */
async function renderHooks(): Promise<{
  store: PresentationStore;
  presentation: () => UsePresentationResult;
  slide: () => UseSlideResult;
  zoom: () => UseZoomResult;
}> {
  const store = createPresentationStore();
  await store.load(fixture);

  const latest: {
    presentation?: UsePresentationResult;
    slide?: UseSlideResult;
    zoom?: UseZoomResult;
  } = {};

  function Probe() {
    latest.presentation = usePresentation();
    latest.slide = useSlide();
    latest.zoom = useZoom();
    return null;
  }

  render(
    <Presentation.Provider store={store}>
      <Probe />
    </Presentation.Provider>,
  );

  return {
    store,
    presentation: () => latest.presentation!,
    slide: () => latest.slide!,
    zoom: () => latest.zoom!,
  };
}

describe("usePresentation", () => {
  it("exposes status, progress, and parsed data", async () => {
    const { presentation } = await renderHooks();
    expect(presentation().status).toBe("ready");
    expect(presentation().progress).toBe(100);
    expect(presentation().error).toBeNull();
    expect(presentation().presentation?.slides).toHaveLength(FIXTURE_SLIDE_COUNT);
  });

  it("throws when used outside a Presentation tree", () => {
    function Naked() {
      usePresentation();
      return null;
    }
    expect(() => render(<Naked />)).toThrow(/must be used inside/);
  });
});

describe("useSlide", () => {
  it("reports slide identity, position, and boundary flags", async () => {
    const { slide } = await renderHooks();
    expect(slide().index).toBe(0);
    expect(slide().total).toBe(FIXTURE_SLIDE_COUNT);
    expect(slide().isFirst).toBe(true);
    expect(slide().isLast).toBe(false);
    expect(slide().slideId).toBe(slide().slide?.id);
  });

  it("re-renders on navigation via hook actions", async () => {
    const { slide } = await renderHooks();

    act(() => slide().next());
    expect(slide().index).toBe(1);
    expect(slide().isFirst).toBe(false);

    act(() => slide().goToIndex(FIXTURE_SLIDE_COUNT - 1));
    expect(slide().isLast).toBe(true);

    act(() => slide().prev());
    expect(slide().index).toBe(FIXTURE_SLIDE_COUNT - 2);

    const targetId = slide().slide!.id;
    act(() => slide().goTo(targetId));
    expect(slide().slideId).toBe(targetId);
  });

  it("reflects store-driven navigation", async () => {
    const { store, slide } = await renderHooks();
    act(() => store.goToIndex(2));
    expect(slide().index).toBe(2);
  });
});

describe("useZoom", () => {
  it("exposes zoom state and actions with clamping", async () => {
    const { zoom } = await renderHooks();
    expect(zoom().zoom).toBe(1);

    act(() => zoom().zoomIn());
    expect(zoom().zoom).toBe(1.25);

    act(() => zoom().zoomOut(0.5));
    expect(zoom().zoom).toBe(0.75);

    act(() => zoom().setZoom(99));
    expect(zoom().zoom).toBe(4);
  });

  it("fitTo computes the containing zoom", async () => {
    const { zoom } = await renderHooks();
    // Fixture deck is 1280x720.
    act(() => zoom().fitTo(640, 360, 0));
    expect(zoom().zoom).toBeCloseTo(0.5, 9);
  });
});
