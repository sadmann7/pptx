import { describe, expect, it } from "vitest";

import { computePanelScale } from "../../utils/preview-scale";

describe("computePanelScale", () => {
  it("scales from the element's intrinsic size when available", () => {
    const result = computePanelScale({
      panelWidth: 640,
      elementWidth: 1280,
      elementHeight: 720,
      fallbackWidth: 960,
      fallbackHeight: 540,
    });
    expect(result).toEqual({ scale: 0.5, scaledHeight: 360 });
  });

  it("falls back to the fallback dimensions when element size is missing", () => {
    const result = computePanelScale({
      panelWidth: 480,
      elementWidth: null,
      elementHeight: null,
      fallbackWidth: 960,
      fallbackHeight: 540,
    });
    expect(result).toEqual({ scale: 0.5, scaledHeight: 270 });
  });

  it("falls back per-dimension when one element dimension is unusable", () => {
    const result = computePanelScale({
      panelWidth: 500,
      elementWidth: 1000,
      elementHeight: 0, // non-positive → fallback height
      fallbackWidth: 2000,
      fallbackHeight: 800,
    });
    expect(result).toEqual({ scale: 0.5, scaledHeight: 400 });
  });

  it("returns null for non-positive or non-finite panel widths", () => {
    const base = { fallbackWidth: 100, fallbackHeight: 100 };
    expect(computePanelScale({ panelWidth: 0, ...base })).toBeNull();
    expect(computePanelScale({ panelWidth: -10, ...base })).toBeNull();
    expect(computePanelScale({ panelWidth: Number.NaN, ...base })).toBeNull();
    expect(computePanelScale({ panelWidth: Number.POSITIVE_INFINITY, ...base })).toBeNull();
  });

  it("returns null when no usable base dimensions exist", () => {
    expect(
      computePanelScale({
        panelWidth: 100,
        fallbackWidth: 0,
        fallbackHeight: 100,
      }),
    ).toBeNull();
    expect(
      computePanelScale({
        panelWidth: 100,
        fallbackWidth: 100,
        fallbackHeight: Number.NaN,
      }),
    ).toBeNull();
  });
});
