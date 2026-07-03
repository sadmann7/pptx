import { describe, expect, it } from "vitest";

import type { ServerPerSlideMetrics, SlideVisualMetricFields } from "../utils/e2e-compare";
import {
  mergeServerMetricsIntoSlides,
  resolveComparablePdfPages,
  resolveComparePanelState,
  resolveCompareSlideCounts,
} from "../utils/e2e-compare";

describe("resolveComparablePdfPages", () => {
  it("maps visible slides to sequential pdf pages", () => {
    expect(resolveComparablePdfPages([{}, {}, {}], 3)).toEqual([0, 1, 2]);
  });

  it("skips hidden slides so later slides compare against earlier pdf pages", () => {
    // PowerPoint PDF export omits hidden slides.
    expect(resolveComparablePdfPages([{}, { hidden: true }, {}], 2)).toEqual([0, null, 1]);
  });

  it("returns null once pdf pages run out", () => {
    expect(resolveComparablePdfPages([{}, {}, {}], 2)).toEqual([0, 1, null]);
  });

  it("normalizes bogus page counts to zero", () => {
    expect(resolveComparablePdfPages([{}, {}], Number.NaN)).toEqual([null, null]);
    expect(resolveComparablePdfPages([{}], -3)).toEqual([null]);
    expect(resolveComparablePdfPages([{}], 1.9)).toEqual([0]);
  });
});

describe("resolveCompareSlideCounts", () => {
  it("shows all pptx slides but only scores slides with matching pdf pages", () => {
    expect(resolveCompareSlideCounts(5, 3)).toEqual({
      displaySlideCount: 5,
      comparableSlideCount: 3,
    });
    expect(resolveCompareSlideCounts(2, 9)).toEqual({
      displaySlideCount: 2,
      comparableSlideCount: 2,
    });
  });

  it("normalizes non-finite and negative counts", () => {
    expect(resolveCompareSlideCounts(Number.NaN, 4)).toEqual({
      displaySlideCount: 0,
      comparableSlideCount: 0,
    });
    expect(resolveCompareSlideCounts(3, -1)).toEqual({
      displaySlideCount: 3,
      comparableSlideCount: 0,
    });
  });
});

describe("resolveComparePanelState", () => {
  it("triple mode shows truth+render and diff only when available", () => {
    expect(resolveComparePanelState("triple", true, false)).toEqual({
      truth: true,
      render: true,
      diff: true,
      compact: false,
      expanded: false,
      fallback: false,
    });
    expect(resolveComparePanelState("triple", false, true).diff).toBe(false);
  });

  it("side-by-side never shows the diff panel", () => {
    const state = resolveComparePanelState("side-by-side", true, false);
    expect(state).toEqual({
      truth: true,
      render: true,
      diff: false,
      compact: false,
      expanded: false,
      fallback: false,
    });
  });

  it("diff-first without a diff falls back to side-by-side with fallback flag", () => {
    const state = resolveComparePanelState("diff-first", false, false);
    expect(state.fallback).toBe(true);
    expect(state.diff).toBe(false);
    expect(state.truth).toBe(true);
  });

  it("diff-first with a diff is compact unless expanded", () => {
    expect(resolveComparePanelState("diff-first", true, false)).toEqual({
      truth: false,
      render: false,
      diff: true,
      compact: true,
      expanded: false,
      fallback: false,
    });
    const expanded = resolveComparePanelState("diff-first", true, true);
    expect(expanded.expanded).toBe(true);
    expect(expanded.truth).toBe(true);
    expect(expanded.compact).toBe(false);
  });
});

describe("mergeServerMetricsIntoSlides", () => {
  const nullMetrics: SlideVisualMetricFields = {
    ssim: null,
    mae: null,
    fgIou: null,
    fgIouTolerant: null,
    chamferScore: null,
    colorHistCorr: null,
    needsReview: null,
    hasDiff: false,
  };

  function slide(index: number, hasComparablePdf = true) {
    return { index, hasComparablePdf, ...nullMetrics };
  }

  it("merges numeric metrics and flags a diff", () => {
    const metrics: ServerPerSlideMetrics[] = [
      { slideIdx: 0, ssim: 0.98, mae: 0.01, fgIou: 0.9, needsReview: false },
    ];
    const [merged] = mergeServerMetricsIntoSlides([slide(0)], metrics);
    expect(merged.ssim).toBe(0.98);
    expect(merged.mae).toBe(0.01);
    expect(merged.fgIou).toBe(0.9);
    expect(merged.fgIouTolerant).toBeNull();
    expect(merged.needsReview).toBe(false);
    expect(merged.hasDiff).toBe(true);
  });

  it("clears metrics for hidden slides and removes pdf comparability", () => {
    const [merged] = mergeServerMetricsIntoSlides(
      [{ ...slide(0), ssim: 0.5, hasDiff: true }],
      [{ slideIdx: 0, hidden: true }],
    );
    expect(merged.hasComparablePdf).toBe(false);
    expect(merged.ssim).toBeNull();
    expect(merged.hasDiff).toBe(false);
  });

  it("nulls metrics when the slide has no comparable pdf or no server entry", () => {
    const [noPdf, noEntry] = mergeServerMetricsIntoSlides(
      [slide(0, false), slide(1)],
      [{ slideIdx: 0, ssim: 0.9 }],
    );
    expect(noPdf.ssim).toBeNull();
    expect(noPdf.hasDiff).toBe(false);
    expect(noEntry.ssim).toBeNull();
    expect(noEntry.hasDiff).toBe(false);
  });

  it("handles a null metrics payload", () => {
    const [merged] = mergeServerMetricsIntoSlides([slide(0)], null);
    expect(merged.ssim).toBeNull();
    expect(merged.hasDiff).toBe(false);
  });
});
