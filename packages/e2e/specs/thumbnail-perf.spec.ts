import { type Page, test } from "@playwright/test";
/**
 * Thumbnail-list scroll performance.
 *
 * Skipped unless PERF=1 (`pnpm test:perf`): it scrolls the real thumbnail list
 * and reports numbers rather than asserting a baseline, because timings are
 * machine-dependent.
 *
 * Two things get measured:
 *
 *   1. per-slide render cost, `renderSlide` against `renderThumbnail`, which is
 *      what decides whether a preview can be filled inside one frame's budget;
 *   2. skeleton exposure while scrolling: for every animation frame of a sweep,
 *      how many previews inside the viewport are still `[data-pending]`, plus
 *      how long the list needs to settle once the scrolling stops.
 *
 * The second number is the one that matches the complaint ("a bunch of skeleton
 * slides as we scroll until the scroll settles"). Each speed runs on a freshly
 * loaded page so an earlier sweep's cached thumbnails cannot flatter a later
 * one.
 */
import { copyFileSync, mkdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const HERE = dirname(fileURLToPath(import.meta.url));
const DECKS_DIR = join(HERE, "..", "decks");
/** A committed real-world deck; the perf picture only means something on real content. */
const SOURCE_DECK = join(
  HERE,
  "..",
  "..",
  "..",
  "apps",
  "video",
  "public",
  "tiny-adventure-club.pptx",
);
const BENCH_DECK = "bench-thumbnails.pptx";
/**
 * Enough slides that the list scrolls for a long time and, importantly, well
 * past both the observer's runway and the list's retention limit, so the sweep
 * cannot be flattered by a deck that fits entirely in either.
 */
const TARGET_SLIDES = 200;

/**
 * Scroll speeds in px per animation frame: a deliberate drag, a fast flick, and
 * a scrollbar yank. At 60Hz these are roughly 7, 18 and 36 thousand px/s.
 */
const SCROLL_STEPS = [120, 300, 600];
const SCROLL_FRAMES = 120;
/**
 * Sweeps per speed. A single sweep's jank count swings by a factor of two or
 * more between otherwise identical runs, which is enough to read a change into
 * noise, so each speed is measured several times and reported as a median.
 */
const SWEEP_REPEATS = 3;

interface SweepResult {
  frames: number;
  pendingFrames: number;
  worstPending: number;
  meanPending: number;
  visiblePreviews: number;
  settleMs: number;
  jankFrames: number;
  renderedPreviews: number;
  totalPreviews: number;
}

async function openList(page: Page): Promise<void> {
  await page.goto(`/thumbnails.html?file=decks/${BENCH_DECK}&slides=${TARGET_SLIDES}`);
  await page.waitForFunction(() => window.__thumbnailsReady === true, undefined, {
    timeout: 120_000,
  });
}

function sweep(page: Page, step: number): Promise<SweepResult> {
  return page.evaluate<SweepResult, { step: number; frames: number }>(
    async ({ step, frames }) => {
      const list = document.getElementById("thumbnail-list");
      if (!list) throw new Error("missing #thumbnail-list");

      const previews = () => Array.from(list.querySelectorAll<HTMLElement>("[aria-hidden='true']"));

      const all = previews();
      const first = all[0];
      const second = all[1];
      if (!first || !second) throw new Error("expected at least two previews");

      // Items are uniform, so the visible range comes out of arithmetic. Reading
      // a rect per preview per frame would be measurement that causes the jank
      // it is trying to measure.
      const pitch = second.offsetTop - first.offsetTop;
      const firstOffset = first.offsetTop;
      const itemHeight = first.offsetHeight;

      /**
       * Indices of the previews whose box overlaps the viewport, inclusive.
       * Preview k spans [firstOffset + k*pitch, +itemHeight] in scroll
       * coordinates, and the viewport spans [scrollTop, +clientHeight].
       */
      const visibleRange = (): [number, number] => {
        const top = list.scrollTop - firstOffset;
        const start = Math.max(0, Math.floor((top - itemHeight) / pitch) + 1);
        const end = Math.min(all.length - 1, Math.floor((top + list.clientHeight - 1) / pitch));
        return [start, end];
      };

      const visiblePreviews = () => {
        const [start, end] = visibleRange();
        return end - start + 1;
      };

      const pendingCount = () => {
        const [start, end] = visibleRange();
        let pending = 0;
        for (let index = start; index <= end; index++) {
          if (all[index]?.dataset.pending !== undefined) pending++;
        }
        return pending;
      };

      const nextFrame = () =>
        new Promise<number>((resolve) => requestAnimationFrame((now) => resolve(now)));

      list.scrollTop = 0;
      // Let the initial viewport fill before measuring, so the sweep reports
      // what scrolling costs rather than what start-up costs.
      for (let i = 0; i < 60; i++) await nextFrame();

      // Mid-scroll a partial preview shows at each edge, so the count runs one
      // higher than at rest; the worst pending count is reported against this.
      let mostVisible = visiblePreviews();
      let pendingFrames = 0;
      let worstPending = 0;
      let totalPending = 0;
      let measured = 0;
      let jankFrames = 0;
      let previousFrameTime = await nextFrame();

      for (let frame = 0; frame < frames; frame++) {
        list.scrollTop += step;
        const frameTime = await nextFrame();
        // Two missed vsyncs at 60Hz; long enough that scrolling visibly hitches.
        if (frameTime - previousFrameTime > 32) jankFrames++;
        previousFrameTime = frameTime;

        const pending = pendingCount();
        mostVisible = Math.max(mostVisible, visiblePreviews());
        measured++;
        totalPending += pending;
        worstPending = Math.max(worstPending, pending);
        if (pending > 0) pendingFrames++;
        if (list.scrollTop + list.clientHeight >= list.scrollHeight) break;
      }

      const settleStart = performance.now();
      while (pendingCount() > 0 && performance.now() - settleStart < 10_000) {
        await nextFrame();
      }
      const settleMs = performance.now() - settleStart;

      return {
        frames: measured,
        pendingFrames,
        worstPending,
        meanPending: measured > 0 ? totalPending / measured : 0,
        visiblePreviews: mostVisible,
        settleMs,
        jankFrames,
        renderedPreviews: all.filter((preview) => preview.dataset.pending === undefined).length,
        totalPreviews: all.length,
      };
    },
    { step, frames: SCROLL_FRAMES },
  );
}

test.describe("thumbnail list scroll performance", () => {
  test.skip(!process.env.PERF, "perf run only: PERF=1 pnpm test:perf");
  test.setTimeout(300_000);

  // decks/ is the directory the harness server exposes, so the deck under test
  // is copied there rather than served from another package.
  test.beforeAll(() => {
    mkdirSync(DECKS_DIR, { recursive: true });
    copyFileSync(SOURCE_DECK, join(DECKS_DIR, BENCH_DECK));
  });

  test("reports render cost and skeleton exposure while scrolling", async ({ page }) => {
    await openList(page);
    const slideCount = await page.evaluate(() => window.__slideCount ?? 0);
    const renderCost = await page.evaluate(() => window.__benchRenderModes?.());
    if (!renderCost) throw new Error("harness did not expose __benchRenderModes");

    const total = (values: number[]) => values.reduce((sum, value) => sum + value, 0);
    const worst = (values: number[]) => Math.max(...values);
    const mean = (values: number[]) => (values.length > 0 ? total(values) / values.length : 0);
    const median = (values: number[]) =>
      [...values].sort((a, b) => a - b)[Math.floor(values.length / 2)] ?? 0;

    const sweeps: Array<{ step: number; runs: SweepResult[] }> = [];
    for (const step of SCROLL_STEPS) {
      const runs: SweepResult[] = [];
      for (let repeat = 0; repeat < SWEEP_REPEATS; repeat++) {
        // A fresh page each time, so one sweep's warm cache cannot flatter the
        // next.
        await openList(page);
        runs.push(await sweep(page, step));
      }
      sweeps.push({ step, runs });
    }

    const lines = [
      "",
      `deck  ${BENCH_DECK} (${slideCount} slides)`,
      "",
      "per-slide render cost",
      `  renderSlide      mean ${mean(renderCost.slide).toFixed(1)}ms  worst ${worst(renderCost.slide).toFixed(1)}ms  deck ${total(renderCost.slide).toFixed(0)}ms`,
      `  renderThumbnail  mean ${mean(renderCost.thumbnail).toFixed(1)}ms  worst ${worst(renderCost.thumbnail).toFixed(1)}ms  deck ${total(renderCost.thumbnail).toFixed(0)}ms`,
      "",
      `scroll sweeps (median of ${SWEEP_REPEATS}; skeletons counted among previews overlapping the viewport)`,
    ];
    for (const { step, runs } of sweeps) {
      const pick = (read: (result: SweepResult) => number) => median(runs.map(read));
      const last = runs[runs.length - 1];
      if (!last) continue;
      lines.push(
        `  ${String(step).padStart(3)}px/frame  skeleton frames ${pick((r) => r.pendingFrames)}/${last.frames}` +
          `  on screen mean ${pick((r) => r.meanPending).toFixed(1)} worst ${pick((r) => r.worstPending)}/${last.visiblePreviews}` +
          `  settle ${pick((r) => r.settleMs).toFixed(0)}ms  jank ${pick((r) => r.jankFrames)}` +
          ` (${runs.map((r) => r.jankFrames).join("/")})` +
          `  kept ${pick((r) => r.renderedPreviews)}/${last.totalPreviews}`,
      );
    }
    lines.push("");
    console.log(lines.join("\n"));
  });
});
