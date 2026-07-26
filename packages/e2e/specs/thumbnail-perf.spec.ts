import { test } from "@playwright/test";
/**
 * Thumbnail-list scroll performance.
 *
 * Skipped unless PERF=1 (`pnpm test:perf`): it builds a long deck, scrolls the
 * real thumbnail list, and reports numbers rather than asserting a baseline,
 * because timings are machine-dependent.
 *
 * Two things get measured:
 *
 *   1. per-slide render cost, `renderSlide` against `renderThumbnail`, which is
 *      what decides whether a preview can be filled within one frame's budget;
 *   2. skeleton exposure during a fast scroll: for every animation frame of the
 *      sweep, how many previews inside the viewport are still `[data-pending]`,
 *      plus how long the list takes to settle once scrolling stops.
 *
 * The second number is the one that matches the complaint ("a bunch of
 * skeleton slides as we scroll until the scroll settles").
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
/** Enough slides that the list scrolls for a while at a realistic speed. */
const TARGET_SLIDES = 64;

/** Pixels per animation frame, roughly a fast but not absurd flick. */
const SCROLL_STEP_PX = 120;
const SCROLL_FRAMES = 90;

interface SweepResult {
  frames: number;
  pendingFrames: number;
  worstPending: number;
  meanPending: number;
  settleMs: number;
  renderedPreviews: number;
  totalPreviews: number;
}

test.describe("thumbnail list scroll performance", () => {
  test.skip(!process.env.PERF, "perf run only: PERF=1 pnpm test:perf");
  test.setTimeout(180_000);

  // decks/ is the directory the harness server exposes, so the deck under test
  // is copied there rather than served from another package.
  test.beforeAll(() => {
    mkdirSync(DECKS_DIR, { recursive: true });
    copyFileSync(SOURCE_DECK, join(DECKS_DIR, BENCH_DECK));
  });

  test("reports render cost and skeleton exposure while scrolling", async ({ page }) => {
    await page.goto(`/thumbnails.html?file=decks/${BENCH_DECK}&slides=${TARGET_SLIDES}`);
    await page.waitForFunction(() => window.__thumbnailsReady === true, undefined, {
      timeout: 120_000,
    });

    const slideCount = await page.evaluate(() => window.__slideCount ?? 0);

    const renderCost = await page.evaluate(() => window.__benchRenderModes?.());
    if (!renderCost) throw new Error("harness did not expose __benchRenderModes");

    const sweep = await page.evaluate<SweepResult, { step: number; frames: number }>(
      async ({ step, frames }) => {
        const list = document.getElementById("thumbnail-list");
        if (!list) throw new Error("missing #thumbnail-list");

        const previews = () =>
          Array.from(list.querySelectorAll<HTMLElement>("[aria-hidden='true']"));

        /** Previews overlapping the scroller viewport, i.e. what the user sees. */
        const visiblePending = () => {
          const listRect = list.getBoundingClientRect();
          let pending = 0;
          for (const preview of previews()) {
            const rect = preview.getBoundingClientRect();
            if (rect.bottom < listRect.top || rect.top > listRect.bottom) continue;
            if (preview.dataset.pending !== undefined) pending++;
          }
          return pending;
        };

        const nextFrame = () =>
          new Promise<number>((resolve) => requestAnimationFrame((now) => resolve(now)));

        list.scrollTop = 0;
        // Let the initial viewport fill before measuring, so the sweep reports
        // what scrolling costs rather than what start-up costs.
        for (let i = 0; i < 60; i++) await nextFrame();

        let pendingFrames = 0;
        let worstPending = 0;
        let totalPending = 0;
        let measured = 0;

        for (let frame = 0; frame < frames; frame++) {
          list.scrollTop += step;
          await nextFrame();
          const pending = visiblePending();
          measured++;
          totalPending += pending;
          worstPending = Math.max(worstPending, pending);
          if (pending > 0) pendingFrames++;
          if (list.scrollTop + list.clientHeight >= list.scrollHeight) break;
        }

        const settleStart = performance.now();
        while (visiblePending() > 0 && performance.now() - settleStart < 10_000) {
          await nextFrame();
        }
        const settleMs = performance.now() - settleStart;

        const all = previews();
        return {
          frames: measured,
          pendingFrames,
          worstPending,
          meanPending: measured > 0 ? totalPending / measured : 0,
          settleMs,
          renderedPreviews: all.filter((preview) => preview.dataset.pending === undefined).length,
          totalPreviews: all.length,
        };
      },
      { step: SCROLL_STEP_PX, frames: SCROLL_FRAMES },
    );

    const total = (values: number[]) => values.reduce((sum, value) => sum + value, 0);
    const worst = (values: number[]) => Math.max(...values);
    const mean = (values: number[]) => (values.length > 0 ? total(values) / values.length : 0);

    console.log(
      [
        "",
        `deck                  ${BENCH_DECK} (${slideCount} slides)`,
        "",
        "per-slide render cost",
        `  renderSlide         mean ${mean(renderCost.slide).toFixed(1)}ms   worst ${worst(renderCost.slide).toFixed(1)}ms   deck ${total(renderCost.slide).toFixed(0)}ms`,
        `  renderThumbnail     mean ${mean(renderCost.thumbnail).toFixed(1)}ms   worst ${worst(renderCost.thumbnail).toFixed(1)}ms   deck ${total(renderCost.thumbnail).toFixed(0)}ms`,
        "",
        `fast scroll (${SCROLL_STEP_PX}px/frame)`,
        `  frames with a visible skeleton   ${sweep.pendingFrames}/${sweep.frames}`,
        `  skeletons on screen              mean ${sweep.meanPending.toFixed(1)}   worst ${sweep.worstPending}`,
        `  settle time after scroll stops   ${sweep.settleMs.toFixed(0)}ms`,
        `  previews rendered at the end     ${sweep.renderedPreviews}/${sweep.totalPreviews}`,
        "",
      ].join("\n"),
    );
  });
});
