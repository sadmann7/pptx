/**
 * Temporary: cost of scrolling a thumbnail list whose previews are all
 * rendered already, which is the case the background pass leaves behind.
 */
import { test } from "@playwright/test";

/** A real 100-slide deck, already long enough that nothing has to be grown. */
const BENCH_DECK = "heavy-100.pptx";
const TARGET_SLIDES = 0;

test.describe("warm thumbnail list", () => {
  test.skip(!process.env.PERF, "perf run only");
  test.setTimeout(300_000);

  test("scrolling up and down once everything is rendered", async ({ page }) => {
    const cdp = await page.context().newCDPSession(page);
    await cdp.send("Performance.enable");
    const metric = async (name: string) => {
      const { metrics } = await cdp.send("Performance.getMetrics");
      return metrics.find((m) => m.name === name)?.value ?? 0;
    };

    page.on("pageerror", (error) => console.log(`  page error: ${error.message}`));
    page.on("console", (message) => {
      if (message.type() === "error") console.log(`  console: ${message.text()}`);
    });

    await page.goto(`/thumbnails.html?file=decks/${BENCH_DECK}&slides=${TARGET_SLIDES}`);
    try {
      await page.waitForFunction(() => window.__thumbnailsReady === true, undefined, {
        timeout: 60_000,
      });
    } catch {
      const state = await page.evaluate(() => ({
        count: window.__slideCount,
        error: window.__renderError,
      }));
      throw new Error(`not ready: ${JSON.stringify(state)}`);
    }

    const cost = await page.evaluate(() => window.__benchRenderModes?.());
    if (!cost) throw new Error("no render cost");
    const stat = (values: number[]) => {
      const sorted = [...values].sort((a, b) => a - b);
      const total = values.reduce((sum, value) => sum + value, 0);
      return `mean ${(total / values.length).toFixed(1)}ms  median ${(sorted[Math.floor(sorted.length / 2)] ?? 0).toFixed(1)}ms  worst ${Math.max(...values).toFixed(1)}ms  deck ${total.toFixed(0)}ms`;
    };
    console.log(`\n  renderSlide      ${stat(cost.slide)}`);
    console.log(`  renderThumbnail  ${stat(cost.thumbnail)}`);

    // Let the background pass finish the whole deck.
    await page.waitForTimeout(8000);

    const result = await page.evaluate(async () => {
      const list = document.getElementById("thumbnail-list");
      if (!list) throw new Error("missing #thumbnail-list");
      const all = Array.from(list.querySelectorAll<HTMLElement>("[aria-hidden='true']"));
      const nextFrame = () => new Promise<number>((r) => requestAnimationFrame((n) => r(n)));

      const pending = all.filter((preview) => preview.dataset.pending !== undefined).length;
      const canvases = list.querySelectorAll("canvas").length;
      const nodes = list.querySelectorAll("*").length;

      const oscillate = async (step: number) => {
        let jank = 0;
        let worst = 0;
        const durations: number[] = [];
        let direction = 1;
        let previous = await nextFrame();
        for (let frame = 0; frame < 150; frame++) {
          list.scrollTop += step * direction;
          if (list.scrollTop <= 0 || list.scrollTop + list.clientHeight >= list.scrollHeight - 1) {
            direction *= -1;
          }
          const now = await nextFrame();
          const duration = now - previous;
          previous = now;
          durations.push(duration);
          if (duration > 32) jank++;
          worst = Math.max(worst, duration);
        }
        durations.sort((a, b) => a - b);
        return {
          jank,
          worst: Math.round(worst),
          median: Math.round(durations[Math.floor(durations.length / 2)] ?? 0),
          p90: Math.round(durations[Math.floor(durations.length * 0.9)] ?? 0),
        };
      };

      // Oscillate inside a window that is already attached, so no preview can
      // be mounted during the measured frames and what is left is the cost of
      // scrolling the content itself.
      const pitch = (all[1]?.offsetTop ?? 0) - (all[0]?.offsetTop ?? 0);
      const local = async () => {
        const home = list.scrollTop;
        let jank = 0;
        let worst = 0;
        let previous = await nextFrame();
        let direction = 1;
        for (let frame = 0; frame < 150; frame++) {
          list.scrollTop += 300 * direction;
          if (Math.abs(list.scrollTop - home) > pitch * 3) direction *= -1;
          const now = await nextFrame();
          const duration = now - previous;
          previous = now;
          if (duration > 32) jank++;
          worst = Math.max(worst, duration);
        }
        return { jank, worst: Math.round(worst) };
      };

      const snapshot = () => ({
        attached: all.filter((preview) => preview.dataset.pending === undefined).length,
        nodes: list.querySelectorAll("*").length,
      });

      const localResult = await local();
      const afterLocal = snapshot();
      const slowResult = await oscillate(300);
      const afterSlow = snapshot();
      const fastResult = await oscillate(900);
      const afterFast = snapshot();

      return {
        pending,
        canvases,
        nodes,
        total: all.length,
        local: localResult,
        slow: slowResult,
        fast: fastResult,
        afterLocal,
        afterSlow,
        afterFast,
      };
    });

    console.log(
      [
        "",
        `  ${result.total} previews, ${result.pending} still pending, ${result.nodes} nodes, ${result.canvases} chart canvases`,
        `  heap ${((await metric("JSHeapUsedSize")) / 1e6).toFixed(1)}MB`,
        `  attached window  jank ${result.local.jank}/150  worst ${result.local.worst}ms   after: ${result.afterLocal.attached} attached, ${result.afterLocal.nodes} nodes`,
        `  300px/frame  jank ${result.slow.jank}/150  median ${result.slow.median}ms  p90 ${result.slow.p90}ms  worst ${result.slow.worst}ms   after: ${result.afterSlow.attached} attached, ${result.afterSlow.nodes} nodes`,
        `  900px/frame  jank ${result.fast.jank}/150  median ${result.fast.median}ms  p90 ${result.fast.p90}ms  worst ${result.fast.worst}ms   after: ${result.afterFast.attached} attached, ${result.afterFast.nodes} nodes`,
        "",
      ].join("\n"),
    );
  });
});
