/**
 * Records editor performance on the docs playground (/pg) against a running
 * server, the way tablecn's profile:scroll does for its data grid.
 *
 * Run against a production build on a real display (headed, not headless):
 *   pnpm -F docs build && pnpm -F docs start -p 3003
 *   pnpm bench-deck            # once: recordings/perf/decks/bench-300.pptx
 *   pnpm profile:editor <label> [scenario]
 *
 * Writes a Chrome trace per scenario to recordings/perf/<label>/ (open in the
 * DevTools Performance panel) and prints frame stats. Use a different label per
 * change to compare before and after, and only compare runs made back to back.
 *
 * BASE_URL picks the server (default http://localhost:3003), DECK the deck.
 * Set EXTRA_CSS to inject a stylesheet, or STRIP_CSS to a regex to delete the
 * style rules whose selector matches it, to A/B CSS without a rebuild. Set
 * CPU_THROTTLE (for example 4) to emulate a slower CPU.
 */
import { type CDPSession, chromium, type Page } from "@playwright/test";
import { existsSync } from "node:fs";
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BASE_URL = process.env.BASE_URL ?? "http://localhost:3003";
const PAGE_PATH = process.env.PAGE_PATH ?? "/pg?layout=editing";
const DECK =
  process.env.DECK ?? path.join(HERE, "..", "recordings", "perf", "decks", "bench-300.pptx");
const EDITOR = "[data-slot=presentation-editor]";
const LIST = `${EDITOR} [role=listbox]`;
const SHAPE = `${EDITOR} [data-slot=presentation-slide] [data-pptx-node-id]`;
const TRACE_CATEGORIES = [
  "devtools.timeline",
  "disabled-by-default-devtools.timeline",
  "disabled-by-default-devtools.timeline.frame",
  "blink.user_timing",
  "v8.execute",
  "disabled-by-default-v8.cpu_profiler",
  "latencyInfo",
];

interface Scenario {
  name: string;
  /** Counts blank thumbnails on screen per frame. */
  thumbnails?: boolean;
  run: (page: Page) => Promise<void>;
}

const scenarios: Scenario[] = [
  {
    name: "thumbs-slow-wheel",
    thumbnails: true,
    run: async (page) => {
      await hoverList(page);
      await wheel(page, { deltaY: 40, events: 240, intervalMs: 8 });
    },
  },
  {
    name: "thumbs-fast-fling",
    thumbnails: true,
    run: async (page) => {
      await hoverList(page);
      for (let delta = 600; delta > 20; delta *= 0.97) {
        await page.mouse.wheel(0, delta);
        await page.waitForTimeout(8);
      }
    },
  },
  {
    // Driven by scrollTop so it does not depend on overlay scrollbars
    name: "thumbs-scrollbar-drag",
    thumbnails: true,
    run: (page) =>
      page.evaluate(
        ({ list, durationMs }) =>
          new Promise<void>((resolve) => {
            const scroller = getScroller(document.querySelector<HTMLElement>(list));
            if (!scroller) return resolve();
            const max = scroller.scrollHeight - scroller.clientHeight;
            const start = performance.now();
            const step = (now: number) => {
              const progress = Math.min((now - start) / durationMs, 1);
              scroller.scrollTop = max * progress;
              if (progress < 1) requestAnimationFrame(step);
              else resolve();
            };
            requestAnimationFrame(step);

            function getScroller(element: HTMLElement | null) {
              let current = element;
              while (current && current.scrollHeight <= current.clientHeight) {
                current = current.parentElement;
              }
              return current;
            }
          }),
        { list: LIST, durationMs: 3000 },
      ),
  },
  {
    // Holding ArrowDown on the thumbnails: every step renders the next slide
    name: "slide-step",
    run: async (page) => {
      await page.locator(`${LIST} [role=option]`).first().focus();
      for (let i = 0; i < 40; i++) {
        await page.keyboard.press("ArrowDown");
        await page.waitForTimeout(33);
      }
    },
  },
  {
    // Dragging a shape around the slide: overlay, snapping, live transform
    name: "shape-drag",
    run: async (page) => {
      const box = await page.locator(SHAPE).nth(1).boundingBox();
      if (!box) return;
      const cx = box.x + box.width / 2;
      const cy = box.y + box.height / 2;
      await page.mouse.move(cx, cy);
      await page.mouse.down();
      for (let i = 0; i < 240; i++) {
        const angle = (i / 60) * Math.PI * 2;
        await page.mouse.move(cx + Math.cos(angle) * 120, cy + Math.sin(angle) * 80);
        await page.waitForTimeout(8);
      }
      await page.mouse.up();
      await page.keyboard.press("ControlOrMeta+z");
    },
  },
];

async function hoverList(page: Page) {
  const box = await page.locator(LIST).boundingBox();
  if (box) await page.mouse.move(box.x + box.width / 2, box.y + 200);
}

async function wheel(
  page: Page,
  { deltaY, events, intervalMs }: { deltaY: number; events: number; intervalMs: number },
) {
  for (let i = 0; i < events; i++) {
    await page.mouse.wheel(0, deltaY);
    await page.waitForTimeout(intervalMs);
  }
}

interface FrameStats {
  frames: number;
  fps: number;
  p50: number;
  p95: number;
  p99: number;
  max: number;
  droppedFrames: number;
  longAnimationFrames: number;
  loafScriptMs: number;
  /** Share of frames with a blank (still rendering) thumbnail on screen, in %. */
  blankFramesPct: number;
  mainThreadMs: number;
}

interface TraceEvent {
  name: string;
  ph: string;
  pid: number;
  tid: number;
  dur?: number;
  args?: { name?: string };
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});

async function main() {
  const label = process.argv[2] ?? new Date().toISOString().replace(/[:.]/g, "-");
  const only = process.argv[3];
  const outDir = path.join(HERE, "..", "recordings", "perf", label);
  await mkdir(outDir, { recursive: true });
  if (!existsSync(DECK)) throw new Error(`No deck at ${DECK}; run \`pnpm bench-deck\` first.`);

  const browser = await chromium.launch({ headless: false });
  const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });
  const cdp = await page.context().newCDPSession(page);
  const cpuThrottle = Number(process.env.CPU_THROTTLE ?? 1);
  if (cpuThrottle > 1) {
    await cdp.send("Emulation.setCPUThrottlingRate", { rate: cpuThrottle });
  }

  // tsx keeps function names via an `__name` helper that page.evaluate does not serialize
  await page.addInitScript("globalThis.__name = (fn) => fn");
  await page.goto(`${BASE_URL}${PAGE_PATH}`);
  await page.locator('input[type=file][accept=".pptx"]').setInputFiles(DECK);
  await page
    .locator(`${EDITOR} [data-pptx-selection]`)
    .waitFor({ state: "attached", timeout: 120_000 });
  if (process.env.EXTRA_CSS) await page.addStyleTag({ content: process.env.EXTRA_CSS });
  if (process.env.STRIP_CSS) {
    const stripped = await page.evaluate(stripStyleRules, process.env.STRIP_CSS);
    console.log(`Stripped ${stripped} style rules matching ${process.env.STRIP_CSS}`);
  }
  // Let parsing, fonts, and the first thumbnails settle
  await page.waitForTimeout(3000);

  const vsync = await page.evaluate(measureVsync);
  console.log(`Display refresh: ${Math.round(1000 / vsync)} Hz`);

  const results: Record<string, FrameStats> = {};
  for (const scenario of scenarios) {
    if (only && scenario.name !== only) continue;

    // Frame stats come from an untraced run, since tracing adds overhead to every frame
    await reset(page);
    await page.evaluate(startFrameRecorder, scenario.thumbnails ? LIST : null);
    await scenario.run(page);
    // Let thumbnails catch up inside the measurement
    await page.waitForTimeout(300);
    const recorded = await page.evaluate(stopFrameRecorder);
    const stats = summarize(recorded, vsync);

    await reset(page);
    stats.mainThreadMs = await recordTrace(
      page,
      cdp,
      path.join(outDir, `${scenario.name}.json`),
      async () => {
        await scenario.run(page);
        await page.waitForTimeout(300);
      },
    );
    results[scenario.name] = stats;
  }

  await browser.close();

  console.table(
    Object.fromEntries(
      Object.entries(results).map(([name, stats]) => [
        name,
        Object.fromEntries(Object.entries(stats).map(([key, value]) => [key, round(value)])),
      ]),
    ),
  );
  await writeFile(path.join(outDir, "summary.json"), JSON.stringify(results, null, 2));
  console.log(`Traces and summary written to ${outDir}`);
}

/** Back to the first slide, nothing selected, thumbnails scrolled to the top. */
async function reset(page: Page) {
  await page.keyboard.press("Escape");
  await page.locator(`${LIST} [role=option]`).first().click();
  await page.evaluate((list) => {
    let scroller = document.querySelector<HTMLElement>(list);
    while (scroller && scroller.scrollHeight <= scroller.clientHeight) {
      scroller = scroller.parentElement;
    }
    scroller?.scrollTo({ top: 0, behavior: "instant" });
  }, LIST);
  await page.mouse.move(0, 0);
  await page.waitForTimeout(800);
}

function stripStyleRules(pattern: string) {
  const regex = new RegExp(pattern);
  let count = 0;
  const strip = (rules: CSSRuleList, owner: CSSStyleSheet | CSSGroupingRule) => {
    for (let index = rules.length - 1; index >= 0; index--) {
      const rule = rules[index];
      if (rule instanceof CSSStyleRule && regex.test(rule.selectorText)) {
        owner.deleteRule(index);
        count++;
      } else if (rule instanceof CSSGroupingRule) {
        strip(rule.cssRules, rule);
      }
    }
  };
  for (const sheet of document.styleSheets) {
    try {
      strip(sheet.cssRules, sheet);
    } catch {
      // Cross-origin sheets are not readable
    }
  }
  return count;
}

function measureVsync() {
  return new Promise<number>((resolve) => {
    const deltas: number[] = [];
    let last = performance.now();
    const tick = (now: number) => {
      deltas.push(now - last);
      last = now;
      if (deltas.length < 120) requestAnimationFrame(tick);
      else resolve(deltas.slice(1).sort((a, b) => a - b)[59] ?? 0);
    };
    requestAnimationFrame(tick);
  });
}

function startFrameRecorder(list: string | null) {
  const recorder = {
    deltas: [] as number[],
    blankFrames: 0,
    loafs: [] as { duration: number; scriptMs: number }[],
    running: true,
  };
  const listElement = list ? document.querySelector<HTMLElement>(list) : null;
  let last = performance.now();
  const tick = (now: number) => {
    recorder.deltas.push(now - last);
    last = now;
    if (listElement) {
      const { top, bottom } = listElement.getBoundingClientRect();
      const viewTop = Math.max(top, 0);
      const viewBottom = Math.min(bottom, innerHeight);
      for (const preview of listElement.querySelectorAll("[data-pending]")) {
        const rect = preview.getBoundingClientRect();
        if (rect.bottom > viewTop && rect.top < viewBottom) {
          recorder.blankFrames++;
          break;
        }
      }
    }
    if (recorder.running) requestAnimationFrame(tick);
  };
  requestAnimationFrame(tick);
  const observer = new PerformanceObserver((entries) => {
    for (const entry of entries.getEntries()) {
      const scripts =
        (entry as PerformanceEntry & { scripts?: { duration: number }[] }).scripts ?? [];
      recorder.loafs.push({
        duration: entry.duration,
        scriptMs: scripts.reduce((sum, script) => sum + script.duration, 0),
      });
    }
  });
  try {
    observer.observe({ type: "long-animation-frame", buffered: false });
  } catch {
    // Long animation frames are Chromium only
  }
  Object.assign(window, { __frameRecorder: recorder, __frameObserver: observer });
}

function stopFrameRecorder() {
  const { __frameRecorder: recorder, __frameObserver: observer } = window as unknown as {
    __frameRecorder: {
      deltas: number[];
      blankFrames: number;
      loafs: { duration: number; scriptMs: number }[];
      running: boolean;
    };
    __frameObserver: PerformanceObserver;
  };
  recorder.running = false;
  observer.disconnect();
  // The first delta spans the gap before recording started
  return {
    deltas: recorder.deltas.slice(1),
    blankFrames: recorder.blankFrames,
    loafs: recorder.loafs,
  };
}

function summarize(
  {
    deltas,
    blankFrames,
    loafs,
  }: { deltas: number[]; blankFrames: number; loafs: { duration: number; scriptMs: number }[] },
  vsync: number,
): FrameStats {
  const at = (q: number) => percentile(deltas, q);
  const total = deltas.reduce((sum, delta) => sum + delta, 0);
  return {
    frames: deltas.length,
    fps: total > 0 ? (deltas.length * 1000) / total : 0,
    p50: at(0.5),
    p95: at(0.95),
    p99: at(0.99),
    max: Math.max(0, ...deltas),
    droppedFrames: deltas.reduce(
      (sum, delta) => sum + Math.max(0, Math.round(delta / vsync) - 1),
      0,
    ),
    longAnimationFrames: loafs.length,
    loafScriptMs: loafs.reduce((sum, loaf) => sum + loaf.scriptMs, 0),
    blankFramesPct: deltas.length > 0 ? (blankFrames / (deltas.length + 1)) * 100 : 0,
    // Filled in from the traced run
    mainThreadMs: 0,
  };
}

function percentile(values: number[], q: number) {
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.min(sorted.length - 1, Math.floor(q * sorted.length))] ?? 0;
}

async function recordTrace(page: Page, cdp: CDPSession, outFile: string, run: () => Promise<void>) {
  const events: unknown[] = [];
  const onData = ({ value }: { value: unknown[] }) => events.push(...value);
  cdp.on("Tracing.dataCollected", onData);
  await cdp.send("Tracing.start", {
    traceConfig: { includedCategories: TRACE_CATEGORIES },
    transferMode: "ReportEvents",
  });
  // Starting the CPU profiler stalls the page, so keep that out of the scenario
  await page.waitForTimeout(500);
  await run();
  const complete = new Promise<void>((resolve) =>
    cdp.once("Tracing.tracingComplete", () => resolve()),
  );
  await cdp.send("Tracing.end");
  await complete;
  cdp.off("Tracing.dataCollected", onData);
  await writeFile(outFile, JSON.stringify({ traceEvents: events }));
  return getMainThreadBusyMs(events as TraceEvent[]);
}

/** Total duration of top-level tasks on the page's main thread, which shows headroom even when fps is capped. */
function getMainThreadBusyMs(events: TraceEvent[]) {
  const mainThreads = new Set(
    events
      .filter(
        (event) =>
          event.ph === "M" && event.name === "thread_name" && event.args?.name === "CrRendererMain",
      )
      .map((event) => `${event.pid}:${event.tid}`),
  );
  const busyByThread = new Map<string, number>();
  for (const event of events) {
    if (event.ph !== "X" || event.name !== "RunTask" || !event.dur) continue;
    const thread = `${event.pid}:${event.tid}`;
    if (!mainThreads.has(thread)) continue;
    busyByThread.set(thread, (busyByThread.get(thread) ?? 0) + event.dur / 1000);
  }
  // Other renderers (extensions, about:blank) barely run, so the busiest one is the page
  return Math.max(0, ...busyByThread.values());
}

function round(value: number) {
  return Math.round(value * 10) / 10;
}
