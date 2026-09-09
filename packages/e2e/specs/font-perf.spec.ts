/**
 * Skipped unless PERF=1 (`pnpm test:perf`): it decodes real embedded fonts both
 * ways and reports numbers rather than asserting a baseline, because timings
 * are machine-dependent.
 *
 * The question it answers is whether the Worker pool earns its place. The pool
 * costs a second, inlined copy of the decode pipeline in the published bundle
 * (tsdown has no worker support, so `tsdown.config.ts` bundles the worker into
 * a string), and the decoder is the bulk of that chunk. So the comparison is
 * per deck:
 *
 *   worker  the shipped path: parts decoded across a pool, registered here
 *   main    the shipped fallback (SSR, worker load failure): decoded here,
 *           handing the event loop a turn between parts
 *   sliced  the fallback without a worker to fall back from: decoded here,
 *           yielding only once a slice has run long, which is what dropping
 *           the worker would actually mean
 *
 * Wall clock is the weaker half of the answer. The pool exists to keep the main
 * thread free, so what matters is `blocked` (time frames arrived late, summed)
 * and `worst frame`: with a worker those stay near idle, on the main thread they
 * track decode cost directly.
 *
 * Both are run under CPU throttling too, since the decision is about low-end
 * devices; a desktop decodes a 30 KB part fast enough that neither path shows.
 */
import { test } from "@playwright/test";

/**
 * The deck under test: a committed fixture, or any deck in `decks/` through
 * `FONT_PERF_DECK` (e.g. `FONT_PERF_DECK=decks/real.pptx`), which is how a
 * PowerPoint-authored deck gets measured without committing one.
 */
const DECK = process.env.FONT_PERF_DECK ?? "the-good-room-soft-editorial.pptx";

/**
 * The two payload kinds, because they decide the whole answer.
 *
 * Every fixture here embeds plain EOT, where decoding strips a header and costs
 * microseconds. PowerPoint writes MTX-compressed parts (LZCOMP over three
 * streams), the only case with real work in it, so the harness can substitute
 * those payloads. A deck named explicitly is measured as it is, since the point
 * of naming one is to see what it actually carries.
 */
const PAYLOADS = process.env.FONT_PERF_DECK
  ? ([{ label: "as embedded", mtx: false }] as const)
  : ([
      { label: "plain EOT", mtx: false },
      { label: "MTX", mtx: true },
    ] as const);

/** 1 is this machine; 4 stands in for a mid-range phone. */
const THROTTLE_RATES = [1, 4] as const;

/**
 * Font parts per deck, as a multiple of what it embeds (3 to 7 today). The
 * multiples cover the "dozens of parts" deck the loader was written for.
 */
const PART_FACTORS = [1, 4, 8] as const;

const MODES: readonly BenchMode[] = ["worker", "main", "sliced"];

/**
 * The published entry, so worker startup is the blob spawn consumers pay and
 * not Vite fetching the worker module graph from source. Requires
 * `pnpm -F @diceui/pptx-core build`; the harness fails loudly without it.
 */
const BUILD = "dist";

type FontBench = Awaited<ReturnType<NonNullable<Window["__benchFonts"]>>>;

interface Run extends FontBench {
  payload: string;
  /** Per-part main-thread decode cost, measured after the run under test. */
  partMs: number[];
}

test.describe("embedded font decode performance", () => {
  test.skip(!process.env.PERF, "perf run only: PERF=1 pnpm test:perf");
  test.setTimeout(300_000);

  for (const rate of THROTTLE_RATES) {
    test(`worker pool against main thread at ${rate}x CPU throttling`, async ({
      page,
      browserName,
    }) => {
      test.skip(browserName !== "chromium", "CPU throttling and long tasks need CDP");

      const cdp = await page.context().newCDPSession(page);

      const runs: Run[] = [];
      for (const payload of PAYLOADS) {
        for (const factor of PART_FACTORS) {
          for (const mode of MODES) {
            const query = new URLSearchParams({
              file: DECK,
              mode,
              parts: String(factor),
              build: BUILD,
              ...(payload.mtx ? { mtx: "1" } : {}),
            });
            await page.goto(`/fonts.html?${query}`);
            await page.waitForFunction(
              () => window.__fontsReady === true || window.__renderError !== undefined,
              undefined,
              { timeout: 120_000 },
            );
            const error = await page.evaluate(() => window.__renderError);
            if (error) throw new Error(`harness failed: ${error}`);

            // Throttling starts after the deck is parsed and ends before the
            // next navigation, so it applies to the decode under test and not
            // to loading the deck, which is not what is being compared.
            await cdp.send("Emulation.setCPUThrottlingRate", { rate });
            const result = await page.evaluate(() => window.__benchFonts?.());
            await cdp.send("Emulation.setCPUThrottlingRate", { rate: 1 });
            if (!result) throw new Error("harness did not expose __benchFonts");

            // After the run, never before: decoding the parts up front would
            // warm nothing the loader reuses but would pollute the frame
            // monitor of the run being measured.
            const parts = (await page.evaluate(() => window.__benchDecodeParts?.())) ?? [];
            runs.push({ ...result, payload: payload.label, partMs: parts.map((part) => part.ms) });
          }
        }
      }

      const kb = (bytes: number) => `${(bytes / 1024).toFixed(0)} KB`;
      const ms = (value: number) => `${value.toFixed(0)}ms`;
      const worst = (values: number[]) => (values.length > 0 ? Math.max(...values) : 0);
      const mean = (values: number[]) =>
        values.length > 0 ? values.reduce((sum, value) => sum + value, 0) / values.length : 0;

      const lines = [
        "",
        `deck                  ${DECK}`,
        `CPU throttling        ${rate}x`,
        "",
        "payload    parts  size      mode    complete  ready   blocked  worst frame  faces",
      ];
      for (const run of runs) {
        lines.push(
          [
            run.payload.padEnd(11),
            `${run.parts}`.padEnd(7),
            kb(run.bytes).padEnd(10),
            run.mode.padEnd(8),
            ms(run.completeMs).padEnd(10),
            ms(run.readyMs).padEnd(8),
            ms(run.blockedMs).padEnd(9),
            ms(run.worstFrameMs).padEnd(13),
            // Fewer faces than parts means parts the decoder or the browser
            // rejected, which no timing here would reveal.
            `${run.faces}`,
          ].join(""),
        );
      }

      lines.push("", "per-part decode cost on the main thread");
      for (const run of runs) {
        if (run.mode !== "sliced") continue;
        lines.push(
          `  ${run.payload.padEnd(11)}${`${run.parts} parts`.padEnd(10)}${kb(run.bytes).padEnd(10)}mean ${mean(run.partMs).toFixed(2)}ms   worst ${worst(run.partMs).toFixed(2)}ms`,
        );
      }
      lines.push("");

      console.log(lines.join("\n"));
    });
  }
});
