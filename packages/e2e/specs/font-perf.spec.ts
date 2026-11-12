/**
 * Skips unless PERF=1 (`pnpm test:perf`), and reports rather than asserts
 * because timings are machine-dependent.
 *
 * It weighs the Worker pool against its cost, a second inlined copy of the
 * decoder in the bundle, by decoding each deck in the harness's worker, main
 * and sliced modes. `blocked` and `worst frame` matter more than wall clock,
 * since the pool exists to keep the main thread free. Runs repeat under CPU
 * throttling, where a 30 KB part stops decoding too fast to show either path.
 */
import { test } from "@playwright/test";

// `FONT_PERF_DECK=decks/real.pptx` measures a PowerPoint-authored deck without
// committing it.
const DECK = process.env.FONT_PERF_DECK ?? "the-good-room-soft-editorial.pptx";

// Fixtures embed plain EOT, which decodes in microseconds, while PowerPoint
// writes MTX, so the harness substitutes those. A named deck is measured as is.
const PAYLOADS = process.env.FONT_PERF_DECK
  ? ([{ label: "as embedded", mtx: false }] as const)
  : ([
      { label: "plain EOT", mtx: false },
      { label: "MTX", mtx: true },
    ] as const);

// 4 stands in for a mid-range phone.
const THROTTLE_RATES = [1, 4] as const;

// Real decks embed 3 to 7 parts; the multiples reach the "dozens of parts" deck
// the loader was written for.
const PART_FACTORS = [1, 4, 8] as const;

const MODES: readonly BenchMode[] = ["worker", "main", "sliced"];

// Worker startup should be the blob spawn consumers pay, not Vite fetching the
// worker from source. Requires `pnpm -F @diceui/pptx-core build`.
const BUILD = "dist";

type FontBench = Awaited<ReturnType<NonNullable<Window["__benchFonts"]>>>;

interface Run extends FontBench {
  payload: string;
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
      test.skip(
        browserName !== "chromium",
        "CPU throttling and long tasks need CDP",
      );

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
              () =>
                window.__fontsReady === true ||
                window.__renderError !== undefined,
              undefined,
              { timeout: 120_000 },
            );
            const error = await page.evaluate(() => window.__renderError);
            if (error) throw new Error(`harness failed: ${error}`);

            // Throttles only the decode under test, not loading the deck.
            await cdp.send("Emulation.setCPUThrottlingRate", { rate });
            const result = await page.evaluate(() => window.__benchFonts?.());
            await cdp.send("Emulation.setCPUThrottlingRate", { rate: 1 });
            if (!result) throw new Error("harness did not expose __benchFonts");

            // Decoding the parts before the run would pollute its frame monitor.
            const parts =
              (await page.evaluate(() => window.__benchDecodeParts?.())) ?? [];
            runs.push({
              ...result,
              payload: payload.label,
              partMs: parts.map((part) => part.ms),
            });
          }
        }
      }

      const kb = (bytes: number) => `${(bytes / 1024).toFixed(0)} KB`;
      const ms = (value: number) => `${value.toFixed(0)}ms`;
      const worst = (values: number[]) =>
        values.length > 0 ? Math.max(...values) : 0;
      const mean = (values: number[]) =>
        values.length > 0
          ? values.reduce((sum, value) => sum + value, 0) / values.length
          : 0;

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
            // Fewer faces than parts means the decoder or the browser rejected some.
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
