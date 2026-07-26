import { chromium } from "@playwright/test";
/**
 * Screenshots a harness slide and dumps the geometry of whatever you want to
 * measure. Companion to probe.ts for investigating a rendering bug by hand.
 *
 * Start the harness first (`pnpm harness`), then:
 *   pnpm shoot --file table-borders.pptx --slide 0 --scale 0.86
 *   pnpm shoot --file decks/customer.pptx --slide 6 --select table --select "table td"
 *   pnpm shoot --file table-borders.pptx --mode transform --dpr 1.5
 *
 * --select prints each match's client rect plus the styles that decide how a
 * hairline lands on the device pixel grid, which is what a probe run is keyed
 * to. Local decks go in decks/ (gitignored); fixtures are addressed by name.
 */
import { mkdir } from "node:fs/promises";
import { dirname, resolve } from "node:path";

import { parseArgs } from "./args";

const args = parseArgs(process.argv.slice(2), {
  numbers: ["slide", "scale", "dpr", "width", "height", "port"],
  strings: ["file", "mode", "out"],
  lists: ["select"],
});

const file = args.strings.file ?? "table-borders.pptx";
const slide = args.numbers.slide ?? 0;
const scale = args.numbers.scale ?? 1;
const mode = args.strings.mode ?? "zoom";
const port = args.numbers.port ?? 5000;
const out = resolve(args.strings.out ?? `out/${file.replace(/[/\\]/g, "-")}-${slide}.png`);
const url = `http://localhost:${port}/?file=${encodeURIComponent(file)}&slide=${slide}&scale=${scale}&mode=${mode}`;

const browser = await chromium.launch();
const page = await browser.newPage({
  viewport: { width: args.numbers.width ?? 1400, height: args.numbers.height ?? 900 },
  deviceScaleFactor: args.numbers.dpr ?? 1,
});
page.on("pageerror", (error) => console.error("[pageerror]", error.message));

await page.goto(url, { waitUntil: "load" });
await page.waitForFunction(
  () => window.__renderDone === true || window.__renderError !== undefined,
  null,
  { timeout: 30_000 },
);
const renderError = await page.evaluate(() => window.__renderError);
if (renderError) throw new Error(renderError);

for (const selector of args.lists.select) {
  const found = await page.evaluate(
    (sel) =>
      // Everything stays inline: named functions do not survive the transpiler's
      // rewrite when the body is serialized into the page.
      [...document.querySelectorAll(sel)].map((el) => {
        const rect = el.getBoundingClientRect();
        const style = getComputedStyle(el);
        return {
          text: (el.textContent ?? "").trim().slice(0, 16),
          x: Math.round(rect.x * 100) / 100,
          y: Math.round(rect.y * 100) / 100,
          right: Math.round(rect.right * 100) / 100,
          bottom: Math.round(rect.bottom * 100) / 100,
          w: Math.round(rect.width * 100) / 100,
          h: Math.round(rect.height * 100) / 100,
          zoom: style.zoom,
          transform: style.transform,
          background: style.backgroundColor,
          borders: [style.borderTop, style.borderRight, style.borderBottom, style.borderLeft],
        };
      }),
    selector,
  );
  console.log(`\n${selector} (${found.length})`);
  for (const entry of found) console.log(" ", JSON.stringify(entry));
}

await mkdir(dirname(out), { recursive: true });
await page.locator("#slide-container").screenshot({ path: out });
await browser.close();
console.log(`\nwrote ${out}`);
