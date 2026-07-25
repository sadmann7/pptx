import { mkdir } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { chromium } from "playwright";

import { parseArgs } from "./args.mjs";

/**
 * Screenshot a slide from the dev harness and dump the geometry of any
 * elements you want to measure.
 *
 * Start the harness first (`pnpm --filter @pptx/visual dev`), then:
 *   node shoot.mjs --slide 7 --scale 0.86 --out out/slide7.png
 *   node shoot.mjs --slide 5 --select table --select "table td"
 *   node shoot.mjs --slide 5 --mode transform --dpr 1.5
 *
 * `--select` prints each match's client rect plus the styles that decide how a
 * hairline lands on the device pixel grid, which is what the geometry in a
 * probe run is keyed to.
 */
const args = parseArgs(process.argv.slice(2), {
  numbers: ["slide", "scale", "dpr", "width", "height", "port"],
  lists: ["select"],
  strings: ["file", "mode", "out", "url"],
});

const slide = args.slide ?? 1;
const scale = args.scale ?? 1;
const port = args.port ?? 5399;
const out = resolve(args.out ?? `out/slide-${slide}.png`);
const url =
  args.url ??
  `http://localhost:${port}/?file=${args.file ?? "deck.pptx"}&slide=${slide}&scale=${scale}&mode=${args.mode ?? "zoom"}`;

const browser = await chromium.launch();
const page = await browser.newPage({
  viewport: { width: args.width ?? 1400, height: args.height ?? 900 },
  deviceScaleFactor: args.dpr ?? 1,
});
page.on("pageerror", (error) => console.error("[pageerror]", error.message));

await page.goto(url, { waitUntil: "load" });
await page.waitForFunction(() => document.title !== "slide render", null, { timeout: 30_000 });
const title = await page.title();
if (title.startsWith("error:")) throw new Error(title);
// Fonts and chart canvases settle a frame or two after the slide is in the DOM.
await page.waitForTimeout(500);

for (const selector of args.select ?? []) {
  const found = await page.evaluate((sel) => {
    const round = (value) => Math.round(value * 100) / 100;
    return [...document.querySelectorAll(sel)].map((el) => {
      const rect = el.getBoundingClientRect();
      const style = getComputedStyle(el);
      return {
        text: (el.textContent ?? "").trim().slice(0, 16),
        x: round(rect.x),
        y: round(rect.y),
        right: round(rect.right),
        bottom: round(rect.bottom),
        w: round(rect.width),
        h: round(rect.height),
        zoom: style.zoom,
        transform: style.transform,
        background: style.backgroundColor,
        borders: [style.borderTop, style.borderRight, style.borderBottom, style.borderLeft],
      };
    });
  }, selector);
  console.log(`\n${selector} (${found.length})`);
  for (const entry of found) console.log(" ", JSON.stringify(entry));
}

await mkdir(dirname(out), { recursive: true });
await page.screenshot({ path: out });
await browser.close();
console.log(`\nwrote ${out}`);
