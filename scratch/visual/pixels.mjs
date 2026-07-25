import { chromium } from "file:///D:/Code/web/pptx/node_modules/.pnpm/playwright@1.61.1/node_modules/playwright/index.mjs";

const scale = process.argv[2] ?? "0.853";
const dpr = Number(process.argv[3] ?? "1");
const out = process.argv[4] ?? "D:/Code/web/pptx/scratch/visual/scaled.png";

const browser = await chromium.launch();
const page = await browser.newPage({
  viewport: { width: 1400, height: 900 },
  deviceScaleFactor: dpr,
});
await page.goto(`http://localhost:5399/?slide=7&scale=${scale}`, { waitUntil: "load" });
await page.waitForFunction(() => document.title !== "slide render", null, { timeout: 30000 });
await page.waitForTimeout(800);

const geo = await page.evaluate(() => {
  const rows = [...document.querySelectorAll("table tr")];
  const cell = (r, c) => rows[r].querySelectorAll("td")[c].getBoundingClientRect();
  const header = cell(0, 0);
  const body = cell(1, 0);
  const body2 = cell(2, 0);
  return {
    vBorderX: header.right,
    headerMidY: header.top + header.height / 2,
    bodyMidY: body.top + body.height / 2,
    hBorderHeaderBodyY: header.bottom,
    hBorderBodyBodyY: body2.top,
    sampleX: header.left + 100,
  };
});
console.log("geometry:", JSON.stringify(geo));
await page.screenshot({ path: out });
await browser.close();
