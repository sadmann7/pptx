import { chromium } from "file:///D:/Code/web/pptx/node_modules/.pnpm/playwright@1.61.1/node_modules/playwright/index.mjs";

const browser = await chromium.launch();
const page = await browser.newPage({
  viewport: { width: 1300, height: 800 },
  deviceScaleFactor: 1,
});
await page.goto("http://localhost:5399/?slide=7&scale=0.853", { waitUntil: "load" });
await page.waitForFunction(() => document.title !== "slide render", null, { timeout: 30000 });
await page.waitForTimeout(500);

// Mimic a table whose cells differ within a row: the renderer skips the row
// backdrop in that case, so clear it here too.
const geo = await page.evaluate(() => {
  const rows = [...document.querySelectorAll("table tr")];
  rows.forEach((tr, r) => {
    tr.style.backgroundColor = "";
    [...tr.querySelectorAll("td")].forEach((td, c) => {
      td.style.backgroundColor = (r + c) % 2 === 0 ? "#0F0F0F" : "#EFE9D9";
    });
  });
  const cell = (r, c) => rows[r].querySelectorAll("td")[c].getBoundingClientRect();
  const a = cell(1, 0);
  return { vBorderX: a.right, midY: a.top + a.height / 2 };
});
console.log(JSON.stringify(geo));
await page.locator("#stage").screenshot({ path: "D:/Code/web/pptx/scratch/visual/checker.png" });
await browser.close();
