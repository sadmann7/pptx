import { chromium } from "file:///D:/Code/web/pptx/node_modules/.pnpm/playwright@1.61.1/node_modules/playwright/index.mjs";

const browser = await chromium.launch();
const page = await browser.newPage({
  viewport: { width: 1400, height: 900 },
  deviceScaleFactor: 1,
});
page.on("pageerror", (err) => console.log("[pageerror]", err.message));

await page.goto("http://localhost:5399/?slide=7", { waitUntil: "load" });
await page.waitForFunction(() => document.title !== "slide render", null, { timeout: 30000 });
await page.waitForTimeout(800);

const info = await page.evaluate(() => {
  const table = document.querySelector("table");
  const cells = [...table.querySelectorAll("td")];
  const wrapper = table.parentElement;
  const rowOf = (td) => td.parentElement.rowIndex;
  const out = cells.slice(0, 8).map((td) => {
    const cs = getComputedStyle(td);
    const r = td.getBoundingClientRect();
    return {
      row: rowOf(td),
      text: td.textContent.trim().slice(0, 12),
      bg: cs.backgroundColor,
      top: `${cs.borderTopWidth} ${cs.borderTopColor}`,
      bottom: `${cs.borderBottomWidth} ${cs.borderBottomColor}`,
      left: `${cs.borderLeftWidth} ${cs.borderLeftColor}`,
      right: `${cs.borderRightWidth} ${cs.borderRightColor}`,
      rect: `${r.x.toFixed(2)},${r.y.toFixed(2)} ${r.width.toFixed(2)}x${r.height.toFixed(2)}`,
    };
  });
  return {
    wrapperStyle: wrapper.getAttribute("style"),
    tableRect: JSON.stringify(table.getBoundingClientRect()),
    cells: out,
  };
});
console.log(JSON.stringify(info, null, 2));

// Zoom the header/body boundary and a body/body boundary for visual inspection.
const table = page.locator("table");
const box = await table.boundingBox();
await page.screenshot({
  path: "D:/Code/web/pptx/scratch/visual/zoom-boundary.png",
  clip: { x: box.x + 340, y: box.y + 20, width: 140, height: 120 },
  scale: "css",
});

await browser.close();
