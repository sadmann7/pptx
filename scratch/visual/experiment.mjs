import { chromium } from "file:///D:/Code/web/pptx/node_modules/.pnpm/playwright@1.61.1/node_modules/playwright/index.mjs";

const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 800, height: 700 }, deviceScaleFactor: 1 });
await page.goto("http://localhost:5399/experiment.html", { waitUntil: "load" });
await page.waitForTimeout(400);

const geo = await page.evaluate(() =>
  ["a", "b", "c", "d"].map((id) => {
    const rows = [...document.querySelectorAll(`#${id} tr`)];
    const dark = rows[0].querySelectorAll("td")[0].getBoundingClientRect();
    const light = rows[1].querySelectorAll("td")[0].getBoundingClientRect();
    return {
      id,
      vBorderX: dark.right,
      darkMidY: dark.top + dark.height / 2,
      lightMidY: light.top + light.height / 2,
      hBorderY: dark.bottom,
      sampleX: dark.left + 40,
    };
  }),
);

await page.screenshot({ path: "D:/Code/web/pptx/scratch/visual/experiment.png" });
console.log(JSON.stringify(geo));
await browser.close();
