import { chromium } from "file:///D:/Code/web/pptx/node_modules/.pnpm/playwright@1.61.1/node_modules/playwright/index.mjs";

const browser = await chromium.launch();
const page = await browser.newPage({
  viewport: { width: 1300, height: 800 },
  deviceScaleFactor: 1,
});
page.on("pageerror", (err) => console.log("[pageerror]", err.message));

for (let i = 1; i <= 8; i++) {
  await page.goto(`http://localhost:5399/?slide=${i}&scale=0.853`, { waitUntil: "load" });
  await page.waitForFunction(() => document.title !== "slide render", null, { timeout: 30000 });
  await page.waitForTimeout(700);
  const title = await page.title();
  await page
    .locator("#stage")
    .screenshot({ path: `D:/Code/web/pptx/scratch/visual/deck-${i}.png` });
  console.log(`slide ${i}: ${title}`);
}

await browser.close();
