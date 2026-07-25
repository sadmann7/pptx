import { chromium } from "file:///D:/Code/web/pptx/node_modules/.pnpm/playwright@1.61.1/node_modules/playwright/index.mjs";

const slide = process.argv[2] ?? "7";
const out = process.argv[3] ?? "D:/Code/web/pptx/scratch/visual/slide.png";

const browser = await chromium.launch();
const page = await browser.newPage({
  viewport: { width: 1400, height: 900 },
  deviceScaleFactor: 2,
});
page.on("console", (msg) => console.log("[console]", msg.text()));
page.on("pageerror", (err) => console.log("[pageerror]", err.message));

await page.goto(`http://localhost:5399/?slide=${slide}`, { waitUntil: "load" });
await page.waitForFunction(() => document.title !== "slide render", null, { timeout: 30000 });
console.log("title:", await page.title());
await page.waitForTimeout(1500);

const stage = page.locator("#stage");
await stage.screenshot({ path: out });
console.log("saved", out);

await browser.close();
