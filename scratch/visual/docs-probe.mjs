import { chromium } from "file:///D:/Code/web/pptx/node_modules/.pnpm/playwright@1.61.1/node_modules/playwright/index.mjs";

const dpr = Number(process.argv[2] ?? 1);
const w = Number(process.argv[3] ?? 1600);
const browser = await chromium.launch();
const page = await browser.newPage({
  viewport: { width: w, height: 1000 },
  deviceScaleFactor: dpr,
});
page.on("pageerror", (e) => console.log("[pageerror]", e.message));
page.on("console", (m) => {
  if (m.type() === "error") console.log("[console]", m.text());
});

await page.goto("http://localhost:3000/pg", { waitUntil: "load" });
await page.waitForTimeout(4000);
await page.setInputFiles('input[type="file"]', "D:/Code/web/pptx/scratch/visual/deck.pptx");
try {
  await page.waitForSelector("table", { state: "attached", timeout: 60000 });
} catch {
  await page.screenshot({ path: "D:/Code/web/pptx/scratch/visual/docs-fail.png" });
  console.log("status text:", await page.locator("body").innerText());
  await browser.close();
  process.exit(1);
}
await page.waitForTimeout(1500);
for (const t of await page.locator("button").all()) {
  const txt = (await t.textContent()) ?? "";
  if (txt.trim().startsWith("7")) {
    await t.click();
    break;
  }
}
await page.waitForTimeout(1500);

const info = await page.evaluate(() => {
  const main = [...document.querySelectorAll("table")]
    .map((t) => ({ t, r: t.getBoundingClientRect() }))
    .sort((a, b) => b.r.width - a.r.width)[0];
  const { t, r } = main;
  const cells = [...t.querySelector("tr").querySelectorAll("td")];
  const stage = t.closest("[style*='scale'],[style*='matrix']");
  return {
    scale: getComputedStyle(stage).transform,
    tableRect: { x: r.x, y: r.y, w: r.width, h: r.height },
    dividers: cells.slice(0, -1).map((td) => td.getBoundingClientRect().right),
  };
});
console.log(`dpr=${dpr} width=${w}`, JSON.stringify(info));
await page.screenshot({ path: `D:/Code/web/pptx/scratch/visual/docs-t7-${w}-${dpr}.png` });
await browser.close();
