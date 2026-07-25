import { chromium } from "file:///D:/Code/web/pptx/node_modules/.pnpm/playwright@1.61.1/node_modules/playwright/index.mjs";

const browser = await chromium.launch();

for (const dpr of [1, 1.25, 1.5]) {
  for (const v of ["a", "b", "c"]) {
    const page = await browser.newPage({
      viewport: { width: 1300, height: 700 },
      deviceScaleFactor: dpr,
    });
    await page.goto(`http://localhost:5399/corner.html?v=${v}`, { waitUntil: "load" });
    await page.waitForTimeout(300);

    const geo = await page.evaluate(() => {
      const t = document.querySelector("table");
      const r = t.getBoundingClientRect();
      const cells = [...t.querySelector("tr").querySelectorAll("td")];
      return {
        top: r.top,
        left: r.left,
        right: r.right,
        dividers: cells.slice(0, -1).map((td) => td.getBoundingClientRect().right),
      };
    });

    const shot = await page.screenshot();
    const { createCanvas, loadImage } =
      await import("file:///D:/Code/web/pptx/node_modules/.pnpm/@napi-rs+canvas@0.1.80/node_modules/@napi-rs/canvas/index.js").catch(
        () => ({}),
      );
    await page.screenshot({ path: `D:/Code/web/pptx/scratch/visual/corner-${dpr}-${v}.png` });
    console.log(`dpr=${dpr} v=${v}`, JSON.stringify(geo));
    void shot;
    void createCanvas;
    void loadImage;
    await page.close();
  }
}

await browser.close();
