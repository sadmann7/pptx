import { chromium } from "file:///D:/Code/web/pptx/node_modules/.pnpm/playwright@1.61.1/node_modules/playwright/index.mjs";

const browser = await chromium.launch();

for (const dpr of [1, 2]) {
  for (const scale of ["0.853", "1", "1.07", "1.25", "0.6"]) {
    const page = await browser.newPage({
      viewport: { width: 1500, height: 1000 },
      deviceScaleFactor: dpr,
    });
    await page.goto(`http://localhost:5399/?slide=7&scale=${scale}`, { waitUntil: "load" });
    await page.waitForFunction(() => document.title !== "slide render", null, { timeout: 30000 });
    await page.waitForTimeout(500);

    const report = await page.evaluate(() => {
      const table = document.querySelector("table");
      const firstRow = table.querySelector("tr");
      const cells = [...firstRow.querySelectorAll("td")];
      return {
        tableTop: table.getBoundingClientRect().top,
        tableRect: (({ x, y, width, height }) => ({ x, y, width, height }))(
          table.getBoundingClientRect(),
        ),
        wrapperRect: (({ x, y, width, height }) => ({ x, y, width, height }))(
          table.parentElement.getBoundingClientRect(),
        ),
        cellTops: cells.map((td) => td.getBoundingClientRect().top),
        dividers: cells.slice(0, -1).map((td) => td.getBoundingClientRect().right),
        rowHeightStyle: firstRow.style.height,
        tableHeightStyle: table.style.height,
      };
    });
    console.log(`dpr=${dpr} scale=${scale}`, JSON.stringify(report));
    await page.screenshot({ path: `D:/Code/web/pptx/scratch/visual/poke-${dpr}-${scale}.png` });
    await page.close();
  }
}

await browser.close();
