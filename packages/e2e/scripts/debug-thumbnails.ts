import { chromium } from "@playwright/test";

const url =
  process.argv[2] ??
  "http://localhost:5000/thumbnails.html?file=decks/bench-thumbnails.pptx&slides=64";

const browser = await chromium.launch();
const page = await browser.newPage();
page.on("console", (message) => console.log(`[${message.type()}]`, message.text()));
page.on("pageerror", (error) => console.log("[pageerror]", error.message));

await page.goto(url);
for (let i = 0; i < 12; i++) {
  await page.waitForTimeout(2500);
  const state = await page.evaluate(() => ({
    ready: window.__thumbnailsReady,
    slides: window.__slideCount,
    error: window.__renderError,
    items: document.querySelectorAll("[role='option']").length,
  }));
  console.log(state);
  if (state.ready) break;
}

await browser.close();
