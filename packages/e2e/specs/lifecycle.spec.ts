/**
 * A viewer that unmounts has to take its blob URLs and embedded fonts with it,
 * or every deck opened in a long-lived page stays in memory for good.
 */
import { expect, type Page, test } from "@playwright/test";

async function mountViewer(page: Page, file: string, store: ViewerStore = "root"): Promise<void> {
  await page.evaluate(([name, from]) => window.__mountViewer?.(name, from), [file, store] as const);
}

async function waitForLoad(page: Page): Promise<void> {
  await page.waitForFunction(
    () => window.__viewerLoaded === true || window.__renderError !== undefined,
  );
  expect(await page.evaluate(() => window.__renderError)).toBeUndefined();
}

async function isBlobAlive(page: Page, url: string): Promise<boolean> {
  return page.evaluate(
    (src) =>
      fetch(src).then(
        () => true,
        () => false,
      ),
    url,
  );
}

test.beforeEach(async ({ page }) => {
  await page.goto("/lifecycle.html");
});

test("unmounting the viewer revokes the slide's picture URLs", async ({ page }) => {
  await mountViewer(page, "pictures.pptx");
  await waitForLoad(page);

  const picture = page.locator("[data-presentation-slide] img");
  await expect(picture).toHaveCount(1);
  const src = await picture.getAttribute("src");
  if (!src?.startsWith("blob:")) throw new Error(`expected a blob URL, got ${src}`);
  expect(await isBlobAlive(page, src)).toBe(true);

  await page.evaluate(() => window.__unmountViewer?.());
  expect(await isBlobAlive(page, src)).toBe(false);
});

const FONT_DECK = "adventure-club-pin-and-paper.pptx";

const STORES: readonly { store: ViewerStore; label: string }[] = [
  { store: "root", label: "Root's own store" },
  { store: "hook", label: "useCreatePresentationStore" },
];

for (const { store, label } of STORES) {
  test.describe(`with ${label}`, () => {
    test("unmounting the viewer removes the deck's embedded fonts", async ({ page }) => {
      await mountViewer(page, FONT_DECK, store);
      await waitForLoad(page);
      expect(await page.evaluate(() => document.fonts.size)).toBeGreaterThan(0);

      await page.evaluate(() => window.__unmountViewer?.());
      expect(await page.evaluate(() => document.fonts.size)).toBe(0);
    });

    test("unmounting the viewer mid-load removes the fonts it registered", async ({ page }) => {
      await mountViewer(page, FONT_DECK, store);
      await page.waitForFunction(() => document.fonts.size > 0);

      await page.evaluate(() => window.__unmountViewer?.());
      await page.waitForFunction(() => document.fonts.size === 0);
    });
  });
}
