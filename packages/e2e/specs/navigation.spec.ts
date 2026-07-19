import { expect, test } from "@playwright/test";

import { openSlide, slideContainer, waitForRender } from "./helpers";

test.describe("slide navigation", () => {
  // Regression: switching between structurally identical slides skipped the
  // rebuild, so the main view kept showing the previous slide's content.
  test("re-renders when switching to another slide", async ({ page }) => {
    await openSlide(page, "basic.pptx", 0);
    await expect(slideContainer(page)).toContainText("Slide one");

    await page.evaluate(() => window.__showSlide?.(1));
    await waitForRender(page);
    await expect(slideContainer(page)).toContainText("Slide two");
    await expect(slideContainer(page)).not.toContainText("Slide one");

    await page.evaluate(() => window.__showSlide?.(2));
    await waitForRender(page);
    await expect(slideContainer(page)).toContainText("Slide three");
  });

  test("reports the deck's slide count", async ({ page }) => {
    await openSlide(page, "basic.pptx", 0);
    expect(await page.evaluate(() => window.__slideCount)).toBe(3);
  });
});
