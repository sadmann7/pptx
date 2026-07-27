/**
 * Smoke specs for the exported decks (specs/decks.ts).
 *
 * The oracle scores these decks against PowerPoint, but SSIM is a blunt
 * instrument: a dropped image on a busy slide, or a shape that renders as an
 * empty box, can cost less score than a font substitution and hide under the
 * baseline tolerance. These walk every slide and assert the things a score
 * cannot see: no thrown errors, no empty slide, no unresolved media.
 */
import { expect, type Page, test } from "@playwright/test";

import { EXPORTED_DECKS } from "./decks";
import { getStructure, openSlide, slideContainer } from "./utils";

// One test covers a whole deck, so it pays for nine harness loads of a full
// presentation and runs past the default 30s once the rest of the suite is
// competing for the CPU.
test.describe.configure({ timeout: 120_000 });

/**
 * Sources that failed to decode. Covers both <img> elements and the blob URLs
 * the renderer hands to background-image for fills and crops.
 */
async function brokenMedia(page: Page): Promise<string[]> {
  return page.evaluate(async () => {
    const container = document.querySelector("#slide-container");
    if (!container) return ["#slide-container is missing"];

    const failures: string[] = [];
    for (const img of container.querySelectorAll("img")) {
      if (!img.complete || img.naturalWidth === 0) failures.push(`img ${img.src.slice(0, 80)}`);
    }

    const urls = new Set<string>();
    for (const element of container.querySelectorAll("*")) {
      const background = getComputedStyle(element).backgroundImage;
      for (const match of background.matchAll(/url\("([^"]+)"\)/g)) urls.add(match[1]);
    }
    await Promise.all(
      [...urls].map(
        (url) =>
          new Promise<void>((resolve) => {
            const probe = new Image();
            probe.onload = () => resolve();
            probe.onerror = () => {
              failures.push(`background ${url.slice(0, 80)}`);
              resolve();
            };
            probe.src = url;
          }),
      ),
    );

    return failures;
  });
}

for (const { name, slides } of EXPORTED_DECKS) {
  test.describe(`exported deck: ${name}`, () => {
    test("renders every slide with content and no failures", async ({ page }) => {
      const failures: string[] = [];
      page.on("pageerror", (error) => failures.push(`pageerror: ${error.message}`));
      page.on("console", (message) => {
        if (message.type() === "error") failures.push(`console: ${message.text()}`);
      });

      await openSlide(page, `${name}.pptx`, 0);
      const structure = await getStructure(page);
      expect(structure.slideCount).toBe(slides);

      for (const [index, slide] of structure.slides.entries()) {
        expect(slide.nodes.length, `slide ${index + 1} parsed no nodes`).toBeGreaterThan(0);
      }

      for (let slide = 0; slide < slides; slide++) {
        await openSlide(page, `${name}.pptx`, slide);
        await expect(slideContainer(page)).toBeVisible();
        expect(await brokenMedia(page), `slide ${slide + 1} has unresolved media`).toEqual([]);
      }

      expect(failures).toEqual([]);
    });
  });
}
