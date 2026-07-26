import type { SerializedPresentation } from "@diceui/pptx-core";
import { expect, type Page } from "@playwright/test";

export interface OpenSlideOptions {
  /** Display scale; 1 (native size) unless a spec cares about scaled output. */
  scale?: number;
  /**
   * How the scale is applied. "zoom" is what the viewer ships; "transform" is
   * the raster-scaling alternative, used as a control in hairline specs.
   */
  mode?: "zoom" | "transform";
}

/** Navigates the harness to a fixture slide and waits for render to settle. */
export async function openSlide(
  page: Page,
  file: string,
  slide = 0,
  options: OpenSlideOptions = {},
): Promise<void> {
  const query = new URLSearchParams({ file, slide: String(slide) });
  if (options.scale !== undefined) query.set("scale", String(options.scale));
  if (options.mode !== undefined) query.set("mode", options.mode);
  await page.goto(`/?${query}`);
  await waitForRender(page);
}

/** Waits for the harness to flag the current render as done (and error-free). */
export async function waitForRender(page: Page): Promise<void> {
  await page.waitForFunction(
    () => window.__renderDone === true || window.__renderError !== undefined,
  );
  const error = await page.evaluate(() => window.__renderError);
  expect(error, "harness reported a render error").toBeUndefined();
}

export function slideContainer(page: Page) {
  return page.locator("#slide-container");
}

/** Returns the serialized presentation structure from the loaded harness. */
export async function getStructure(page: Page): Promise<SerializedPresentation> {
  const structure = await page.evaluate(() => window.__getStructure?.());
  if (!structure) throw new Error("harness did not expose a presentation structure");
  return structure;
}
