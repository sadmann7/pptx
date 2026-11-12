import type { SerializedPresentation } from "@diceui/pptx-core";
import { expect, type Page } from "@playwright/test";

export interface OpenSlideOptions {
  scale?: number;
  // "zoom" is what the viewer ships; "transform" is a control in hairline specs.
  mode?: "zoom" | "transform";
}

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

export async function getStructure(page: Page): Promise<SerializedPresentation> {
  const structure = await page.evaluate(() => window.__getStructure?.());
  if (!structure) throw new Error("harness did not expose a presentation structure");
  return structure;
}
