import type { SerializedPresentation } from "@diceui/pptx-core";
import { expect, type Page } from "@playwright/test";

/** Navigates the harness to a fixture slide and waits for render to settle. */
export async function openSlide(page: Page, file: string, slide = 0): Promise<void> {
  await page.goto(`/?file=${encodeURIComponent(file)}&slide=${slide}`);
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

declare global {
  interface Window {
    __renderDone?: boolean;
    __renderError?: string;
    __slideCount?: number;
    __slideWidth?: number;
    __slideHeight?: number;
    __showSlide?: (index: number) => Promise<void>;
    __getStructure?: () => SerializedPresentation;
  }
}
