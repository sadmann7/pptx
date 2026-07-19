/**
 * Slide-level rendering behavior (the container `renderSlide` produces).
 * Currently covers the `clipContent` option: edit-mode pasteboard behavior
 * keeps off-slide shapes visible; viewing clips them.
 */
import { describe, expect, it } from "vitest";

import { buildPresentation } from "../../model/presentation";
import { readPptx } from "../../ooxml/zip";
import type { SlideRendererOptions } from "../../renderer/slide";
import { renderSlide } from "../../renderer/slide";
import { buildRichPptx } from "../fixtures/rich-pptx";

async function renderContainer(options?: SlideRendererOptions): Promise<HTMLElement> {
  const buffer = await buildRichPptx({});
  const files = await readPptx(buffer);
  const presentation = buildPresentation(files);
  return renderSlide(presentation, presentation.slides[0], options).element;
}

describe("slide container clipping", () => {
  it("clips content by default", async () => {
    const element = await renderContainer();
    expect(element.style.overflow).toBe("hidden");
  });

  it("clips content when clipContent is true", async () => {
    const element = await renderContainer({ clipContent: true });
    expect(element.style.overflow).toBe("hidden");
  });

  it("leaves content visible when clipContent is false", async () => {
    const element = await renderContainer({ clipContent: false });
    expect(element.style.overflow).toBe("visible");
  });
});
