import { applySlideScale, buildPresentation, readPptx, renderSlide } from "@diceui/pptx-core";

/**
 * Renders one slide of one deck, isolated from the viewer and the docs app.
 *
 * Query parameters:
 *   file   deck under decks/ (default: deck.pptx)
 *   slide  1-based slide number (default: 1)
 *   scale  display scale (default: 1)
 *   mode   "zoom" (viewer default, layout scaling) or "transform" (raster scaling)
 *
 * The title becomes "ready" once the slide is in the DOM, which is what the
 * Playwright drivers wait on; on failure it carries the error message.
 */
async function main(): Promise<void> {
  const params = new URLSearchParams(location.search);
  const file = params.get("file") ?? "deck.pptx";
  const slideIndex = Number(params.get("slide") ?? "1") - 1;
  const scale = Number(params.get("scale") ?? "1");
  const mode = params.get("mode") ?? "zoom";

  const response = await fetch(`./decks/${file}`);
  if (!response.ok) throw new Error(`${file}: ${response.status}`);
  const files = await readPptx(await response.arrayBuffer());
  const presentation = buildPresentation(files);
  const slide = presentation.slides[slideIndex];
  if (!slide) throw new Error(`slide ${slideIndex + 1} of ${presentation.slides.length}`);

  const stage = document.getElementById("stage");
  if (!stage) throw new Error("missing #stage");
  stage.style.width = `${presentation.width * scale}px`;
  stage.style.height = `${presentation.height * scale}px`;

  const handle = renderSlide(presentation, slide);
  if (mode === "transform") {
    handle.element.style.transformOrigin = "top left";
    handle.element.style.transform = `scale(${scale})`;
  } else {
    applySlideScale(handle.element, scale);
  }
  stage.appendChild(handle.element);

  document.title = "ready";
}

main().catch((error: unknown) => {
  document.title = `error: ${error instanceof Error ? error.message : String(error)}`;
  console.error(error);
});
