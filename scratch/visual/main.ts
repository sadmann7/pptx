import { buildPresentation, readPptx, renderSlide } from "../../packages/core/src/index";

const SLIDE_INDEX = Number(new URLSearchParams(location.search).get("slide") ?? "7") - 1;

async function main() {
  const res = await fetch("./deck.pptx");
  const buffer = await res.arrayBuffer();
  const files = await readPptx(buffer);
  const presentation = buildPresentation(files);
  const slide = presentation.slides[SLIDE_INDEX];

  const scale = Number(new URLSearchParams(location.search).get("scale") ?? "1");

  const stage = document.getElementById("stage")!;
  stage.style.width = `${presentation.width * scale}px`;
  stage.style.height = `${presentation.height * scale}px`;

  const handle = renderSlide(presentation, slide);
  handle.element.style.transformOrigin = "top left";
  handle.element.style.transform = `scale(${scale})`;
  stage.appendChild(handle.element);

  document.title = "ready";
}

main().catch((error) => {
  document.title = `error: ${error}`;
  console.error(error);
});
