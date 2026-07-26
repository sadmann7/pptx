/**
 * E2E render harness.
 *
 * Loads a fixture deck and renders one slide, then signals completion via
 * window globals the Playwright specs poll for:
 *
 *   window.__renderDone   true once the slide (incl. async media/charts) settled
 *   window.__renderError  error message when load/render failed
 *   window.__slideCount   number of slides in the loaded deck
 *   window.__showSlide(i) re-renders another slide of the loaded deck
 *   window.__getStructure() serialized presentation structure (for structural specs)
 *
 * Query params:
 *   file   deck to load, served from fixtures/ (or decks/ for local scratch decks)
 *   slide  0-based slide index (default 0)
 *   scale  display scale (default 1)
 *   mode   how `scale` is applied: "zoom" (what the viewer ships) or
 *          "transform" (the raster-scaling alternative). Hairline rendering
 *          differs between the two, so specs comparing them need both.
 */
import type { PresentationData, SerializedPresentation, SlideHandle } from "@diceui/pptx-core";
import {
  applySlideScale,
  buildPresentation,
  readPptx,
  renderSlide,
  serializePresentation,
} from "@diceui/pptx-core";

type ScaleMode = "zoom" | "transform";

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

const container = document.getElementById("slide-container") as HTMLDivElement;
const mediaUrlCache = new Map<string, string>();

let presentation: PresentationData | undefined;
let currentHandle: SlideHandle | undefined;
let scale = 1;
let scaleMode: ScaleMode = "zoom";

function applyScale(element: HTMLElement): void {
  if (scaleMode === "transform") {
    element.style.transformOrigin = "top left";
    element.style.transform = `scale(${scale})`;
    return;
  }
  applySlideScale(element, scale);
}

async function showSlide(index: number): Promise<void> {
  if (!presentation) throw new Error("no presentation loaded");
  window.__renderDone = false;
  window.__renderError = undefined;

  currentHandle?.dispose();
  currentHandle = undefined;
  container.innerHTML = "";

  try {
    const slide = presentation.slides[index];
    if (!slide) throw new Error(`slide index ${index} out of range`);

    const handle = renderSlide(presentation, slide, { mediaUrlCache });
    currentHandle = handle;

    container.style.width = `${presentation.width * scale}px`;
    container.style.height = `${presentation.height * scale}px`;
    applyScale(handle.element);
    container.appendChild(handle.element);

    await handle.ready;
    window.__renderDone = true;
  } catch (error) {
    window.__renderError = error instanceof Error ? error.message : String(error);
    throw error;
  }
}

async function main(): Promise<void> {
  const params = new URLSearchParams(location.search);
  const file = params.get("file");
  const slideIndex = Number.parseInt(params.get("slide") ?? "0", 10);
  scale = Number.parseFloat(params.get("scale") ?? "1");
  scaleMode = params.get("mode") === "transform" ? "transform" : "zoom";

  if (!file) {
    window.__renderError = "missing ?file= query param";
    return;
  }

  try {
    const response = await fetch(`/${file}`);
    if (!response.ok) throw new Error(`fetch ${file}: HTTP ${response.status}`);
    const buffer = await response.arrayBuffer();

    const files = await readPptx(buffer);
    presentation = buildPresentation(files);

    window.__slideCount = presentation.slides.length;
    window.__slideWidth = presentation.width;
    window.__slideHeight = presentation.height;
    window.__showSlide = showSlide;
    window.__getStructure = () => {
      if (!presentation) throw new Error("no presentation loaded");
      return serializePresentation(presentation);
    };

    await showSlide(slideIndex);
  } catch (error) {
    window.__renderError = error instanceof Error ? error.message : String(error);
  }
}

void main();
