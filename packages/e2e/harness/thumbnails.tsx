/**
 * Mounts the real `Presentation.ThumbnailList` in a fixed-height scroller for
 * thumbnail-perf.spec.ts to scroll.
 *
 * Query params:
 *   file   deck to load, served from fixtures/ (or decks/ for local decks)
 *   slides grows the deck to this many slides by duplicating its own slides,
 *          so the list is long enough to scroll through (default: as loaded)
 *   width  thumbnail column width in px (default 180)
 *   height scroller height in px (default 720)
 */
import * as React from "react";

import type { PresentationStore } from "@diceui/pptx";
import {
  Presentation,
  useCreatePresentationStore,
  usePresentation,
} from "@diceui/pptx";
import type {
  PresentationData,
  SlideData,
  SlideHandle,
} from "@diceui/pptx-core";
import { renderSlide, renderThumbnail } from "@diceui/pptx-core";
import { createRoot } from "react-dom/client";

const params = new URLSearchParams(location.search);
const file = params.get("file");
const targetSlides = Number.parseInt(params.get("slides") ?? "0", 10);
const width = Number.parseInt(params.get("width") ?? "180", 10);
const height = Number.parseInt(params.get("height") ?? "720", 10);

interface RenderTimings {
  render: number[];
  mount: number[];
}

// Mount time is a lower bound: the host is off-screen, so it covers insertion
// and the forced layout but not paint.
function timeRenders(
  presentation: PresentationData,
  render: (presentation: PresentationData, slide: SlideData) => SlideHandle,
): RenderTimings {
  // Off to the side at the size a preview really occupies: laid out in full,
  // but never painted and never disturbing the list being measured.
  const host = document.createElement("div");
  host.style.cssText = `position:absolute;top:0;left:-100000px;width:${width}px;height:${height}px;overflow:hidden;pointer-events:none`;
  document.body.appendChild(host);

  const timings: RenderTimings = { render: [], mount: [] };
  for (const slide of presentation.slides) {
    const renderStart = performance.now();
    const handle = render(presentation, slide);
    timings.render.push(performance.now() - renderStart);

    const mountStart = performance.now();
    host.appendChild(handle.element);
    // Appending only schedules the layout; reading a geometric property is
    // what makes it happen here rather than in some later frame.
    void host.offsetHeight;
    timings.mount.push(performance.now() - mountStart);

    handle.element.remove();
    handle.dispose();
  }

  host.remove();
  return timings;
}

// Grown in the browser rather than written to disk, since the parser needs the
// DOM APIs.
async function growDeck(store: PresentationStore): Promise<void> {
  const sourceIds =
    store.getState().presentation?.slides.map((slide) => slide.id) ?? [];
  if (sourceIds.length === 0) return;

  let index = 0;
  while ((store.getState().presentation?.slides.length ?? 0) < targetSlides) {
    const slideId = sourceIds[index % sourceIds.length];
    if (!slideId) break;
    await store.edit({ type: "duplicateSlide", slideId });
    index++;
  }
}

function Harness({ store }: { store: PresentationStore }) {
  const { presentation, status } = usePresentation();

  // Duplicating a slide mutates the presentation in place, so its identity is
  // unchanged and only the store notification reports the new length.
  const slideCount = React.useSyncExternalStore(
    store.subscribe,
    () => store.getState().presentation?.slides.length ?? 0,
    () => 0,
  );

  React.useEffect(() => {
    if (status !== "ready" || !presentation) return;
    window.__slideCount = slideCount;
    window.__benchRenderModes = () => ({
      slide: timeRenders(presentation, (p, s) => renderSlide(p, s, {})),
      thumbnail: timeRenders(presentation, (p, s) => renderThumbnail(p, s, {})),
    });
    window.__thumbnailsReady = slideCount >= targetSlides;
  }, [status, presentation, slideCount]);

  return (
    <Presentation.ThumbnailList
      id="thumbnail-list"
      style={{
        width: `${width}px`,
        height: `${height}px`,
        overflowY: "auto",
        display: "flex",
        flexDirection: "column",
        gap: "8px",
        padding: "8px",
        boxSizing: "border-box",
      }}
    />
  );
}

function App({ deck }: { deck: ArrayBuffer }) {
  const store = useCreatePresentationStore();

  return (
    <Presentation.Provider store={store}>
      <Presentation.Root
        file={deck}
        readOnly={false}
        onLoad={(loaded) => {
          growDeck(loaded).catch((error: unknown) => {
            window.__renderError =
              error instanceof Error ? error.message : String(error);
          });
        }}
        style={{ display: "flex" }}
      >
        <Harness store={store} />
      </Presentation.Root>
    </Presentation.Provider>
  );
}

async function main(): Promise<void> {
  const root = document.getElementById("thumbnail-root");
  if (!root) throw new Error("missing #thumbnail-root");

  if (!file) {
    window.__renderError = "missing ?file= query param";
    return;
  }

  try {
    const response = await fetch(`/${file}`);
    if (!response.ok) throw new Error(`fetch ${file}: HTTP ${response.status}`);
    const deck = await response.arrayBuffer();
    createRoot(root).render(<App deck={deck} />);
  } catch (error) {
    window.__renderError =
      error instanceof Error ? error.message : String(error);
  }
}

void main();
