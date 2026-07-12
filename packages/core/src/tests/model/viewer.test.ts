import { afterEach, describe, expect, it } from "vitest";

import { normalizePptxSource, PptxViewer } from "../../viewer";
import { buildCustomPptx } from "../fixtures/fixture-extras";
import { buildPptxWithShapes } from "../fixtures/minimal-pptx";

function textShape(id: number, text: string): string {
  return `<p:sp>
<p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="4572000" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody>
</p:sp>`;
}

function threeSlideDeck(): Promise<ArrayBuffer> {
  return buildCustomPptx({
    slides: [textShape(2, "Alpha slide"), textShape(2, "Beta slide"), textShape(2, "Gamma slide")],
  });
}

function makeContainer(): HTMLElement {
  const container = document.createElement("div");
  document.body.appendChild(container);
  return container;
}

const cleanups: (() => void)[] = [];

afterEach(() => {
  for (const cleanup of cleanups.splice(0)) cleanup();
  document.body.innerHTML = "";
});

function track(viewer: PptxViewer): PptxViewer {
  cleanups.push(() => viewer.destroy());
  return viewer;
}

describe("PptxViewer.open (list mode)", () => {
  it("parses, loads, and mounts every slide into the container", async () => {
    const container = makeContainer();
    const viewer = track(await PptxViewer.open(await threeSlideDeck(), container));

    expect(viewer.slideCount).toBe(3);
    expect(viewer.slideWidth).toBeCloseTo(1280, 6);
    expect(viewer.slideHeight).toBeCloseTo(720, 6);
    expect(viewer.presentationData).not.toBeNull();

    const items = container.querySelectorAll("[data-slide-index]");
    expect(items).toHaveLength(3);
    expect(container.textContent).toContain("Alpha slide");
    expect(container.textContent).toContain("Gamma slide");

    expect(viewer.getMountedSlides()).toEqual([0, 1, 2]);
    expect(viewer.isSlideMounted(0)).toBe(true);
    expect(viewer.isSlideMounted(5)).toBe(false);
  });

  it("scales slides to the configured width", async () => {
    const container = makeContainer();
    const viewer = track(await PptxViewer.open(await threeSlideDeck(), container, { width: 640 }));
    expect(viewer.fitMode).toBe("contain");

    const wrapper = container.querySelector<HTMLElement>("[data-slide-index='0']")
      ?.firstElementChild as HTMLElement;
    expect(wrapper.style.width).toBe("640px");
    expect(wrapper.style.height).toBe("360px");

    const slideEl = wrapper.firstElementChild as HTMLElement;
    expect(slideEl.style.transform).toBe("scale(0.5)");
  });

  it("re-renders at the new scale after setZoom", async () => {
    const container = makeContainer();
    const viewer = track(await PptxViewer.open(await threeSlideDeck(), container, { width: 640 }));
    await viewer.setZoom(200);
    expect(viewer.zoomPercent).toBe(200);

    const wrapper = container.querySelector<HTMLElement>("[data-slide-index='0']")
      ?.firstElementChild as HTMLElement;
    expect(wrapper.style.width).toBe("1280px");
    expect((wrapper.firstElementChild as HTMLElement).style.transform).toBe("scale(1)");
  });

  it("clamps zoom to the 10–400 percent range", async () => {
    const container = makeContainer();
    const viewer = track(await PptxViewer.open(await threeSlideDeck(), container));
    await viewer.setZoom(5000);
    expect(viewer.zoomPercent).toBe(400);
    await viewer.setZoom(1);
    expect(viewer.zoomPercent).toBe(10);
    await viewer.setZoom(Number.NaN);
    expect(viewer.zoomPercent).toBe(100);
  });

  it("shows slide labels when requested", async () => {
    const container = makeContainer();
    track(
      await PptxViewer.open(await threeSlideDeck(), container, {
        listOptions: { showSlideLabels: true },
      }),
    );
    expect(container.textContent).toContain("Slide 1");
    expect(container.textContent).toContain("Slide 3");
  });

  it("fires renderstart/rendercomplete/sliderendered/slidechange events", async () => {
    const container = makeContainer();
    const events: string[] = [];
    const renderedIndexes: number[] = [];

    const viewer = track(
      new PptxViewer(container, {
        onRenderStart: () => events.push("start"),
        onRenderComplete: () => events.push("complete"),
        onSlideRendered: (index) => renderedIndexes.push(index),
        onSlideChange: (index) => events.push(`change:${index}`),
      }),
    );
    await viewer.open(await threeSlideDeck());

    expect(events[0]).toBe("start");
    expect(events).toContain("complete");
    expect(events).toContain("change:0");
    expect(renderedIndexes).toEqual([0, 1, 2]);
    expect(viewer.isRendering).toBe(false);
  });

  it("supports typed on/off event helpers", async () => {
    const container = makeContainer();
    const viewer = track(new PptxViewer(container));
    let renderedCount = 0;
    const listener = (): void => {
      renderedCount++;
    };
    viewer.on("sliderendered", listener);
    await viewer.open(await threeSlideDeck());
    expect(renderedCount).toBe(3);

    viewer.off("sliderendered", listener);
    await viewer.renderList();
    expect(renderedCount).toBe(3);
  });
});

describe("PptxViewer slide mode and navigation", () => {
  it("renders a single slide and navigates with goToSlide", async () => {
    const container = makeContainer();
    const viewer = track(
      await PptxViewer.open(await threeSlideDeck(), container, { renderMode: "slide" }),
    );

    expect(viewer.currentSlideIndex).toBe(0);
    expect(container.textContent).toContain("Alpha slide");
    expect(container.textContent).not.toContain("Beta slide");

    const changes: number[] = [];
    viewer.on("slidechange", (e) => changes.push(e.detail.index));

    await viewer.goToSlide(1);
    expect(viewer.currentSlideIndex).toBe(1);
    expect(container.textContent).toContain("Beta slide");
    expect(container.textContent).not.toContain("Alpha slide");
    expect(changes).toEqual([1]);
  });

  it("clamps goToSlide to the valid slide range", async () => {
    const container = makeContainer();
    const viewer = track(
      await PptxViewer.open(await threeSlideDeck(), container, { renderMode: "slide" }),
    );
    await viewer.goToSlide(99);
    expect(viewer.currentSlideIndex).toBe(2);
    await viewer.goToSlide(-5);
    expect(viewer.currentSlideIndex).toBe(0);
  });
});

describe("PptxViewer external rendering and search", () => {
  it("renders a slide into an external container with a disposable handle", async () => {
    const container = makeContainer();
    const viewer = track(await PptxViewer.open(await threeSlideDeck(), container));

    const external = makeContainer();
    const handle = viewer.renderSlideToContainer(1, external, 0.5);
    expect(handle).not.toBeNull();
    expect(external.textContent).toContain("Beta slide");
    expect(handle!.element.style.transform).toBe("scale(0.5)");

    // Unlike thumbnail handles, slide handles only release resources; the
    // caller owns DOM removal.
    handle!.dispose();
    expect(external.contains(handle!.element)).toBe(true);
    handle!.element.remove();
    expect(external.textContent).not.toContain("Beta slide");

    expect(viewer.renderSlideToContainer(99, external)).toBeNull();
  });

  it("renders sized thumbnails preserving the aspect ratio", async () => {
    const container = makeContainer();
    const viewer = track(await PptxViewer.open(await threeSlideDeck(), container));

    const external = makeContainer();
    const handle = viewer.renderThumbnailToContainer(0, external, { width: 320 });
    expect(handle).not.toBeNull();
    expect(handle!.element.dataset.pptxThumbnail).toBe("true");
    expect(handle!.element.style.width).toBe("320px");
    expect(handle!.element.style.height).toBe("180px");

    handle!.dispose();
    expect(external.querySelector("[data-pptx-thumbnail]")).toBeNull();
  });

  it("searches loaded text and highlights results", async () => {
    const container = makeContainer();
    const viewer = track(await PptxViewer.open(await threeSlideDeck(), container));

    const results = viewer.searchText("Beta");
    expect(results).toHaveLength(1);
    expect(results[0].slideIndex).toBe(1);

    const highlight = await viewer.highlightSearchResult(results[0], { scrollIntoView: false });
    expect(highlight).not.toBeNull();
    expect(highlight!.element.dataset.pptxSearchHighlight).toBe("true");
    expect(container.contains(highlight!.element)).toBe(true);

    viewer.clearSearchHighlights();
    expect(container.querySelector("[data-pptx-search-highlight]")).toBeNull();
  });

  it("returns no results for text that is not present", async () => {
    const container = makeContainer();
    const viewer = track(await PptxViewer.open(await threeSlideDeck(), container));
    expect(viewer.searchText("Missing")).toHaveLength(0);
  });
});

describe("PptxViewer destroy", () => {
  it("clears the container and presentation state", async () => {
    const container = makeContainer();
    const viewer = await PptxViewer.open(await threeSlideDeck(), container);

    viewer.destroy();
    expect(container.innerHTML).toBe("");
    expect(viewer.presentationData).toBeNull();
    expect(viewer.slideCount).toBe(0);
    expect(viewer.getMountedSlides()).toEqual([]);
  });

  it("can reopen after destroy", async () => {
    const container = makeContainer();
    const viewer = track(await PptxViewer.open(await threeSlideDeck(), container));
    viewer.destroy();

    await viewer.open(await buildPptxWithShapes(textShape(2, "Reopened")));
    expect(viewer.slideCount).toBe(1);
    expect(container.textContent).toContain("Reopened");
  });
});

describe("normalizePptxSource", () => {
  it("passes ArrayBuffers through unchanged", async () => {
    const buffer = new ArrayBuffer(8);
    await expect(normalizePptxSource(buffer)).resolves.toBe(buffer);
  });

  it("copies Uint8Array views into standalone buffers", async () => {
    const backing = new Uint8Array([1, 2, 3, 4, 5, 6]);
    const view = backing.subarray(2, 5);
    const result = await normalizePptxSource(view);
    expect(result.byteLength).toBe(3);
    expect(new Uint8Array(result)).toEqual(new Uint8Array([3, 4, 5]));
  });

  it("reads Blobs via arrayBuffer()", async () => {
    const blob = new Blob([new Uint8Array([7, 8, 9])]);
    const result = await normalizePptxSource(blob);
    expect(new Uint8Array(result)).toEqual(new Uint8Array([7, 8, 9]));
  });
});
