import { act, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";

import { Presentation } from "../index";
import { createStore } from "../store";
import { distanceFromViewport, slideHasChart } from "../thumbnail-list";
import { FIXTURE_SLIDE_COUNT } from "./minimal-pptx";
import { loadedStore, withStore } from "./test-utils";

/**
 * Stand-in for IntersectionObserver, which happy-dom does not implement and
 * which never fires on its own here anyway: nothing lays out or scrolls.
 */
class FakeIntersectionObserver {
  static instances: FakeIntersectionObserver[] = [];

  private targets = new Set<Element>();

  constructor(private callback: IntersectionObserverCallback) {
    FakeIntersectionObserver.instances.push(this);
  }

  observe(target: Element) {
    this.targets.add(target);
  }

  unobserve(target: Element) {
    this.targets.delete(target);
  }

  disconnect() {
    this.targets.clear();
  }

  report(isIntersecting: boolean) {
    const entries = Array.from(this.targets, (target) => ({
      target,
      isIntersecting,
    })) as IntersectionObserverEntry[];
    this.callback(entries, this as unknown as IntersectionObserver);
  }

  static reportAll(isIntersecting: boolean) {
    for (const instance of FakeIntersectionObserver.instances) instance.report(isIntersecting);
  }
}

/** Renders with observation and frame scheduling under the test's control. */
function withObservedStore(...args: Parameters<typeof withStore>) {
  const frames: FrameRequestCallback[] = [];
  FakeIntersectionObserver.instances = [];

  vi.stubGlobal("IntersectionObserver", FakeIntersectionObserver);
  vi.stubGlobal("requestAnimationFrame", (callback: FrameRequestCallback) => frames.push(callback));
  vi.stubGlobal("cancelAnimationFrame", () => undefined);
  // Push the background pass onto its frame-callback fallback, so its work
  // lands in `frames` with everything else and the test decides when it runs.
  vi.stubGlobal("requestIdleCallback", undefined);

  const flushFrames = () => {
    // Draining is itself what queues the next frame, so loop until it settles.
    for (let pass = 0; pass < 10 && frames.length > 0; pass++) {
      const pending = frames.splice(0, frames.length);
      for (const callback of pending) callback(performance.now());
    }
  };

  return { ...withStore(...args), flushFrames };
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("Presentation.ThumbnailList", () => {
  it("renders a listbox with one option per slide and navigates on click", async () => {
    const store = await loadedStore();
    withStore(store, <Presentation.ThumbnailList />);

    expect(screen.getByRole("listbox")).toBeDefined();
    const options = screen.getAllByRole("option");
    expect(options).toHaveLength(FIXTURE_SLIDE_COUNT);
    expect(options[0].getAttribute("aria-selected")).toBe("true");
    expect(options[0].textContent).toBe("1");

    act(() => options[2].click());
    expect(store.getActiveSlideIndex()).toBe(2);
    expect(options[2].getAttribute("aria-selected")).toBe("true");
    expect(options[0].getAttribute("aria-selected")).toBe("false");
  });

  it("renders nothing before the presentation is ready", () => {
    const store = createStore();
    const { container } = withStore(store, <Presentation.ThumbnailList />);
    expect(container.querySelectorAll("button")).toHaveLength(0);
  });

  it("exposes slides and navigation through the children render function", async () => {
    const store = await loadedStore();
    withStore(
      store,
      <Presentation.ThumbnailList>
        {({ slides, activeIndex, goToIndex }) => (
          <>
            <span>count: {slides.length}</span>
            <span>active: {activeIndex}</span>
            <button type="button" onClick={() => goToIndex(1)}>
              jump
            </button>
          </>
        )}
      </Presentation.ThumbnailList>,
    );

    expect(screen.getByText(`count: ${FIXTURE_SLIDE_COUNT}`)).toBeDefined();
    expect(screen.getByText("active: 0")).toBeDefined();

    act(() => screen.getByText("jump").click());
    expect(store.getActiveSlideIndex()).toBe(1);
  });

  it("keeps a rendered preview once it leaves the observed area", async () => {
    const store = await loadedStore();
    const { container, flushFrames } = withObservedStore(store, <Presentation.ThumbnailList />);

    const preview = container.querySelector<HTMLElement>('[aria-hidden="true"]');
    expect(preview?.dataset.pending).toBe("");

    act(() => FakeIntersectionObserver.reportAll(true));
    act(() => flushFrames());

    expect(preview?.dataset.pending).toBeUndefined();
    expect(preview?.childElementCount).toBe(1);

    // The whole point of the retention model: scrolling away does not undo the
    // render, so scrolling back has no placeholder to show.
    act(() => FakeIntersectionObserver.reportAll(false));

    expect(preview?.dataset.pending).toBeUndefined();
    expect(preview?.childElementCount).toBe(1);
  });

  it("renders ahead of the observer so a preview attaches straight from cache", async () => {
    const store = await loadedStore();
    const { container, flushFrames } = withObservedStore(store, <Presentation.ThumbnailList />);
    const preview = container.querySelector<HTMLElement>('[aria-hidden="true"]');

    // Nothing has been observed yet, so this is the background pass alone. It
    // fills the list's cache; mounting still waits on the observer.
    act(() => flushFrames());
    expect(preview?.dataset.pending).toBe("");

    // A cache hit, so no frame has to be drained for the miniature to appear.
    act(() => FakeIntersectionObserver.reportAll(true));

    expect(preview?.dataset.pending).toBeUndefined();
    expect(preview?.childElementCount).toBe(1);
  });

  it("renders a preview only once while it stays observed", async () => {
    const store = await loadedStore();
    const { container, flushFrames } = withObservedStore(store, <Presentation.ThumbnailList />);
    const preview = container.querySelector<HTMLElement>('[aria-hidden="true"]');

    act(() => FakeIntersectionObserver.reportAll(true));
    act(() => flushFrames());
    const rendered = preview?.firstElementChild;

    act(() => FakeIntersectionObserver.reportAll(true));
    act(() => flushFrames());

    expect(preview?.firstElementChild).toBe(rendered);
    expect(preview?.childElementCount).toBe(1);
  });
});

describe("distanceFromViewport", () => {
  function elementAt(top: number, bottom: number): HTMLElement {
    const element = document.createElement("div");
    element.getBoundingClientRect = () => ({ top, bottom }) as DOMRect;
    Object.defineProperty(element.ownerDocument.defaultView, "innerHeight", {
      value: 800,
      configurable: true,
    });
    return element;
  }

  it("is zero for anything on screen, including partially", () => {
    expect(distanceFromViewport(elementAt(100, 300))).toBe(0);
    expect(distanceFromViewport(elementAt(-50, 20))).toBe(0);
    expect(distanceFromViewport(elementAt(790, 900))).toBe(0);
  });

  it("grows with the gap above and below the viewport", () => {
    expect(distanceFromViewport(elementAt(900, 1000))).toBe(100);
    expect(distanceFromViewport(elementAt(-300, -120))).toBe(120);
  });
});

describe("slideHasChart", () => {
  it("is true only for slides carrying a chart", async () => {
    const store = await loadedStore();
    const slides = store.getState().presentation?.slides ?? [];

    expect(slides.length).toBeGreaterThan(0);
    for (const slide of slides) {
      expect(slideHasChart(slide)).toBe(slide.nodes.some((n) => n.nodeType === "chart"));
    }
  });
});
