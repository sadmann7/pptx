// @vitest-environment jsdom
/**
 * A drag overlay mounts a second copy of a thumbnail while the original is
 * still in the list. These tests lock the two things that makes possible: the
 * copy stays out of the list's semantics, and the two previews never fight
 * over one rendered slide element.
 *
 * Runs under jsdom (not the default happy-dom) for the same reason as
 * `thumbnail-reorder.test.tsx`: happy-dom's XML parser drops namespaced
 * attributes the renderer depends on.
 */

import { act, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";

import { Presentation } from "../index";
import type { Store } from "../store";
import { editableStore } from "./test-utils";

/**
 * jsdom ships no IntersectionObserver, and the preview only renders once it
 * reports the element as visible. This stub reports everything as visible
 * immediately, which is what the strip sees in a normal viewport.
 */
function stubIntersectionObserver(): void {
  vi.stubGlobal(
    "IntersectionObserver",
    class {
      constructor(private callback: IntersectionObserverCallback) {}
      observe(target: Element) {
        this.callback(
          [{ isIntersecting: true, target } as unknown as IntersectionObserverEntry],
          this as unknown as IntersectionObserver,
        );
      }
      unobserve() {}
      disconnect() {}
    },
  );
}

afterEach(() => {
  vi.unstubAllGlobals();
});

/**
 * Lets the list's rAF-batched render queue drain.
 *
 * The queue drains within an 8ms-per-frame budget, so a single callback
 * normally clears in one tick. A loaded CI runner can blow that budget on
 * the very first callback, pushing the rest to a second frame; a handful of
 * extra ticks costs nothing locally and removes that flake.
 */
async function flushPreviews(): Promise<void> {
  for (let tick = 0; tick < 5; tick++) {
    await act(async () => {
      await new Promise((resolve) => requestAnimationFrame(() => setTimeout(resolve, 0)));
    });
  }
}

function previewOf(item: Element): Element | null {
  return item.querySelector("[aria-hidden='true']");
}

interface StripProps {
  store: Store;
  /** Slide currently "being dragged", mirrored into an overlay copy. */
  overlaySlideId?: string;
}

function Strip({ store, overlaySlideId }: StripProps) {
  return (
    <Presentation.Provider store={store}>
      <Presentation.ThumbnailList>
        {({ slides }) => (
          <>
            {slides.map((slide) => (
              <Presentation.ThumbnailItem key={slide.id} slideId={slide.id} data-testid="item">
                <Presentation.ThumbnailItemNumber />
                <Presentation.ThumbnailItemPreview />
              </Presentation.ThumbnailItem>
            ))}
            {overlaySlideId ? (
              <Presentation.ThumbnailItem decorative slideId={overlaySlideId} data-testid="overlay">
                <Presentation.ThumbnailItemNumber />
                <Presentation.ThumbnailItemPreview />
              </Presentation.ThumbnailItem>
            ) : null}
          </>
        )}
      </Presentation.ThumbnailList>
    </Presentation.Provider>
  );
}

describe("decorative thumbnail items", () => {
  it("keeps the copy out of the listbox and out of the tab order", async () => {
    const store = await editableStore();
    const slides = store.getState().presentation!.slides;

    render(<Strip store={store} overlaySlideId={slides[0].id} />);

    // The copy renders, but the deck is still announced as N slides.
    expect(screen.getByTestId("overlay")).toBeTruthy();
    expect(screen.getAllByRole("option")).toHaveLength(slides.length);

    const overlay = screen.getByTestId("overlay");
    expect(overlay.getAttribute("role")).toBeNull();
    expect(overlay.getAttribute("aria-hidden")).toBe("true");
    expect(overlay.tabIndex).toBe(-1);
  });

  it("does not navigate when the copy is clicked or focused", async () => {
    const store = await editableStore();
    const slides = store.getState().presentation!.slides;
    store.goTo(slides[1].id);

    render(<Strip store={store} overlaySlideId={slides[0].id} />);

    act(() => {
      screen.getByTestId("overlay").dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(store.getState().activeSlideId).toBe(slides[1].id);
  });

  it("leaves the original registered for roving focus after the copy unmounts", async () => {
    const store = await editableStore();
    const slides = store.getState().presentation!.slides;

    const { rerender } = render(<Strip store={store} overlaySlideId={slides[0].id} />);
    rerender(<Strip store={store} />);

    const options = screen.getAllByRole("option");
    options[0].focus();
    act(() => {
      options[0].dispatchEvent(
        new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true, cancelable: true }),
      );
    });
    // Roving focus is scheduled a tick later so the browser finishes the keydown.
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 0));
    });

    // A copy that unregistered slide 0 on unmount would leave nothing to
    // arrow from, and focus would never reach slide 1.
    expect(document.activeElement).toBe(options[1]);
  });
});

describe("previews of the same slide in two places", () => {
  /**
   * Mirrors a real drag: the list renders and caches its miniatures first, and
   * only then does the overlay mount. That order matters, because the copy
   * hits a warm cache and the shared element is the one already on screen.
   */
  it("renders its own element rather than taking the one already on screen", async () => {
    stubIntersectionObserver();
    const store = await editableStore();
    const slides = store.getState().presentation!.slides;

    const { rerender } = render(<Strip store={store} />);
    await flushPreviews();

    const original = previewOf(screen.getAllByTestId("item")[0]);
    const cachedElement = original?.firstElementChild;
    expect(cachedElement).toBeTruthy();

    rerender(<Strip store={store} overlaySlideId={slides[0].id} />);
    await flushPreviews();

    const overlay = previewOf(screen.getByTestId("overlay"));
    expect(overlay?.firstElementChild).toBeTruthy();
    expect(overlay?.firstElementChild).not.toBe(cachedElement);
    // The list item keeps the element it was already showing.
    expect(original?.firstElementChild).toBe(cachedElement);
  });

  it("leaves the original rendered after the copy unmounts", async () => {
    stubIntersectionObserver();
    const store = await editableStore();
    const slides = store.getState().presentation!.slides;

    const { rerender } = render(<Strip store={store} />);
    await flushPreviews();

    rerender(<Strip store={store} overlaySlideId={slides[0].id} />);
    await flushPreviews();

    rerender(<Strip store={store} />);
    await flushPreviews();

    // The copy tearing down a shared element is what would blank this out.
    const original = previewOf(screen.getAllByTestId("item")[0]);
    expect(original?.firstElementChild).toBeTruthy();
    expect(original?.hasAttribute("data-pending")).toBe(false);
  });
});
