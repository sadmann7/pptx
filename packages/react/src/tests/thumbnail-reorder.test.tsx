// @vitest-environment jsdom
/**
 * Runs under jsdom because `moveSlide` rewrites `sldIdLst`, and happy-dom's XML
 * parser drops its `r:id`.
 */

import { act, render, screen } from "@testing-library/react";
import { describe, expect, it } from "vitest";

import { Presentation } from "../index";
import { editableStore } from "./test-utils";

function renderedOrder(): (string | null)[] {
  return screen
    .getAllByRole("option")
    .map((option) => option.getAttribute("data-slide-id"));
}

describe("consumer-driven reordering", () => {
  it("composes a consumer's ref, handlers, and style onto the item", async () => {
    const store = await editableStore();
    const slides = store.getState().presentation!.slides;
    const captured: (HTMLButtonElement | null)[] = [];
    let pointerDowns = 0;

    render(
      <Presentation.Provider store={store}>
        <Presentation.ThumbnailList>
          {({ slides: renderedSlides }) =>
            renderedSlides.map((slide) => (
              <Presentation.ThumbnailItem
                key={slide.id}
                slideId={slide.id}
                ref={(element) => {
                  captured.push(element);
                }}
                onPointerDown={() => pointerDowns++}
                style={{ transform: "translateY(8px)" }}
                data-draggable="true"
              >
                <Presentation.ThumbnailItemNumber />
              </Presentation.ThumbnailItem>
            ))
          }
        </Presentation.ThumbnailList>
      </Presentation.Provider>,
    );

    const options = screen.getAllByRole("option");
    expect(options).toHaveLength(slides.length);

    // A sortable library needs the element to attach its sensors to.
    expect(captured.filter(Boolean).length).toBe(slides.length);

    expect(options[0].getAttribute("data-draggable")).toBe("true");
    expect(options[0].getAttribute("role")).toBe("option");

    // Sortables animate through the transform style.
    expect(options[0].style.transform).toBe("translateY(8px)");
    expect(options[0].style.width).toBe("100%");

    act(() => {
      options[0].dispatchEvent(
        new MouseEvent("pointerdown", { bubbles: true }),
      );
    });
    expect(pointerDowns).toBe(1);
  });

  it("re-renders the strip in the new order after a moveSlide edit", async () => {
    const store = await editableStore();
    const slides = store.getState().presentation!.slides;
    const [first, second, third] = [slides[0].id, slides[1].id, slides[2].id];

    render(
      <Presentation.Provider store={store}>
        <Presentation.ThumbnailList />
      </Presentation.Provider>,
    );

    expect(renderedOrder()).toEqual([first, second, third]);

    await act(async () => {
      await store.edit({ type: "moveSlide", slideId: first, toIndex: 2 });
    });

    expect(renderedOrder()).toEqual([second, third, first]);

    await act(async () => {
      await store.undo();
    });
    expect(renderedOrder()).toEqual([first, second, third]);
  });

  it("keeps the item numbers in step with the new positions", async () => {
    const store = await editableStore();
    // The slides array is reordered in place, so hold the id, not the index.
    const movedId = store.getState().presentation!.slides[0].id;

    render(
      <Presentation.Provider store={store}>
        <Presentation.ThumbnailList />
      </Presentation.Provider>,
    );

    expect(
      screen.getAllByRole("option").map((option) => option.textContent),
    ).toEqual(["1", "2", "3"]);

    await act(async () => {
      await store.edit({ type: "moveSlide", slideId: movedId, toIndex: 2 });
    });

    const options = screen.getAllByRole("option");
    expect(options.map((option) => option.textContent)).toEqual([
      "1",
      "2",
      "3",
    ]);
    expect(options[2].getAttribute("data-slide-id")).toBe(movedId);
  });

  it("hands the roving tab stop to the slide an undo jumps to", async () => {
    const store = await editableStore();
    const slides = store.getState().presentation!.slides;
    const [first, third] = [slides[0].id, slides[2].id];

    render(
      <Presentation.Provider store={store}>
        <Presentation.ThumbnailList />
      </Presentation.Provider>,
    );

    act(() => screen.getAllByRole("option")[2].focus());
    expect(store.getState().activeSlideId).toBe(third);

    await act(async () => {
      await store.edit({ type: "moveSlide", slideId: first, toIndex: 1 });
    });
    expect(store.getState().activeSlideId).toBe(third);

    await act(async () => {
      await store.undo();
    });

    // Undo navigated to the slide it touched, so the tab stop and the focus
    // riding on it follow instead of staying behind on the pressed thumbnail.
    expect(store.getState().activeSlideId).toBe(first);
    const undoneItem = screen
      .getAllByRole("option")
      .find((option) => option.getAttribute("data-slide-id") === first);
    expect(undoneItem?.tabIndex).toBe(0);
    expect(document.activeElement).toBe(undoneItem);
  });

  it("lets onSelect suppress navigation while a drag is in progress", async () => {
    const store = await editableStore();
    const slides = store.getState().presentation!.slides;
    let isDragging = false;

    render(
      <Presentation.Provider store={store}>
        <Presentation.ThumbnailList>
          {({ slides: renderedSlides }) =>
            renderedSlides.map((slide) => (
              <Presentation.ThumbnailItem
                key={slide.id}
                slideId={slide.id}
                onSelect={(event) => {
                  if (isDragging) event.preventDefault();
                }}
              >
                <Presentation.ThumbnailItemNumber />
              </Presentation.ThumbnailItem>
            ))
          }
        </Presentation.ThumbnailList>
      </Presentation.Provider>,
    );

    const options = screen.getAllByRole("option");

    isDragging = true;
    act(() => options[2].click());
    expect(store.getState().activeSlideId).toBe(slides[0].id);

    isDragging = false;
    act(() => options[2].click());
    expect(store.getState().activeSlideId).toBe(slides[2].id);
  });
});
