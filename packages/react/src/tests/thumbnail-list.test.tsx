import * as React from "react";

import { act, screen } from "@testing-library/react";
import { describe, expect, it } from "vitest";

import { Presentation } from "../index";
import { createStore } from "../store";
import { FIXTURE_SLIDE_COUNT } from "./minimal-pptx";
import { editableStore, loadedStore, withStore } from "./test-utils";

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

  it("lets onSelect veto navigation, from focus as well as click", async () => {
    const store = await loadedStore();
    const selected: string[] = [];
    withStore(
      store,
      <Presentation.ThumbnailList>
        {({ slides }) =>
          slides.map((slide) => (
            <Presentation.ThumbnailItem
              key={slide.id}
              slideId={slide.id}
              onSelect={(event) => {
                selected.push(event.slideId);
                event.preventDefault();
              }}
            >
              <Presentation.ThumbnailItemNumber />
            </Presentation.ThumbnailItem>
          ))
        }
      </Presentation.ThumbnailList>,
    );

    const options = screen.getAllByRole("option");

    act(() => options[2].click());
    expect(selected).not.toHaveLength(0);
    expect(store.getActiveSlideIndex()).toBe(0);

    act(() => options[1].focus());
    expect(store.getActiveSlideIndex()).toBe(0);
  });

  it("navigates as usual when onSelect does not prevent", async () => {
    const store = await loadedStore();
    let calls = 0;
    withStore(
      store,
      <Presentation.ThumbnailList>
        {({ slides }) =>
          slides.map((slide) => (
            <Presentation.ThumbnailItem
              key={slide.id}
              slideId={slide.id}
              onSelect={() => {
                calls++;
              }}
            >
              <Presentation.ThumbnailItemNumber />
            </Presentation.ThumbnailItem>
          ))
        }
      </Presentation.ThumbnailList>,
    );

    const options = screen.getAllByRole("option");

    act(() => options[2].click());
    expect(calls).toBeGreaterThan(0);
    expect(store.getActiveSlideIndex()).toBe(2);

    act(() => options[1].focus());
    expect(store.getActiveSlideIndex()).toBe(1);
  });

  it("leaves focus alone when a deck loads", async () => {
    // A deck can arrive unasked, fetched on mount or restored on navigation,
    // so taking focus by default would steal the page's tab position.
    const store = await loadedStore();
    withStore(store, <Presentation.ThumbnailList />);

    expect(screen.getAllByRole("option")).toHaveLength(FIXTURE_SLIDE_COUNT);
    expect(document.activeElement).toBe(document.body);
  });

  it("focuses the active thumbnail when initialFocus asks for it", async () => {
    const store = await loadedStore();
    withStore(store, <Presentation.ThumbnailList initialFocus />);

    const options = screen.getAllByRole("option");
    expect(document.activeElement).toBe(options[0]);
  });

  it("focuses the element initialFocus names, by ref or by return value", async () => {
    const store = await loadedStore();
    const outside = document.createElement("button");
    document.body.appendChild(outside);

    withStore(
      store,
      <Presentation.ThumbnailList initialFocus={{ current: outside }} />,
    );
    expect(document.activeElement).toBe(outside);

    outside.remove();
  });

  it("falls back to the active thumbnail when the named element is absent", async () => {
    // A ref that has not attached, or a resolver returning null, is a focus
    // request that cannot name its target, not a request to stay put.
    const store = await loadedStore();
    withStore(
      store,
      <Presentation.ThumbnailList initialFocus={{ current: null }} />,
    );

    expect(document.activeElement).toBe(screen.getAllByRole("option")[0]);
  });

  it("lets an initialFocus function decide per load", async () => {
    const store = await loadedStore();
    let asked = 0;
    withStore(
      store,
      <Presentation.ThumbnailList
        initialFocus={() => {
          asked += 1;
          return false;
        }}
      />,
    );

    expect(asked).toBe(1);
    expect(document.activeElement).toBe(document.body);
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

  it("re-renders only the previews a navigation or an edit touches", async () => {
    const store = await editableStore();
    const slides = store.getState().presentation!.slides;
    const renders = new Map<string, number>();

    function CountedPreview({
      slideId,
      ...props
    }: React.ComponentProps<"div"> & { slideId?: string }) {
      if (slideId) renders.set(slideId, (renders.get(slideId) ?? 0) + 1);
      return <div {...props} />;
    }
    const previews = new Map(
      slides.map((slide) => [
        slide.id,
        <CountedPreview key={slide.id} slideId={slide.id} />,
      ]),
    );

    withStore(
      store,
      <Presentation.ThumbnailList>
        {({ slides: listed }) =>
          listed.map((slide) => (
            <Presentation.ThumbnailItem key={slide.id} slideId={slide.id}>
              <Presentation.ThumbnailItemPreview
                render={previews.get(slide.id)}
              />
            </Presentation.ThumbnailItem>
          ))
        }
      </Presentation.ThumbnailList>,
    );

    renders.clear();
    act(() => store.goToIndex(2));
    expect([...renders.keys()].sort()).toEqual(
      [slides[0].id, slides[2].id].sort(),
    );

    renders.clear();
    const node = slides[2].nodes[0];
    await act(() =>
      store.edit({
        type: "setNodeTransform",
        slideId: slides[2].id,
        nodeId: node.id,
        position: { x: node.position.x + 10, y: node.position.y },
      }),
    );
    expect([...renders.keys()]).toEqual([slides[2].id]);
  });
});
