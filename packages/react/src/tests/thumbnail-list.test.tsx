import { act, screen } from "@testing-library/react";
import { describe, expect, it } from "vitest";

import { Presentation } from "../index";
import { createStore } from "../store";
import { FIXTURE_SLIDE_COUNT } from "./minimal-pptx";
import { loadedStore, withStore } from "./test-utils";

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

    // Focus is the other route into navigation and must respect the veto too.
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

    // Focus navigates too, which is what makes the veto test's focus case meaningful.
    act(() => options[1].focus());
    expect(store.getActiveSlideIndex()).toBe(1);
  });

  it("leaves focus alone when a deck loads", async () => {
    // A deck can arrive without being asked for (fetched on mount, restored on
    // navigation), so capturing focus by default would take the page's tab
    // position away from wherever it belongs.
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

    withStore(store, <Presentation.ThumbnailList initialFocus={{ current: outside }} />);
    expect(document.activeElement).toBe(outside);

    outside.remove();
  });

  it("falls back to the active thumbnail when the named element is absent", async () => {
    // A ref that has not attached, or a resolver returning null, is a focus
    // request that cannot name its target, not a request to stay put.
    const store = await loadedStore();
    withStore(store, <Presentation.ThumbnailList initialFocus={{ current: null }} />);

    expect(document.activeElement).toBe(screen.getAllByRole("option")[0]);
  });

  it("lets an initialFocus function decide per load", async () => {
    // The shape a page with both an on-mount fetch and a file input needs: the
    // consumer knows which load the user asked for, and this one did not.
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
});
