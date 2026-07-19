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
