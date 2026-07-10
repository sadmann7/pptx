import * as React from "react";

import { act, render, screen, waitFor } from "@testing-library/react";
import { beforeAll, describe, expect, it } from "vitest";

import { Presentation } from "../index";
import type { Store } from "../store";
import { createStore } from "../store";
import { buildMinimalPptx, FIXTURE_SLIDE_COUNT } from "./minimal-pptx";

let fixture: ArrayBuffer;

beforeAll(async () => {
  fixture = await buildMinimalPptx();
});

async function loadedStore(): Promise<Store> {
  const store = createStore();
  await store.load(fixture);
  return store;
}

function withStore(store: Store, ui: React.ReactNode) {
  return render(<Presentation.Provider store={store}>{ui}</Presentation.Provider>);
}

describe("Presentation.Loading", () => {
  it("renders only while loading, with progress via render function", async () => {
    const store = createStore();
    // Kick off a load without awaiting: status flips to "loading" synchronously.
    const pending = store.load(fixture);

    withStore(store, <Presentation.Loading>{(p) => <span>at {p}%</span>}</Presentation.Loading>);
    expect(screen.getByText(/at \d+%/)).toBeDefined();

    await act(async () => {
      await pending;
    });
    expect(screen.queryByText(/at \d+%/)).toBeNull();
  });

  it("renders nothing when idle", () => {
    const store = createStore();
    const { container } = withStore(store, <Presentation.Loading>busy</Presentation.Loading>);
    expect(container.textContent).toBe("");
  });
});

describe("Presentation.Error", () => {
  it("renders the error via render function after a failed load", async () => {
    const store = createStore();
    await store.load(new ArrayBuffer(8)).catch(() => undefined);

    withStore(
      store,
      <Presentation.Error>{(err: Error) => <span>failed: {err.message}</span>}</Presentation.Error>,
    );
    expect(screen.getByText(/^failed: /).textContent).toContain("failed:");
  });

  it("renders nothing when there is no error", async () => {
    const store = await loadedStore();
    const { container } = withStore(store, <Presentation.Error>boom</Presentation.Error>);
    expect(container.textContent).toBe("");
  });
});

describe("Presentation.Slide", () => {
  it("mounts the wrapper with data-status and renders slide content when ready", async () => {
    const store = await loadedStore();
    const { container } = withStore(store, <Presentation.Slide data-testid="slide" />);

    const wrapper = screen.getByTestId("slide");
    expect(wrapper.getAttribute("data-status")).toBe("ready");
    // The parser's slide DOM is appended inside the scaled container.
    await waitFor(() => {
      expect(
        container.querySelector("[data-slide-root], svg, .pptx-slide, div div div"),
      ).not.toBeNull();
    });
  });

  it("keeps the wrapper mounted with data-status=idle before load", () => {
    const store = createStore();
    withStore(store, <Presentation.Slide data-testid="slide" />);
    expect(screen.getByTestId("slide").getAttribute("data-status")).toBe("idle");
  });
});

describe("Presentation.Viewport", () => {
  it("renders children and supports render-prop replacement with state", async () => {
    const store = await loadedStore();
    store.setZoom(2);

    withStore(
      store,
      <Presentation.Viewport
        data-testid="viewport"
        render={(props, state) => (
          <section {...props} data-zoom={state.zoom}>
            {props.children}
          </section>
        )}
      >
        <span>inside</span>
      </Presentation.Viewport>,
    );

    const viewport = screen.getByTestId("viewport");
    expect(viewport.tagName).toBe("SECTION");
    expect(viewport.getAttribute("data-zoom")).toBe("2");
    expect(screen.getByText("inside")).toBeDefined();
  });
});

describe("Presentation.ThumbnailList", () => {
  it("renders a listbox with one option per slide and navigates on click", async () => {
    const store = await loadedStore();
    withStore(store, <Presentation.ThumbnailList />);

    expect(screen.getByRole("listbox")).toBeDefined();
    const options = screen.getAllByRole("option");
    expect(options).toHaveLength(FIXTURE_SLIDE_COUNT);
    expect(options[0].getAttribute("aria-selected")).toBe("true");
    expect(options[0].getAttribute("aria-label")).toBe("Slide 1");

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
