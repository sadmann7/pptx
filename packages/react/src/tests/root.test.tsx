import { act } from "react";

import { render, waitFor } from "@testing-library/react";
import { describe, expect, it } from "vitest";

import { useStoreContext } from "../context";
import { Presentation } from "../index";
import type { SlideChangeEvent, Store } from "../store";
import { createStore } from "../store";
import { loadFixture } from "./test-utils";

/** Captures the store visible at its position in the tree. */
function StoreProbe({ onStore }: { onStore: (store: Store) => void }) {
  onStore(useStoreContext("StoreProbe"));
  return null;
}

describe("Presentation.Root store resolution", () => {
  it("creates an internal store when no Provider is present", () => {
    let captured: Store | null = null;
    render(
      <Presentation.Root>
        <StoreProbe onStore={(s) => (captured = s)} />
      </Presentation.Root>,
    );
    expect(captured).not.toBeNull();
    expect(captured!.getState().status).toBe("idle");
  });

  it("keeps the same internal store across re-renders", () => {
    const seen = new Set<Store>();
    const ui = (
      <Presentation.Root>
        <StoreProbe onStore={(s) => seen.add(s)} />
      </Presentation.Root>
    );
    const { rerender } = render(ui);
    rerender(ui);
    expect(seen.size).toBe(1);
  });

  it("inherits the store from Presentation.Provider", () => {
    const store = createStore();
    let captured: Store | null = null;
    render(
      <Presentation.Provider store={store}>
        <Presentation.Root>
          <StoreProbe onStore={(s) => (captured = s)} />
        </Presentation.Root>
      </Presentation.Provider>,
    );
    expect(captured).toBe(store);
  });

  it("exposes the provider store to components outside Root", () => {
    const store = createStore();
    let outside: Store | null = null;
    let inside: Store | null = null;
    render(
      <Presentation.Provider store={store}>
        <StoreProbe onStore={(s) => (outside = s)} />
        <Presentation.Root>
          <StoreProbe onStore={(s) => (inside = s)} />
        </Presentation.Root>
      </Presentation.Provider>,
    );
    expect(outside).toBe(store);
    expect(inside).toBe(store);
  });
});

describe("Presentation.Root file prop", () => {
  it("loads the file into the internal store and calls onLoad", async () => {
    const fixture = await loadFixture();
    let loaded: Store | null = null;
    render(<Presentation.Root file={fixture} onLoad={(store) => (loaded = store)} />);

    await waitFor(() => expect(loaded).not.toBeNull());
    expect(loaded!.getState().status).toBe("ready");
    expect(loaded!.getState().presentation?.slides).toHaveLength(3);
  });

  it("calls onError for an invalid file", async () => {
    let error: Error | null = null;
    render(<Presentation.Root file={new ArrayBuffer(8)} onError={(err) => (error = err)} />);
    await waitFor(() => expect(error).not.toBeNull());
    expect(error).toBeInstanceOf(Error);
  });

  it("calls onSlideChange for the start slide and for later navigation", async () => {
    const fixture = await loadFixture();
    const store = createStore();
    const events: SlideChangeEvent[] = [];

    render(
      <Presentation.Provider store={store}>
        <Presentation.Root file={fixture} onSlideChange={(event) => events.push(event)} />
      </Presentation.Provider>,
    );

    await waitFor(() => expect(events).toHaveLength(1));
    expect(events[0]).toMatchObject({ index: 0, reason: "load", previousSlideId: null });

    act(() => store.next());
    expect(events).toHaveLength(2);
    expect(events[1]).toMatchObject({ index: 1, reason: "navigate" });
    expect(events[1].previousSlideId).toBe(events[0].slideId);
  });

  it("calls onStatusChange through the load lifecycle", async () => {
    const fixture = await loadFixture();
    const transitions: string[] = [];

    render(
      <Presentation.Root
        file={fixture}
        onStatusChange={({ status, previousStatus }) =>
          transitions.push(`${previousStatus}->${status}`)
        }
      />,
    );

    await waitFor(() => expect(transitions).toContain("loading->ready"));
    expect(transitions).toEqual(["idle->loading", "loading->ready"]);
  });

  it("opens at defaultZoom and keeps auto-fit off it", async () => {
    const fixture = await loadFixture();
    const store = createStore();

    render(
      <Presentation.Provider store={store}>
        <Presentation.Root file={fixture} defaultZoom={0.5}>
          <Presentation.Viewport autoFit />
        </Presentation.Root>
      </Presentation.Provider>,
    );

    await waitFor(() => expect(store.getState().status).toBe("ready"));
    expect(store.getState().zoom).toBe(0.5);
    expect(store.getState().zoomLevel).toBe(0.5);
  });

  it("does not touch a provider store when file is omitted", () => {
    const store = createStore();
    render(
      <Presentation.Provider store={store}>
        <Presentation.Root />
      </Presentation.Provider>,
    );
    expect(store.getState().status).toBe("idle");
  });

  it("loads through a provider store when file is passed", async () => {
    const fixture = await loadFixture();
    const store = createStore();
    render(
      <Presentation.Provider store={store}>
        <Presentation.Root file={fixture} />
      </Presentation.Provider>,
    );
    await waitFor(() => expect(store.getState().status).toBe("ready"));
    expect(store.getState().presentation?.slides).toHaveLength(3);
  });
});
