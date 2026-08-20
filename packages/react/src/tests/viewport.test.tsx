import * as React from "react";

import { act, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";

import { Presentation } from "../index";
import type { ZoomChangeEvent } from "../store";
import { loadedStore, withStore } from "./test-utils";

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

  it("calls onZoomChange with the new and previous zoom", async () => {
    const store = await loadedStore();
    const events: ZoomChangeEvent[] = [];

    withStore(store, <Presentation.Viewport onZoomChange={(event) => events.push(event)} />);

    act(() => store.setZoom(2));
    expect(events).toEqual([{ zoom: 2, previousZoom: 1, reason: "zoom" }]);

    act(() => store.zoomOut(0.5));
    expect(events.at(-1)).toEqual({ zoom: 1.5, previousZoom: 2, reason: "zoom" });
  });

  describe("autoFit", () => {
    /**
     * The test DOM reports a zero-sized element and never lays anything out,
     * so the container size and the resize notification both have to be
     * supplied by hand to exercise the fit path at all.
     */
    function stubResizeObserver(): () => void {
      const callbacks = new Set<ResizeObserverCallback>();
      vi.stubGlobal(
        "ResizeObserver",
        class {
          constructor(callback: ResizeObserverCallback) {
            callbacks.add(callback);
          }
          observe() {}
          unobserve() {}
          disconnect() {}
        },
      );
      return () => {
        for (const callback of callbacks) callback([], {} as ResizeObserver);
      };
    }

    function stubSize(element: HTMLElement, width: number, height: number): void {
      Object.defineProperty(element, "clientWidth", { value: width, configurable: true });
      Object.defineProperty(element, "clientHeight", { value: height, configurable: true });
    }

    afterEach(() => {
      vi.unstubAllGlobals();
    });

    it('fits to the container while the level is "fit"', async () => {
      const store = await loadedStore();
      const notifyResize = stubResizeObserver();
      withStore(store, <Presentation.Viewport autoFit data-testid="viewport" />);

      stubSize(screen.getByTestId("viewport"), 640, 360);
      act(() => notifyResize());

      expect(store.getState().zoom).toBe(0.5);
      expect(store.getState().zoomLevel).toBe("fit");
    });

    it("leaves a pinned level alone when the container resizes", async () => {
      const store = await loadedStore();
      const notifyResize = stubResizeObserver();
      withStore(store, <Presentation.Viewport autoFit data-testid="viewport" />);

      stubSize(screen.getByTestId("viewport"), 640, 360);
      act(() => store.setZoom(2));
      act(() => notifyResize());

      expect(store.getState().zoom).toBe(2);
    });

    it('resolves setZoom("fit") against the live container size', async () => {
      const store = await loadedStore();
      stubResizeObserver();
      withStore(store, <Presentation.Viewport autoFit data-testid="viewport" />);

      stubSize(screen.getByTestId("viewport"), 640, 360);
      act(() => store.setZoom(2));
      act(() => store.setZoom("fit"));

      expect(store.getState().zoom).toBe(0.5);
      expect(store.getState().zoomLevel).toBe("fit");
    });
  });

  describe("scrollNavigation", () => {
    function wheel(element: HTMLElement, deltaY: number, timeStamp?: number): void {
      const event = new WheelEvent("wheel", { deltaY, bubbles: true, cancelable: true });
      // WheelEvent timestamps are read-only; override to control the cooldown.
      if (timeStamp !== undefined) {
        Object.defineProperty(event, "timeStamp", { value: timeStamp });
      }
      act(() => {
        element.dispatchEvent(event);
      });
    }

    it("advances to the next slide on wheel down when there is nothing to scroll", async () => {
      const store = await loadedStore();
      withStore(store, <Presentation.Viewport scrollNavigation data-testid="viewport" />);

      wheel(screen.getByTestId("viewport"), 100);
      expect(store.getActiveSlideIndex()).toBe(1);
    });

    it("swallows momentum ticks within the cooldown window", async () => {
      const store = await loadedStore();
      withStore(store, <Presentation.Viewport scrollNavigation data-testid="viewport" />);
      const viewport = screen.getByTestId("viewport");

      wheel(viewport, 100, 1000);
      wheel(viewport, 100, 1100);
      expect(store.getActiveSlideIndex()).toBe(1);

      // Past the cooldown, the next tick navigates again.
      wheel(viewport, 100, 2000);
      expect(store.getActiveSlideIndex()).toBe(2);
    });

    it("goes back on wheel up and no-ops at the first slide", async () => {
      const store = await loadedStore();
      withStore(store, <Presentation.Viewport scrollNavigation data-testid="viewport" />);
      const viewport = screen.getByTestId("viewport");

      wheel(viewport, -100, 1000);
      expect(store.getActiveSlideIndex()).toBe(0);

      wheel(viewport, 100, 2000);
      wheel(viewport, -100, 3000);
      expect(store.getActiveSlideIndex()).toBe(0);
    });

    it("ignores ctrl+wheel (pinch-zoom gesture)", async () => {
      const store = await loadedStore();
      withStore(store, <Presentation.Viewport scrollNavigation data-testid="viewport" />);

      const event = new WheelEvent("wheel", { deltaY: 100, bubbles: true, cancelable: true });
      // happy-dom's WheelEvent constructor drops modifier keys from the init.
      Object.defineProperty(event, "ctrlKey", { value: true });
      act(() => {
        screen.getByTestId("viewport").dispatchEvent(event);
      });
      expect(store.getActiveSlideIndex()).toBe(0);
    });

    it("ignores shift+wheel (horizontal scroll gesture)", async () => {
      const store = await loadedStore();
      withStore(store, <Presentation.Viewport scrollNavigation data-testid="viewport" />);

      const event = new WheelEvent("wheel", { deltaY: 100, bubbles: true, cancelable: true });
      // happy-dom's WheelEvent constructor drops modifier keys from the init.
      Object.defineProperty(event, "shiftKey", { value: true });
      act(() => {
        screen.getByTestId("viewport").dispatchEvent(event);
      });
      expect(store.getActiveSlideIndex()).toBe(0);
    });

    it("does not navigate when the prop is off", async () => {
      const store = await loadedStore();
      withStore(store, <Presentation.Viewport data-testid="viewport" />);

      wheel(screen.getByTestId("viewport"), 100);
      expect(store.getActiveSlideIndex()).toBe(0);
    });
  });
});
