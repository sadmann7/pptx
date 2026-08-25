import { act, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";

import { Presentation } from "../index";
import type { ZoomChangeEvent } from "../store";
import { MAX_ZOOM } from "../store";
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

    afterEach(() => {
      vi.unstubAllGlobals();
    });

    it("fits the slide to the container on resize", async () => {
      const store = await loadedStore();
      const notifyResize = stubResizeObserver();
      withStore(store, <Presentation.Viewport autoFit data-testid="viewport" />);

      const viewport = screen.getByTestId("viewport");
      Object.defineProperty(viewport, "clientWidth", { value: 640, configurable: true });
      Object.defineProperty(viewport, "clientHeight", { value: 360, configurable: true });
      act(() => notifyResize());

      // Fixture deck is 1280x720.
      expect(store.getState().zoom).toBe(0.5);
    });

    it("stops fitting on resize once an explicit level is picked", async () => {
      const store = await loadedStore();
      const notifyResize = stubResizeObserver();
      withStore(store, <Presentation.Viewport autoFit data-testid="viewport" />);

      const viewport = screen.getByTestId("viewport");
      Object.defineProperty(viewport, "clientWidth", { value: 640, configurable: true });
      Object.defineProperty(viewport, "clientHeight", { value: 360, configurable: true });

      act(() => store.setZoom(2));
      act(() => notifyResize());
      expect(store.getState().zoom).toBe(2);
    });

    it("refits when auto-fit is turned back on", async () => {
      const store = await loadedStore();
      stubResizeObserver();
      withStore(store, <Presentation.Viewport autoFit data-testid="viewport" />);

      const viewport = screen.getByTestId("viewport");
      Object.defineProperty(viewport, "clientWidth", { value: 640, configurable: true });
      Object.defineProperty(viewport, "clientHeight", { value: 360, configurable: true });

      act(() => store.setZoom(2));
      // No resize in between: re-arming fits against the container as it is.
      act(() => store.setAutoFit(true));
      expect(store.getState().zoom).toBe(0.5);
    });

    it("leaves zoom alone when the prop is off", async () => {
      const store = await loadedStore();
      const notifyResize = stubResizeObserver();
      withStore(store, <Presentation.Viewport data-testid="viewport" />);

      const viewport = screen.getByTestId("viewport");
      Object.defineProperty(viewport, "clientWidth", { value: 640, configurable: true });
      Object.defineProperty(viewport, "clientHeight", { value: 360, configurable: true });
      act(() => notifyResize());

      expect(store.getState().isAutoFit).toBe(false);
      expect(store.getState().zoom).toBe(1);
    });

    it("reports the fit as the reason", async () => {
      const store = await loadedStore();
      const reasons: string[] = [];
      withStore(
        store,
        <Presentation.Viewport onZoomChange={({ reason }) => reasons.push(reason)} />,
      );

      act(() => store.fitTo(640, 360));
      expect(reasons).toEqual(["fit"]);
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

    it("ignores ctrl/cmd+wheel (zoom gesture)", async () => {
      const store = await loadedStore();
      withStore(store, <Presentation.Viewport scrollNavigation data-testid="viewport" />);

      for (const modifier of ["ctrlKey", "metaKey"]) {
        const event = new WheelEvent("wheel", { deltaY: 100, bubbles: true, cancelable: true });
        // happy-dom's WheelEvent constructor drops modifier keys from the init.
        Object.defineProperty(event, modifier, { value: true });
        act(() => {
          screen.getByTestId("viewport").dispatchEvent(event);
        });
        expect(store.getActiveSlideIndex()).toBe(0);
      }
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

  describe("scrollZoom", () => {
    afterEach(() => {
      vi.unstubAllGlobals();
    });

    /**
     * Ctrl/Cmd+wheel, sending Ctrl for both since the handler treats them
     * alike. Assigned rather than passed: happy-dom's WheelEvent constructor
     * drops modifier keys from the init.
     */
    function zoomWheel(
      element: HTMLElement,
      deltaY: number,
      init: { deltaMode?: number; clientX?: number; clientY?: number } = {},
    ): WheelEvent {
      const event = new WheelEvent("wheel", {
        deltaY,
        deltaMode: init.deltaMode,
        bubbles: true,
        cancelable: true,
      });
      Object.defineProperty(event, "ctrlKey", { value: true });
      Object.defineProperty(event, "deltaMode", { value: init.deltaMode ?? 0 });
      if (init.clientX !== undefined) {
        Object.defineProperty(event, "clientX", { value: init.clientX });
      }
      if (init.clientY !== undefined) {
        Object.defineProperty(event, "clientY", { value: init.clientY });
      }
      act(() => {
        element.dispatchEvent(event);
      });
      return event;
    }

    it("zooms in on ctrl/cmd+wheel up and out on ctrl/cmd+wheel down", async () => {
      const store = await loadedStore();
      withStore(store, <Presentation.Viewport scrollZoom data-testid="viewport" />);
      const viewport = screen.getByTestId("viewport");

      // One mouse notch, applied multiplicatively.
      zoomWheel(viewport, -100);
      expect(store.getState().zoom).toBeCloseTo(Math.exp(0.15), 5);

      zoomWheel(viewport, 100);
      expect(store.getState().zoom).toBeCloseTo(1, 5);
    });

    it("claims the gesture so the page does not zoom", async () => {
      const store = await loadedStore();
      withStore(store, <Presentation.Viewport scrollZoom data-testid="viewport" />);

      const event = zoomWheel(screen.getByTestId("viewport"), -100);
      expect(event.defaultPrevented).toBe(true);
    });

    it("keeps the gesture at the ends of the zoom range", async () => {
      const store = await loadedStore();
      withStore(store, <Presentation.Viewport scrollZoom data-testid="viewport" />);
      act(() => store.setZoom(MAX_ZOOM));

      const event = zoomWheel(screen.getByTestId("viewport"), -100);
      expect(store.getState().zoom).toBe(MAX_ZOOM);
      // Releasing it here would hand a page zoom to someone still pinching in.
      expect(event.defaultPrevented).toBe(true);
    });

    it("normalizes deltas reported in lines", async () => {
      const store = await loadedStore();
      withStore(store, <Presentation.Viewport scrollZoom data-testid="viewport" />);

      zoomWheel(screen.getByTestId("viewport"), -3, { deltaMode: 1 });
      expect(store.getState().zoom).toBeCloseTo(Math.exp(48 * 0.0015), 5);
    });

    it("releases auto-fit, like any other explicit level", async () => {
      const store = await loadedStore();
      withStore(store, <Presentation.Viewport autoFit scrollZoom data-testid="viewport" />);
      expect(store.getState().isAutoFit).toBe(true);

      zoomWheel(screen.getByTestId("viewport"), -100);
      expect(store.getState().isAutoFit).toBe(false);
    });

    it("scrolls to keep the point under the pointer in place", async () => {
      const store = await loadedStore();
      withStore(
        store,
        <Presentation.Viewport scrollZoom data-testid="viewport">
          {/* Longhands: the test DOM does not expand the `overflow` shorthand. */}
          <div data-testid="scroller" style={{ overflowX: "auto", overflowY: "auto" }}>
            <div data-testid="content" />
          </div>
        </Presentation.Viewport>,
      );

      // The test DOM lays nothing out, so the slide's box has to grow with the
      // zoom by hand for there to be anything to correct for.
      // Scroll offsets clamp to a scrollable range the test DOM never has, so
      // they stand in as plain properties to record what the correction wrote.
      const scroller = screen.getByTestId("scroller");
      Object.defineProperty(scroller, "scrollLeft", { value: 0, writable: true });
      Object.defineProperty(scroller, "scrollTop", { value: 0, writable: true });

      const content = screen.getByTestId("content");
      content.getBoundingClientRect = () => {
        const size = 100 * store.getState().zoom;
        return { left: 0, top: 0, width: size, height: size } as DOMRect;
      };
      vi.stubGlobal("requestAnimationFrame", (callback: FrameRequestCallback) => {
        callback(0);
        return 0;
      });

      // Pointer at the center of the slide, which stays the center as it grows,
      // so the correction is half the growth on both axes.
      zoomWheel(content, -100, { clientX: 50, clientY: 50 });

      const growth = 100 * (store.getState().zoom - 1);
      expect(scroller.scrollLeft).toBeCloseTo(growth / 2, 5);
      expect(scroller.scrollTop).toBeCloseTo(growth / 2, 5);
    });

    it("zooms without navigating when scrollNavigation is on too", async () => {
      const store = await loadedStore();
      withStore(
        store,
        <Presentation.Viewport scrollNavigation scrollZoom data-testid="viewport" />,
      );

      // Cmd+wheel: the navigation handler runs first, so it has to pass on the
      // modifiers the zoom handler claims or one gesture does both. Downwards,
      // because navigating back off the first slide is a no-op either way.
      const event = new WheelEvent("wheel", { deltaY: 100, bubbles: true, cancelable: true });
      Object.defineProperty(event, "metaKey", { value: true });
      act(() => {
        screen.getByTestId("viewport").dispatchEvent(event);
      });

      expect(store.getState().zoom).toBeCloseTo(Math.exp(-0.15), 5);
      expect(store.getActiveSlideIndex()).toBe(0);
    });

    it("leaves the gesture to the browser when the prop is off", async () => {
      const store = await loadedStore();
      withStore(store, <Presentation.Viewport data-testid="viewport" />);

      const event = zoomWheel(screen.getByTestId("viewport"), -100);
      expect(store.getState().zoom).toBe(1);
      expect(event.defaultPrevented).toBe(false);
    });
  });
});
