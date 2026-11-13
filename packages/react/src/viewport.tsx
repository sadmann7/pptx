import * as React from "react";

import { useStoreContext, useStoreEvent, useStoreSelector } from "./context";
import type { PrimitiveProps } from "./render";
import { renderElement } from "./render";
import type { AutoFitPadding, ZoomChangeEvent } from "./store";
import { MAX_ZOOM, MIN_ZOOM } from "./store";
import { clamp } from "./utils";

// Trackpad momentum sends a burst of wheel events per flick, which would otherwise skip several slides.
const WHEEL_NAVIGATION_COOLDOWN_MS = 300;

// One mouse notch (~100px) zooms ~14%, which keeps a trackpad pinch (a few px per event) smooth.
const ZOOM_WHEEL_SENSITIVITY = 0.0015;

export interface ViewportState {
  /** Current zoom level, where `1` is 100%. */
  zoom: number;
}

interface ZoomAnchor {
  scroller: HTMLElement;
  content: HTMLElement;
  ratioX: number;
  ratioY: number;
  pointerX: number;
  pointerY: number;
}

function getIsScrollable(overflow: string): boolean {
  return overflow === "auto" || overflow === "scroll";
}

// Firefox reports lines, or pages when scrolling by page.
function getWheelDeltaPx(event: WheelEvent): number {
  if (event.deltaMode === 1) return event.deltaY * 16;
  if (event.deltaMode === 2) return event.deltaY * 100;
  return event.deltaY;
}

export interface ViewportProps extends PrimitiveProps<"div", ViewportState> {
  /**
   * When `true`, scales the slide to fill the viewport whenever it resizes. `setZoom`, `zoomIn`
   * and `zoomOut` turn fitting off, and `setAutoFit(true)` turns it back on.
   *
   * @default false
   */
  autoFit?: boolean;

  /**
   * When `true`, scrolling past the current slide's scroll range moves to the next or previous
   * slide, like PowerPoint. It captures wheel events from an embedding page.
   *
   * @default false
   */
  scrollNavigation?: boolean;

  /**
   * When `true`, Ctrl/Cmd+wheel and trackpad pinch zoom the deck around the pointer instead of
   * zooming the page. Ctrl/Cmd with plus or minus still zooms the page.
   *
   * @default false
   */
  scrollZoom?: boolean;

  /**
   * Padding in pixels around the slide when `autoFit` is `true`, for every side or per side.
   *
   * @example
   * autoFitPadding={32}
   * autoFitPadding={{ top: 48, bottom: 16 }}
   * autoFitPadding={{ top: 8, right: 16, bottom: 8, left: 16 }}
   *
   * @default 0
   */
  autoFitPadding?: AutoFitPadding;

  /**
   * Event handler called whenever the zoom level changes, including fits while `autoFit` is on.
   *
   * ```tsx
   * onZoomChange={({ zoom }) => setZoom(`${Math.round(zoom * 100)}%`)}
   * ```
   */
  onZoomChange?: (event: ZoomChangeEvent) => void;
}

export const Viewport = React.forwardRef<HTMLDivElement, ViewportProps>(
  function Viewport(
    {
      autoFit = false,
      autoFitPadding = 0,
      scrollNavigation = false,
      scrollZoom = false,
      render,
      onZoomChange,
      ...viewportProps
    },
    forwardedRef,
  ) {
    const viewportRef = React.useRef<HTMLDivElement>(null);

    const store = useStoreContext("PresentationViewport");
    const zoom = useStoreSelector(store, (s) => s.zoom, 1);

    useStoreEvent(store, "zoomChange", onZoomChange);

    // The prop seeds the store's fit mode rather than owning it, so it never fights a `setAutoFit` call.
    React.useEffect(() => {
      store.setAutoFit(autoFit);
    }, [autoFit, store]);

    React.useEffect(() => {
      if (!viewportRef.current) return;

      const viewportElement = viewportRef.current;
      const fit = () => {
        if (!store.getState().isAutoFit) return;
        if (viewportElement.clientWidth > 0 && viewportElement.clientHeight > 0)
          store.fitTo(
            viewportElement.clientWidth,
            viewportElement.clientHeight,
            autoFitPadding,
          );
      };

      fit();

      // Observed even while fitting is off, so re-arming it picks up a resize in the meantime.
      const resizeObserver = new ResizeObserver(fit);
      resizeObserver.observe(viewportElement);

      // A new deck's aspect ratio may differ, so a reload re-fits as well.
      let { presentation: lastPresentation, isAutoFit: wasAutoFit } =
        store.getState();
      const unsubscribe = store.subscribe(() => {
        const { presentation, isAutoFit } = store.getState();
        const isArmed = isAutoFit && !wasAutoFit;
        const isReloaded = presentation !== lastPresentation;
        lastPresentation = presentation;
        wasAutoFit = isAutoFit;
        if (isArmed || isReloaded) fit();
      });

      return () => {
        unsubscribe();
        resizeObserver.disconnect();
      };
    }, [autoFitPadding, store]);

    // As in PowerPoint, the slide scrolls first and a tick past its boundary changes slides.
    React.useEffect(() => {
      if (!scrollNavigation || !viewportRef.current) return;

      const viewportElement = viewportRef.current;
      let cooldownUntil = 0;

      // Resolved per event rather than assumed to be the slide wrapper, for custom `render` layouts.
      function findScroller(from: EventTarget | null): HTMLElement | null {
        let element = from instanceof HTMLElement ? from : null;
        while (element) {
          if (element.scrollHeight > element.clientHeight + 1) {
            const { overflowY } = getComputedStyle(element);
            if (overflowY === "auto" || overflowY === "scroll") return element;
          }
          if (element === viewportElement) break;
          element = element.parentElement;
        }
        return null;
      }

      function onWheel(event: WheelEvent): void {
        if (
          event.deltaY === 0 ||
          event.ctrlKey ||
          event.metaKey ||
          event.shiftKey
        )
          return;

        const goingDown = event.deltaY > 0;
        const scroller = findScroller(event.target);
        if (scroller) {
          const atBottom =
            scroller.scrollTop + scroller.clientHeight >=
            scroller.scrollHeight - 1;
          const atTop = scroller.scrollTop <= 0;
          if (goingDown ? !atBottom : !atTop) return;
        }

        if (event.timeStamp < cooldownUntil) {
          event.preventDefault();
          return;
        }

        if (goingDown ? !store.canGoNext() : !store.canGoPrev()) return;

        event.preventDefault();
        cooldownUntil = event.timeStamp + WHEEL_NAVIGATION_COOLDOWN_MS;
        if (goingDown) store.next();
        else store.prev();

        if (scroller) {
          // As in PowerPoint, the new slide starts from the edge it was entered from, a frame later
          // once its content height is in place.
          requestAnimationFrame(() => {
            scroller.scrollTop = goingDown ? 0 : scroller.scrollHeight;
          });
        }
      }

      // Non-passive so preventDefault can stop scroll chaining to the page.
      viewportElement.addEventListener("wheel", onWheel, { passive: false });
      return () => viewportElement.removeEventListener("wheel", onWheel);
    }, [scrollNavigation, store]);

    React.useEffect(() => {
      if (!scrollZoom || !viewportRef.current) return;

      const viewportElement = viewportRef.current;

      // Matched on `overflow`, not current scrollability: at a fitted zoom there is nothing to scroll yet.
      function findScroller(from: EventTarget | null): HTMLElement | null {
        let element = from instanceof HTMLElement ? from : null;
        while (element) {
          const { overflowX, overflowY } = getComputedStyle(element);
          if (getIsScrollable(overflowX) || getIsScrollable(overflowY))
            return element;
          if (element === viewportElement) break;
          element = element.parentElement;
        }
        return null;
      }

      // Fractions, not scroll offsets: `margin: auto` centers the slide until it outgrows the viewport.
      function getAnchor(event: WheelEvent): ZoomAnchor | null {
        const scroller = findScroller(event.target);
        const content = scroller?.firstElementChild;
        if (!scroller || !(content instanceof HTMLElement)) return null;

        const rect = content.getBoundingClientRect();
        if (rect.width === 0 || rect.height === 0) return null;

        return {
          scroller,
          content,
          ratioX: clamp((event.clientX - rect.left) / rect.width, 0, 1),
          ratioY: clamp((event.clientY - rect.top) / rect.height, 0, 1),
          pointerX: event.clientX,
          pointerY: event.clientY,
        };
      }

      function onWheel(event: WheelEvent): void {
        if (!(event.ctrlKey || event.metaKey) || event.deltaY === 0) return;

        // Claimed before the level is checked: releasing the gesture at either end
        // of the range would zoom the page.
        event.preventDefault();

        const { zoom } = store.getState();
        const next = clamp(
          zoom * Math.exp(-getWheelDeltaPx(event) * ZOOM_WHEEL_SENSITIVITY),
          MIN_ZOOM,
          MAX_ZOOM,
        );
        if (next === zoom) return;

        const anchor = getAnchor(event);
        store.setZoom(next);
        if (!anchor) return;

        // Measured after layout, since `margin: auto` and the scrollbars both move the box.
        requestAnimationFrame(() => {
          const rect = anchor.content.getBoundingClientRect();
          anchor.scroller.scrollLeft +=
            rect.left + anchor.ratioX * rect.width - anchor.pointerX;
          anchor.scroller.scrollTop +=
            rect.top + anchor.ratioY * rect.height - anchor.pointerY;
        });
      }

      viewportElement.addEventListener("wheel", onWheel, { passive: false });
      return () => viewportElement.removeEventListener("wheel", onWheel);
    }, [scrollZoom, store]);

    return renderElement(
      "div",
      { render },
      {
        state: { zoom },
        ref: [viewportRef, forwardedRef],
        props: [viewportProps],
      },
    );
  },
);

export namespace Viewport {
  export type State = ViewportState;
  export type Props = ViewportProps;
}
