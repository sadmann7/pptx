import * as React from "react";

import { useStoreContext, useStoreEvent, useZoom } from "./context";
import type { RenderProp } from "./render";
import { renderElement } from "./render";
import type { AutoFitPadding, ZoomChangeEvent } from "./store";
import { MAX_ZOOM, MIN_ZOOM } from "./store";

/**
 * Ignore wheel events for this long after a wheel-triggered navigation.
 * Trackpads emit a burst of momentum events for one physical gesture;
 * without a cooldown a single flick would skip through several slides.
 */
const WHEEL_NAVIGATION_COOLDOWN_MS = 300;

/**
 * Zoom change per pixel of wheel delta, applied multiplicatively so a step is
 * the same proportion at every level. Tuned so one mouse notch (~100px) is
 * ~14%, which leaves a trackpad pinch (a few px per event) smooth.
 */
const ZOOM_WHEEL_SENSITIVITY = 0.0015;

export interface ViewportState {
  /** Current zoom level (1 = 100%, 0.5 = 50%). */
  zoom: number;
}

/** The point a zoom gesture has to keep still, and what to scroll to keep it. */
interface ZoomAnchor {
  scroller: HTMLElement;
  content: HTMLElement;
  ratioX: number;
  ratioY: number;
  pointerX: number;
  pointerY: number;
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

function getIsScrollable(overflow: string): boolean {
  return overflow === "auto" || overflow === "scroll";
}

/**
 * Wheel delta in pixels. Firefox reports lines and, when scrolling by page,
 * pages, which without normalizing would make a notch there barely register.
 */
function getWheelDeltaPx(event: WheelEvent): number {
  if (event.deltaMode === 1) return event.deltaY * 16;
  if (event.deltaMode === 2) return event.deltaY * 100;
  return event.deltaY;
}

export interface ViewportProps extends React.ComponentProps<"div"> {
  /**
   * When `true`, automatically scales the slide to fill the viewport's
   * dimensions whenever the container resizes.
   *
   * This is the starting mode, not a latch: `setZoom` (and `zoomIn`/
   * `zoomOut`) turn fitting off so an explicit level survives the next
   * resize, and `setAutoFit(true)` turns it back on. A zoom control offering
   * both therefore needs no state of its own.
   *
   * @default false
   */
  autoFit?: boolean;

  /**
   * When `true`, the mouse wheel navigates between slides like PowerPoint:
   * scrolling past the end of the current slide's scroll range advances to
   * the next slide, scrolling past the start goes back to the previous one.
   * When the slide fits the viewport (nothing to scroll), any wheel tick
   * navigates.
   *
   * Off by default: it captures wheel events, which is undesirable when the
   * viewer is embedded in a scrollable page.
   *
   * @default false
   */
  scrollNavigation?: boolean;

  /**
   * When `true`, the browser's zoom gesture zooms the deck instead of the page:
   * Ctrl/Cmd+wheel and trackpad pinch change the zoom level, keeping the point
   * under the pointer in place. Like PowerPoint, and like every canvas editor.
   *
   * Off by default: taking that gesture also takes it away from readers who
   * zoom the page to read, which is the wrong trade for a viewer embedded in a
   * page. Turn it on when the presentation owns the window. Ctrl/Cmd and the
   * plus/minus keys still zoom the page either way.
   *
   * @default false
   */
  scrollZoom?: boolean;

  /**
   * Padding (in pixels) reserved around the slide when fitting.
   * Only used when `autoFit` is `true`.
   *
   * A number applies the same padding on all sides. An object sets
   * per-side values (missing sides default to `0`), like Radix/Base UI
   * `collisionPadding`.
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
   * Replace the viewport container element.
   * - ReactElement: cloned with composed props
   * - Function: `(props, state) => ReactElement`
   */
  render?: RenderProp<ViewportState>;

  /**
   * Event handler called whenever the zoom level changes, including the
   * automatic fits performed while `autoFit` is on.
   *
   * ```tsx
   * onZoomChange={({ zoom }) => setZoom(`${Math.round(zoom * 100)}%`)}
   * ```
   */
  onZoomChange?: (event: ZoomChangeEvent) => void;
}

/**
 * Scrollable container that centers the slide and optionally auto-fits it
 * to the available space.
 *
 * Native `<div>` props are composed (not overwritten) with internals.
 * Place `<Presentation.Slide>` inside to render the current slide.
 */
export const Viewport = React.forwardRef<HTMLDivElement, ViewportProps>(function Viewport(
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
  const { zoom } = useZoom();

  useStoreEvent(store, "zoomChange", onZoomChange);

  // The prop seeds the store's fit mode rather than owning it, so a zoom
  // control can release and re-arm fitting without the consumer mirroring it
  // in state. Runs only when the prop itself changes, so it never fights a
  // `setAutoFit` call made from the UI.
  React.useEffect(() => {
    store.setAutoFit(autoFit);
  }, [autoFit, store]);

  React.useEffect(() => {
    if (!viewportRef.current) return;

    const viewportElement = viewportRef.current;
    // Measurement lives here because only the viewport knows its own box; the
    // store decides whether a measurement should be applied.
    const fit = () => {
      if (!store.getState().isAutoFit) return;
      if (viewportElement.clientWidth > 0 && viewportElement.clientHeight > 0)
        store.fitTo(viewportElement.clientWidth, viewportElement.clientHeight, autoFitPadding);
    };

    fit();

    // Observed even while fitting is off, so re-arming it picks up a resize
    // that happened in the meantime.
    const resizeObserver = new ResizeObserver(fit);
    resizeObserver.observe(viewportElement);

    // Re-fit when fitting is re-armed, and when a new presentation loads: the
    // new deck's aspect ratio may differ so the zoom needs to be recalculated.
    // Both are identity/flag flips, not zoom or navigation updates.
    let { presentation: lastPresentation, isAutoFit: wasAutoFit } = store.getState();
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

  // PowerPoint-style wheel navigation: scroll within the slide first; once
  // the scroll container hits its boundary in the wheel direction (or there
  // is nothing to scroll), a wheel tick changes slides. External system
  // (native wheel events, non-passive for preventDefault), so an effect is
  // the right tool.
  React.useEffect(() => {
    if (!scrollNavigation || !viewportRef.current) return;

    const viewportElement = viewportRef.current;
    let cooldownUntil = 0;

    /**
     * Nearest vertically scrollable element between the event target and the
     * viewport (inclusive). This is normally the `Presentation.Slide`
     * wrapper (`overflow: auto`), but resolving it dynamically keeps the
     * behavior correct with custom `render` layouts.
     */
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
      // Ctrl+wheel = pinch-zoom, shift+wheel = horizontal scroll; leave both alone.
      if (event.deltaY === 0 || event.ctrlKey || event.shiftKey) return;

      const goingDown = event.deltaY > 0;
      const scroller = findScroller(event.target);
      if (scroller) {
        const atBottom = scroller.scrollTop + scroller.clientHeight >= scroller.scrollHeight - 1;
        const atTop = scroller.scrollTop <= 0;
        // Not at the boundary yet: let the browser scroll the slide.
        if (goingDown ? !atBottom : !atTop) return;
      }

      if (event.timeStamp < cooldownUntil) {
        // Momentum tail of the gesture that just navigated: swallow it so it
        // neither chains to the page nor immediately scrolls the new slide.
        event.preventDefault();
        return;
      }

      if (goingDown ? !store.canGoNext() : !store.canGoPrev()) return;

      event.preventDefault();
      cooldownUntil = event.timeStamp + WHEEL_NAVIGATION_COOLDOWN_MS;
      if (goingDown) store.next();
      else store.prev();

      if (scroller) {
        // Start the new slide from the edge you entered it from (top when
        // going forward, bottom when going back), like PowerPoint. Deferred
        // one frame so the new slide's content height is in place.
        requestAnimationFrame(() => {
          scroller.scrollTop = goingDown ? 0 : scroller.scrollHeight;
        });
      }
    }

    // Non-passive: we call preventDefault to stop scroll chaining to the page.
    viewportElement.addEventListener("wheel", onWheel, { passive: false });
    return () => viewportElement.removeEventListener("wheel", onWheel);
  }, [scrollNavigation, store]);

  // Ctrl/Cmd+wheel zoom. External system again (the browser owns this gesture
  // until preventDefault takes it, which a passive listener cannot do).
  React.useEffect(() => {
    if (!scrollZoom || !viewportRef.current) return;

    const viewportElement = viewportRef.current;

    /**
     * Nearest scroll container between the event target and the viewport
     * (inclusive), normally the `Presentation.Slide` wrapper. Matched on
     * `overflow` rather than on current scrollability, because at a fitted
     * zoom there is nothing to scroll yet and zooming in is what creates it.
     */
    function findScroller(from: EventTarget | null): HTMLElement | null {
      let element = from instanceof HTMLElement ? from : null;
      while (element) {
        const { overflowX, overflowY } = getComputedStyle(element);
        if (getIsScrollable(overflowX) || getIsScrollable(overflowY)) return element;
        if (element === viewportElement) break;
        element = element.parentElement;
      }
      return null;
    }

    /**
     * Where the pointer sits on the slide, as a fraction of the slide's box.
     * Fractions rather than scroll offsets because the slide is centered by
     * `margin: auto` until it outgrows the viewport, so its origin moves as
     * the zoom crosses that threshold.
     */
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

      // Claimed before the level is checked: at either end of the range the
      // gesture is still ours, and releasing it there would zoom the page.
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

      // The correction has to wait for the slide to be laid out at the new
      // level; measuring it beats predicting it, since `margin: auto` and the
      // scrollbars themselves both move the box.
      requestAnimationFrame(() => {
        const rect = anchor.content.getBoundingClientRect();
        anchor.scroller.scrollLeft += rect.left + anchor.ratioX * rect.width - anchor.pointerX;
        anchor.scroller.scrollTop += rect.top + anchor.ratioY * rect.height - anchor.pointerY;
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
});

export namespace Viewport {
  export type State = ViewportState;
  export type Props = ViewportProps;
}
