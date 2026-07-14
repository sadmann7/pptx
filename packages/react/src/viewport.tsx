import * as React from "react";

import { useStoreContext, useZoom } from "./context";
import type { RenderProp } from "./render";
import { renderElement } from "./render";
import type { AutoFitPadding } from "./store";

export interface ViewportState {
  /** Current zoom level (1 = 100%, 0.5 = 50%). */
  zoom: number;
}

export interface ViewportProps extends React.ComponentProps<"div"> {
  /**
   * When `true`, automatically scales the slide to fill the viewport's
   * dimensions whenever the container resizes.
   *
   * @default false
   */
  autoFit?: boolean;
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
}

/**
 * Scrollable container that centers the slide and optionally auto-fits it
 * to the available space.
 *
 * Native `<div>` props are composed (not overwritten) with internals.
 * Place `<Presentation.Slide>` inside to render the current slide.
 */
export const Viewport = React.forwardRef<HTMLDivElement, ViewportProps>(function Viewport(
  { autoFit = false, autoFitPadding = 0, render, ...viewportProps },
  forwardedRef,
) {
  const viewportRef = React.useRef<HTMLDivElement>(null);

  const store = useStoreContext("PresentationViewport");
  const { zoom } = useZoom();

  React.useEffect(() => {
    if (!autoFit || !viewportRef.current) return;

    const viewportElement = viewportRef.current;
    const fit = () => {
      if (viewportElement.clientWidth > 0 && viewportElement.clientHeight > 0)
        store.fitTo(viewportElement.clientWidth, viewportElement.clientHeight, autoFitPadding);
    };

    fit();

    // Re-fit when the container is resized.
    const resizeObserver = new ResizeObserver(fit);
    resizeObserver.observe(viewportElement);

    // Re-fit when a new presentation loads: the new slide's aspect ratio may
    // differ so the zoom needs to be recalculated. Only fires on presentation
    // identity changes, not on zoom/navigation/progress updates.
    let lastPresentation = store.getState().presentation;
    const unsubscribe = store.subscribe(() => {
      const presentation = store.getState().presentation;
      if (presentation !== lastPresentation) {
        lastPresentation = presentation;
        fit();
      }
    });

    return () => {
      unsubscribe();
      resizeObserver.disconnect();
    };
  }, [autoFit, autoFitPadding, store]);

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
