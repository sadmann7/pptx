import * as React from "react";

import { usePresentationStore, useZoom } from "./context";
import type { RenderProp } from "./render";
import { renderElement } from "./render";

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
   * Padding (in pixels) applied on all sides when fitting the slide.
   * Only used when `autoFit` is `true`.
   *
   * @default 24
   */
  autoFitPadding?: number;
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
  { children, className, style, autoFit = false, autoFitPadding = 24, render, ...viewportProps },
  forwardedRef,
) {
  const viewportRef = React.useRef<HTMLDivElement>(null);

  const store = usePresentationStore("PresentationViewport");
  const { zoom } = useZoom();

  React.useEffect(() => {
    if (!autoFit || !viewportRef.current) return;

    const viewportElement = viewportRef.current;
    const fit = () => {
      if (viewportElement.clientWidth > 0 && viewportElement.clientHeight > 0)
        store.fitTo(viewportElement.clientWidth, viewportElement.clientHeight, autoFitPadding);
    };
    const unsubscribe = store.subscribe(fit);
    fit();
    const resizeObserver = new ResizeObserver(fit);
    resizeObserver.observe(viewportElement);

    return () => {
      unsubscribe();
      resizeObserver.disconnect();
    };
  }, [autoFit, autoFitPadding, store]);

  return renderElement(
    "div",
    { render, className, style },
    {
      state: { zoom },
      ref: [viewportRef, forwardedRef],
      props: [
        {
          style: {
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            overflow: "auto",
          },
          children,
        },
        viewportProps,
      ],
    },
  );
});

export namespace Viewport {
  export type State = ViewportState;
  export type Props = ViewportProps;
}
