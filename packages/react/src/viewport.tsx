import * as React from "react";

import { usePresentationStore, useZoom } from "./context";
import type { RenderProp } from "./render";
import { renderElement } from "./render";

export interface ViewportState {
  zoom: number;
}

export interface ViewportProps extends React.ComponentProps<"div"> {
  autoFit?: boolean;
  autoFitPadding?: number;
  /**
   * Replace the viewport container element.
   * - ReactElement: cloned with composed props
   * - Function: `(props, state) => ReactElement`
   */
  render?: RenderProp<ViewportState>;
}

/**
 * Scrollable container that optionally auto-fits the slide to its size.
 * Spread any native `<div>` props: they are composed (not overwritten) with internals.
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
      props: {
        ...viewportProps,
        style: {
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          overflow: "auto",
        },
        children,
      },
    },
  );
});

export namespace Viewport {
  export type State = ViewportState;
  export type Props = ViewportProps;
}
