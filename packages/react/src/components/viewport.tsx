import * as React from "react";
import { useZoom, usePresentationStore } from "../context";
import { renderElement } from "../utils/render";
import type { RenderProp } from "../utils/render";

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
 * Spread any native `<div>` props — they are composed (not overwritten) with internals.
 */
export const Viewport = React.forwardRef<HTMLDivElement, ViewportProps>(function Viewport(
  { children, className, style, autoFit = false, autoFitPadding = 24, render, ...elementProps },
  forwardedRef,
) {
  const internalRef = React.useRef<HTMLDivElement>(null);
  const store = usePresentationStore("Viewport");
  const { zoom } = useZoom();

  React.useEffect(() => {
    if (!autoFit || !internalRef.current) return;
    const el = internalRef.current;
    const fit = () => {
      if (el.clientWidth > 0 && el.clientHeight > 0)
        store.fitTo(el.clientWidth, el.clientHeight, autoFitPadding);
    };
    const unsubscribe = store.subscribe(fit);
    fit();
    const ro = new ResizeObserver(fit);
    ro.observe(el);
    return () => {
      unsubscribe();
      ro.disconnect();
    };
  }, [autoFit, autoFitPadding, store]);

  return renderElement(
    "div",
    { render, className, style },
    {
      state: { zoom },
      ref: [internalRef, forwardedRef],
      props: {
        ...elementProps,
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
