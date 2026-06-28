import React from "react";
import { usePresentationStoreRef } from "../context";

export interface ViewportProps {
  children: React.ReactNode;
  className?: string;
  style?: React.CSSProperties;
  /**
   * If true, the viewport computes and applies `fitTo` whenever the container
   * is resized OR when the presentation finishes loading.
   * Requires the ResizeObserver API (available in all modern browsers).
   * @default false
   */
  autoFit?: boolean;
  /** Padding (in px) subtracted from each edge when auto-fitting. @default 24 */
  autoFitPadding?: number;
}

export function Viewport({
  children,
  className,
  style,
  autoFit = false,
  autoFitPadding = 24,
}: ViewportProps) {
  const ref = React.useRef<HTMLDivElement>(null);
  const store = usePresentationStoreRef();

  React.useEffect(() => {
    if (!autoFit || !ref.current) return;

    const el = ref.current;

    const fit = () => {
      if (el.clientWidth > 0 && el.clientHeight > 0) {
        store.fitTo(el.clientWidth, el.clientHeight, autoFitPadding);
      }
    };

    // Re-fit whenever the store state changes (e.g. presentation finishes loading)
    // and whenever the container is resized.
    const unsubscribeStore = store.subscribe(fit);

    fit(); // attempt immediately (no-op if presentation not loaded yet)

    let observer: ResizeObserver | undefined;
    if (typeof ResizeObserver !== "undefined") {
      observer = new ResizeObserver(fit);
      observer.observe(el);
    }

    return () => {
      unsubscribeStore();
      observer?.disconnect();
    };
  }, [autoFit, autoFitPadding, store]);

  const containerStyle: React.CSSProperties = {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    overflow: "auto",
    position: "relative",
    ...style,
  };

  return (
    <div ref={ref} className={className} style={containerStyle}>
      {children}
    </div>
  );
}
