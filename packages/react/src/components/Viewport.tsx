import React from "react";
import { usePresentationStoreRef } from "../context";

export interface ViewportProps {
  children: React.ReactNode;
  className?: string;
  style?: React.CSSProperties;
  autoFit?: boolean;
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
      if (el.clientWidth > 0 && el.clientHeight > 0)
        store.fitTo(el.clientWidth, el.clientHeight, autoFitPadding);
    };
    const unsubscribe = store.subscribe(fit);
    fit();
    const observer = new ResizeObserver(fit);
    observer.observe(el);
    return () => {
      unsubscribe();
      observer.disconnect();
    };
  }, [autoFit, autoFitPadding, store]);

  return (
    <div
      ref={ref}
      className={className}
      style={{
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        overflow: "auto",
        ...style,
      }}
    >
      {children}
    </div>
  );
}
