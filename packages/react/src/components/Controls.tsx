import React from "react";
import { useSlide, useZoom } from "../context";

// ─── Compound control components ─────────────────────────────────────────────

export interface PreviousSlideProps extends React.ButtonHTMLAttributes<HTMLButtonElement> {
  children?: React.ReactNode;
}

export function PreviousSlide({ children, onClick, ...props }: PreviousSlideProps) {
  const { prev, isFirst } = useSlide();
  return (
    <button
      type="button"
      aria-label="Previous slide"
      disabled={isFirst}
      onClick={(e) => {
        prev();
        onClick?.(e);
      }}
      {...props}
    >
      {children ?? "‹"}
    </button>
  );
}

export interface NextSlideProps extends React.ButtonHTMLAttributes<HTMLButtonElement> {
  children?: React.ReactNode;
}

export function NextSlide({ children, onClick, ...props }: NextSlideProps) {
  const { next, isLast } = useSlide();
  return (
    <button
      type="button"
      aria-label="Next slide"
      disabled={isLast}
      onClick={(e) => {
        next();
        onClick?.(e);
      }}
      {...props}
    >
      {children ?? "›"}
    </button>
  );
}

export interface SlideCounterProps {
  className?: string;
  style?: React.CSSProperties;
  /** Format function. Receives 1-based index and total. */
  format?: (current: number, total: number) => string;
}

export function SlideCounter({ className, style, format }: SlideCounterProps) {
  const { index, total } = useSlide();
  const text = format ? format(index + 1, total) : `${index + 1} / ${total}`;
  return (
    <span
      className={className}
      style={{ fontVariantNumeric: "tabular-nums", ...style }}
      aria-live="polite"
      aria-atomic="true"
    >
      {text}
    </span>
  );
}

export interface ZoomInProps extends React.ButtonHTMLAttributes<HTMLButtonElement> {
  step?: number;
}

export function ZoomIn({ step, children, onClick, ...props }: ZoomInProps) {
  const { zoomIn } = useZoom();
  return (
    <button
      type="button"
      aria-label="Zoom in"
      onClick={(e) => {
        zoomIn(step);
        onClick?.(e);
      }}
      {...props}
    >
      {children ?? "+"}
    </button>
  );
}

export interface ZoomOutProps extends React.ButtonHTMLAttributes<HTMLButtonElement> {
  step?: number;
}

export function ZoomOut({ step, children, onClick, ...props }: ZoomOutProps) {
  const { zoomOut } = useZoom();
  return (
    <button
      type="button"
      aria-label="Zoom out"
      onClick={(e) => {
        zoomOut(step);
        onClick?.(e);
      }}
      {...props}
    >
      {children ?? "−"}
    </button>
  );
}

export interface ZoomResetProps extends React.ButtonHTMLAttributes<HTMLButtonElement> {
  targetZoom?: number;
}

export function ZoomReset({ targetZoom = 1, children, onClick, ...props }: ZoomResetProps) {
  const { setZoom } = useZoom();
  return (
    <button
      type="button"
      aria-label="Reset zoom"
      onClick={(e) => {
        setZoom(targetZoom);
        onClick?.(e);
      }}
      {...props}
    >
      {children ?? "100%"}
    </button>
  );
}
