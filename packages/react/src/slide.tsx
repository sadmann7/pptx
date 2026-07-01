import * as React from "react";

import type { PresentationData, SlideData, SlideHandle } from "@diceui/pptx-parser";
import { materializeSlideNodes, renderSlide } from "@diceui/pptx-parser";

import { usePresentation, useSlide, useZoom } from "./context";
import type { RenderProp } from "./render";
import { renderElement } from "./render";
import type { PresentationStatus } from "./store";

export interface SlideState {
  /**
   * Current parse/load status. Reflected as `data-status` on the wrapper
   * element so CSS can target states directly (e.g. `[data-status="loading"]`).
   */
  status: PresentationStatus;
  /** 0-based index of the active slide in the loaded presentation. */
  index: number;
}

export interface SlideProps extends React.ComponentProps<"div"> {
  /**
   * Replace the slide wrapper element.
   * - ReactElement: cloned with composed props
   * - Function: `(props, state) => ReactElement`
   *
   * @example
   * render={(props, { status }) => (
   *   <article {...props} data-slide-status={status} />
   * )}
   */
  render?: RenderProp<SlideState>;
}

/**
 * Renders the active slide inside a centered wrapper `<div>`.
 *
 * The wrapper always mounts so that sibling layout is stable. Slide content
 * is absent until the presentation is `"ready"`. Use `<Presentation.Loading>`
 * and `<Presentation.Error>` alongside this component to cover other states.
 *
 * The wrapper carries a `data-status` attribute matching the store status,
 * enabling CSS-driven state styles without extra JavaScript.
 */
export const Slide = React.forwardRef<HTMLDivElement, SlideProps>(function Slide(
  { children, className, style, render, ...slideProps },
  forwardedRef,
) {
  const { presentation, status } = usePresentation();
  const { slide, index } = useSlide();
  const { zoom } = useZoom();

  const slideContent =
    presentation && slide ? (
      <SlideRenderer presentation={presentation} slide={slide} zoom={zoom}>
        {children}
      </SlideRenderer>
    ) : null;

  return renderElement(
    "div",
    { render, className, style },
    {
      state: { status, index },
      ref: forwardedRef,
      props: {
        ...slideProps,
        "data-status": status,
        style: {
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          overflow: "auto",
        },
        children: slideContent,
      },
    },
  );
});

// ---------------------------------------------------------------------------
// Internal: imperative slide renderer (untouched from original)
// ---------------------------------------------------------------------------

interface SlideRendererProps {
  presentation: PresentationData;
  slide: SlideData;
  zoom: number;
  children?: React.ReactNode;
}

function SlideRenderer({ presentation, slide, zoom, children }: SlideRendererProps) {
  const containerRef = React.useRef<HTMLDivElement>(null);
  const slideHandleRef = React.useRef<SlideHandle | null>(null);
  const mediaUrlCache = React.useRef(new Map<string, string>()).current;

  React.useEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    if (slideHandleRef.current) {
      slideHandleRef.current.dispose();
      slideHandleRef.current = null;
    }
    container.innerHTML = "";

    if (!slide.nodesMaterialized) materializeSlideNodes(presentation, slide);
    const handle = renderSlide(presentation, slide, {
      mediaUrlCache,
      onNodeError: (nodeId, error) => {
        console.warn(`[pptx] Node render error: ${nodeId}`, error);
      },
    });

    container.appendChild(handle.element);
    slideHandleRef.current = handle;

    return () => {
      if (slideHandleRef.current) {
        slideHandleRef.current.dispose();
        slideHandleRef.current = null;
      }
    };
  }, [presentation, slide, mediaUrlCache]);

  const { width, height } = presentation;

  return (
    <div
      style={{ width: width * zoom, height: height * zoom, flexShrink: 0, position: "relative" }}
    >
      <div
        ref={containerRef}
        style={{
          width,
          height,
          transformOrigin: "top left",
          transform: `scale(${zoom})`,
          position: "relative",
          overflow: "hidden",
        }}
      />
      {children && (
        <div style={{ position: "absolute", inset: 0, pointerEvents: "none" }}>{children}</div>
      )}
    </div>
  );
}

export namespace Slide {
  export type State = SlideState;
  export type Props = SlideProps;
}
