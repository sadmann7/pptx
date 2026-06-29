import * as React from "react";
import { usePresentation, useSlide, useZoom } from "../context";
import { renderSlide } from "@diceui/pptx-parser";
import type { PresentationData, SlideData, SlideHandle } from "@diceui/pptx-parser";
import type { PresentationStatus } from "../store";
import { useRenderElement } from "../utils/render";
import type { RenderProp } from "../utils/render";

export interface SlideState {
  /** Current parse/load status. Reflected as `data-status` on the element. */
  status: PresentationStatus;
  /** Zero-based index of the active slide. */
  index: number;
}

export interface SlideProps extends React.ComponentProps<"div"> {
  /**
   * Replace the slide wrapper element.
   * - ReactElement: cloned with composed props
   * - Function: `(props, state) => ReactElement`
   */
  render?: RenderProp<SlideState>;
}

/**
 * Renders the current slide inside a centered wrapper.
 * The wrapper always mounts so that sibling layout is stable — slide content
 * is absent until the presentation is `"ready"`.
 *
 * Use `<Presentation.Loading>` / `<Presentation.Error>` to display status UI.
 *
 * The element carries a `data-status` attribute matching the store status,
 * enabling CSS-driven state styles.
 */
export const Slide = React.forwardRef<HTMLDivElement, SlideProps>(function Slide(
  { children, className, style, render, ...elementProps },
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

  return useRenderElement(
    "div",
    { render, className, style },
    {
      state: { status, index },
      ref: forwardedRef,
      props: {
        ...elementProps,
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
  const handleRef = React.useRef<SlideHandle | null>(null);
  const mediaUrlCache = React.useRef(new Map<string, string>()).current;

  React.useEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    // Prevent Tab from reaching focusable PPTX content (links, forms, etc.)
    container.setAttribute("inert", "");

    if (handleRef.current) {
      handleRef.current.dispose();
      handleRef.current = null;
    }
    container.innerHTML = "";

    const handle = renderSlide(presentation, slide, {
      mediaUrlCache,
      onNodeError: (nodeId, error) => {
        console.warn(`[pptx] Node render error: ${nodeId}`, error);
      },
    });

    container.appendChild(handle.element);
    handleRef.current = handle;

    return () => {
      if (handleRef.current) {
        handleRef.current.dispose();
        handleRef.current = null;
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
