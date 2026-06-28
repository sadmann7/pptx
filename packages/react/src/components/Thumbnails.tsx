import * as React from "react";
import { usePresentation, useSlide } from "../context";
import { renderSlide } from "@pptx/parser";
import type { PresentationData, SlideData, SlideHandle } from "@pptx/parser";
import { renderElement } from "../utils/render";
import type { RenderProp } from "../utils/render";

export interface ThumbnailsState {
  total: number;
  currentIndex: number;
}

export interface ThumbnailsProps extends React.HTMLAttributes<HTMLDivElement> {
  /**
   * Width of each thumbnail in pixels.
   * Falls back to `style.width` for backward compatibility, then `140`.
   */
  thumbWidth?: number;
  /**
   * Replace the thumbnails container element.
   * - ReactElement: cloned with composed props
   * - Function: `(props, state) => ReactElement`
   */
  render?: RenderProp<React.HTMLAttributes<HTMLDivElement>, ThumbnailsState>;
}

/**
 * Scrollable sidebar that lists slide thumbnails.
 * Renders nothing until the presentation is `"ready"`.
 */
export const Thumbnails = React.forwardRef<HTMLDivElement, ThumbnailsProps>(function Thumbnails(
  { className, style, thumbWidth, render, ...elementProps },
  forwardedRef,
) {
  const { presentation, status } = usePresentation();
  const { index: currentIndex, goTo } = useSlide();

  if (status !== "ready" || !presentation) return null;

  const resolvedThumbWidth = thumbWidth ?? (style?.width as number | undefined) ?? 140;

  const state: ThumbnailsState = {
    total: presentation.slides.length,
    currentIndex,
  };

  const children = presentation.slides.map((slide, i) => (
    <ThumbnailItem
      key={i}
      index={i}
      slide={slide}
      presentation={presentation}
      isActive={i === currentIndex}
      thumbWidth={resolvedThumbWidth}
      onClick={() => goTo(i)}
    />
  ));

  return renderElement(
    "div",
    render as RenderProp<Record<string, unknown>, ThumbnailsState> | undefined,
    {
      ...elementProps,
      ref: forwardedRef,
      className,
      style: { overflowY: "auto", padding: 8, ...style },
      children,
    },
    state,
  );
});

// ---------------------------------------------------------------------------
// Internal: individual thumbnail button
// ---------------------------------------------------------------------------

interface ThumbnailItemProps {
  index: number;
  slide: SlideData;
  presentation: PresentationData;
  isActive: boolean;
  thumbWidth: number;
  onClick: () => void;
}

function ThumbnailItem({
  index,
  slide,
  presentation,
  isActive,
  thumbWidth,
  onClick,
}: ThumbnailItemProps) {
  const containerRef = React.useRef<HTMLDivElement>(null);
  const handleRef = React.useRef<SlideHandle | null>(null);
  const mediaUrlCache = React.useRef(new Map<string, string>()).current;

  const { width, height } = presentation;
  const scale = (thumbWidth - 16) / width;
  const thumbHeight = height * scale;

  React.useEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    if (handleRef.current) {
      handleRef.current.dispose();
      handleRef.current = null;
    }
    container.innerHTML = "";

    const handle = renderSlide(presentation, slide, { mediaUrlCache });
    handle.element.style.transform = `scale(${scale})`;
    handle.element.style.transformOrigin = "top left";
    container.appendChild(handle.element);
    handleRef.current = handle;

    return () => {
      if (handleRef.current) {
        handleRef.current.dispose();
        handleRef.current = null;
      }
    };
  }, [presentation, slide, scale, mediaUrlCache]);

  return (
    <button
      onClick={onClick}
      data-active={isActive || undefined}
      style={{
        display: "block",
        width: "100%",
        marginBottom: 8,
        padding: 4,
        border: isActive ? "2px solid #3b82f6" : "1px solid #e5e7eb",
        borderRadius: 4,
        background: "#fff",
        cursor: "pointer",
        overflow: "hidden",
        position: "relative",
      }}
    >
      <div
        ref={containerRef}
        style={{
          width: thumbWidth - 16,
          height: thumbHeight,
          overflow: "hidden",
          pointerEvents: "none",
        }}
      />
      <div style={{ position: "absolute", bottom: 2, right: 4, fontSize: 9, color: "#888" }}>
        {index + 1}
      </div>
    </button>
  );
}
