import React from "react";
import { usePresentation, useSlide } from "../context";
import { renderSlide } from "@aiden0z/pptx-renderer";
import type { PresentationData, SlideData, SlideHandle } from "@aiden0z/pptx-renderer";

export interface ThumbnailsProps {
  className?: string;
  style?: React.CSSProperties;
}

export function Thumbnails({ className, style }: ThumbnailsProps) {
  const { presentation, status } = usePresentation();
  const { index: currentIndex, goTo } = useSlide();

  if (status !== "ready" || !presentation) return null;

  const thumbWidth = (style?.width as number) ?? 140;

  return (
    <div className={className} style={{ overflowY: "auto", padding: 8, ...style }}>
      {presentation.slides.map((slide, i) => (
        <ThumbnailItem
          key={i}
          index={i}
          slide={slide}
          presentation={presentation}
          isActive={i === currentIndex}
          thumbWidth={thumbWidth}
          onClick={() => goTo(i)}
        />
      ))}
    </div>
  );
}

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
