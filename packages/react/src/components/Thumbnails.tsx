import React from "react";
import { usePresentation, useSlide } from "../context";
import { materializeSlideNodes } from "@pptx/parser";
import { SlideCanvas } from "./Slide";

export interface ThumbnailsProps {
  className?: string;
  style?: React.CSSProperties;
}

export function Thumbnails({ className, style }: ThumbnailsProps) {
  const { presentation, status } = usePresentation();
  const { index: currentIndex, goTo } = useSlide();
  const mediaUrlCache = React.useRef(new Map<string, string>()).current;

  if (status !== "ready" || !presentation) return null;

  const thumbWidth = (style?.width as number) ?? 140;
  const scale = (thumbWidth - 16) / presentation.width;

  return (
    <div className={className} style={{ overflowY: "auto", padding: 8, ...style }}>
      {presentation.slides.map((slide, i) => {
        materializeSlideNodes(presentation, slide);
        return (
          <button
            key={i}
            onClick={() => goTo(i)}
            style={{
              display: "block",
              width: "100%",
              marginBottom: 8,
              padding: 4,
              border: i === currentIndex ? "2px solid #3b82f6" : "1px solid #e5e7eb",
              borderRadius: 4,
              background: "#fff",
              cursor: "pointer",
              overflow: "hidden",
              position: "relative",
            }}
          >
            <div style={{ width: presentation.width * scale, height: presentation.height * scale, overflow: "hidden", pointerEvents: "none" }}>
              <SlideCanvas
                width={presentation.width}
                height={presentation.height}
                zoom={scale}
                nodes={slide.nodes}
                presentation={presentation}
                slideRels={slide.rels}
                mediaUrlCache={mediaUrlCache}
              />
            </div>
            <div style={{ position: "absolute", bottom: 2, right: 4, fontSize: 9, color: "#888" }}>{i + 1}</div>
          </button>
        );
      })}
    </div>
  );
}
