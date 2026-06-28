import React from "react";
import { usePresentation, useSlide, useZoom } from "../context";
import { renderSlide } from "@aiden0z/pptx-renderer";
import type { PresentationData, SlideData, SlideHandle } from "@aiden0z/pptx-renderer";

export interface SlideProps {
  children?: React.ReactNode;
  className?: string;
  style?: React.CSSProperties;
}

export function Slide({ children, className, style }: SlideProps) {
  const { presentation, status } = usePresentation();
  const { slide } = useSlide();
  const { zoom } = useZoom();

  if (status === "loading") {
    return (
      <div
        className={className}
        style={{ display: "flex", alignItems: "center", justifyContent: "center", ...style }}
      >
        <span style={{ color: "#666" }}>Loading...</span>
      </div>
    );
  }
  if (status === "error")
    return (
      <div className={className} style={style}>
        Failed to parse
      </div>
    );
  if (!presentation || !slide) return null;

  return (
    <div
      className={className}
      style={{
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        overflow: "auto",
        ...style,
      }}
    >
      <SlideRenderer presentation={presentation} slide={slide} zoom={zoom}>
        {children}
      </SlideRenderer>
    </div>
  );
}

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

    // Dispose previous render
    if (handleRef.current) {
      handleRef.current.dispose();
      handleRef.current = null;
    }
    container.innerHTML = "";

    // Render the slide using the reference library's full renderer
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
