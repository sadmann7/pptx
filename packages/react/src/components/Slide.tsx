import React from "react";
import { usePresentation, useSlide, useZoom } from "../context";
import { fillToCSS, toCSS } from "../render/color";
import { type ElementRendererFn, SlideElementRenderer } from "../elements/index";

export interface SlideProps {
  /**
   * Override the renderer for one or more element types.
   * Return undefined to fall back to the default renderer.
   */
  renderElement?: ElementRendererFn;
  /**
   * Overlay children rendered on top of the slide content.
   * Receives the same coordinate space as the slide (pt units, absolute positioning).
   */
  children?: React.ReactNode;
  className?: string;
  style?: React.CSSProperties;
}

/**
 * <Presentation.Slide>
 *
 * Renders the current slide's elements inside a fixed-pt canvas that is
 * CSS-scaled to match the current zoom level.
 *
 * The slide canvas uses the presentation's native pt dimensions so that:
 *   - All absolute-positioned elements line up pixel-perfectly
 *   - Zoom is a single CSS `scale()` on the container — no re-layout needed
 *   - Children (overlays) share the same coordinate space
 */
export function Slide({ renderElement, children, className, style }: SlideProps) {
  const { presentation, status, progress } = usePresentation();
  const { slide } = useSlide();
  const { zoom } = useZoom();

  if (status === "loading") {
    return (
      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          padding: "2rem",
          fontFamily: "sans-serif",
          color: "#6b7280",
          fontSize: "14px",
        }}
      >
        Parsing… {progress}%
      </div>
    );
  }

  if (status === "error") {
    return (
      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          padding: "2rem",
          fontFamily: "sans-serif",
          color: "#ef4444",
          fontSize: "14px",
        }}
      >
        Failed to parse presentation
      </div>
    );
  }

  if (status !== "ready" || !presentation || !slide) return null;

  const { width, height } = presentation.slideSize;
  const themeColors = presentation.theme.colors;

  // Background: solid color or fill from parsed slide/layout/master
  const bgColor = slide.background ? fillToCSS(slide.background.fill, themeColors) : "#ffffff";

  // Default text color from theme's dk1 (primary dark color).
  // Individual runs with explicit colors override this via inline styles.
  const defaultTextColor = toCSS({ type: "scheme", token: "dk1" }, themeColors);

  const canvasStyle: React.CSSProperties = {
    position: "relative",
    width: `${width}pt`,
    height: `${height}pt`,
    background: bgColor,
    color: defaultTextColor,
    transformOrigin: "top left",
    transform: `scale(${zoom})`,
    overflow: "hidden",
    flexShrink: 0,
    fontFamily: "sans-serif",
  };

  return (
    <div
      className={className}
      style={{
        width: `${width * zoom}pt`,
        height: `${height * zoom}pt`,
        position: "relative",
        ...style,
      }}
    >
      <div style={canvasStyle} role="img" aria-label={`Slide ${slide.index + 1}`}>
        {slide.elements.map((element) => (
          <SlideElementRenderer
            key={element.id}
            element={element}
            theme={themeColors}
            renderElement={renderElement}
          />
        ))}
        {children}
      </div>
    </div>
  );
}
