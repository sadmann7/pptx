import React from "react";
import type { GeometricShape, ThemeColors } from "@pptx/parser";
import { effectsToFilter, fillToSVG, strokeToSVGAttrs } from "../render/color";
import { bodyStyle } from "../render/text";
import { getShapePath } from "../render/shapes";
import { elementStyle } from "../render/transform";
import { ParagraphElement } from "./shared/ParagraphElement";

interface ShapeElementProps {
  element: GeometricShape;
  theme: ThemeColors;
}

export function ShapeElement({ element, theme }: ShapeElementProps) {
  const shape = getShapePath(
    element.shapeType,
    element.size.width,
    element.size.height,
    element.adjustments,
  );
  const fill = fillToSVG(element.fill, theme);
  const strokeAttrs = strokeToSVGAttrs(element.stroke, theme);

  const isLine =
    element.shapeType === "line" ||
    element.shapeType === "arc" ||
    element.shapeType.toLowerCase().includes("connector");

  const filter = effectsToFilter(element.effects, theme);
  const outer: React.CSSProperties = {
    ...elementStyle(element),
    position: "absolute",
    ...(isLine ? { overflow: "visible", minHeight: "1pt" } : {}),
    ...(filter ? { filter } : {}),
  };

  const svgStyle: React.CSSProperties = {
    position: "absolute",
    inset: 0,
    width: "100%",
    height: "100%",
    overflow: "visible",
  };

  const sharedProps = {
    fill: isLine ? "none" : fill,
    stroke: strokeAttrs.stroke,
    strokeWidth: strokeAttrs.strokeWidth,
    ...(strokeAttrs.strokeDasharray ? { strokeDasharray: strokeAttrs.strokeDasharray } : {}),
  };

  return (
    <div style={outer} data-element-type="shape" data-element-id={element.id}>
      <svg style={svgStyle} viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true">
        {renderShapeElement(shape, sharedProps)}
      </svg>

      {element.body && (
        <div
          style={{
            position: "absolute",
            inset: 0,
            ...bodyStyle(element.body.properties, theme),
          }}
        >
          {element.body.paragraphs.map((p, i) => (
            <ParagraphElement key={i} paragraph={p} theme={theme} />
          ))}
        </div>
      )}
    </div>
  );
}

export function renderShapeElement(
  shape: ReturnType<typeof getShapePath>,
  sharedProps: Record<string, string | number | undefined>,
): React.ReactElement {
  // vectorEffect="non-scaling-stroke" prevents the SVG viewBox transform from
  // scaling the stroke-width. Without it a 1pt stroke on a shape that maps to
  // a 2.4× viewBox scale renders 2.4pt thick — looking like a thick border.
  const allProps = { ...shape.attrs, ...sharedProps, vectorEffect: "non-scaling-stroke" };

  switch (shape.element) {
    case "rect":
      return <rect {...allProps} rx={shape.rx} ry={shape.ry} />;
    case "ellipse":
      return <ellipse {...allProps} />;
    case "polygon":
      return <polygon {...allProps} />;
    case "path":
      return <path {...allProps} />;
    case "line":
      return <line {...allProps} />;
  }
}
