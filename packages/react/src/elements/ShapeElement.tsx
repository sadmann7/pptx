import React from "react";
import type { GeometricShape, ThemeColors } from "@pptx/parser";
import {
  effectsToBoxShadow,
  effectsToFilter,
  fillToSVG,
  strokeToSVGAttrs,
} from "../render/color";
import { bodyStyle } from "../render/text";
import { getShapePath } from "../render/shapes";
import { elementStyle } from "../render/transform";
import { ParagraphElement } from "./shared/ParagraphElement";

interface ShapeElementProps {
  element: GeometricShape;
  theme: ThemeColors;
}

export function ShapeElement({ element, theme }: ShapeElementProps) {
  const shape = getShapePath(element.shapeType);
  const fill = fillToSVG(element.fill, theme);
  const strokeAttrs = strokeToSVGAttrs(element.stroke, theme);

  // Shapes that are just lines / connectors have no fill and may be near-zero height.
  const isLine =
    element.shapeType === "line" ||
    element.shapeType === "arc" ||
    element.shapeType.toLowerCase().includes("connector");

  // box-shadow works well for rect/roundRect; filter: drop-shadow for complex shapes.
  const isBoxShadowShape =
    element.shapeType === "rect" || element.shapeType === "roundRect";
  const boxShadow = isBoxShadowShape
    ? effectsToBoxShadow(element.effects, theme)
    : undefined;
  const filter = !isBoxShadowShape
    ? effectsToFilter(element.effects, theme)
    : undefined;
  const outer: React.CSSProperties = {
    ...elementStyle(element),
    position: "absolute",
    ...(isLine ? { overflow: "visible", minHeight: "1pt" } : {}),
    ...(boxShadow ? { boxShadow } : {}),
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
  const allProps = { ...shape.attrs, ...sharedProps };

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
