import React from "react";
import type { ArrowEnd, ConnectorShape, ThemeColors } from "@pptx/parser";
import { strokeToSVGAttrs } from "../render/color";
import { getShapePath } from "../render/shapes";
import { elementStyle } from "../render/transform";
import { renderShapeElement } from "./ShapeElement";

interface ConnectorElementProps {
  element: ConnectorShape;
  theme: ThemeColors;
}

export function ConnectorElement({ element, theme }: ConnectorElementProps) {
  const base = elementStyle(element);
  const outer: React.CSSProperties = {
    ...base,
    // Never clip the stroke — connectors can have near-zero bounding-box height.
    overflow: "visible",
    pointerEvents: "none",
    // Guarantee a visible area even when the bounding box is essentially 0-height.
    minHeight: "1pt",
  };

  const strokeAttrs = strokeToSVGAttrs(element.stroke, theme);
  const strokeColor = strokeAttrs.stroke !== "none" ? strokeAttrs.stroke : "#000000";
  const strokeWidth = strokeAttrs.stroke !== "none" ? strokeAttrs.strokeWidth : "0.75pt";

  const { headEnd, tailEnd } = element.stroke ?? {};
  const headMarkerId = headEnd ? `${element.id}-head` : undefined;
  const tailMarkerId = tailEnd ? `${element.id}-tail` : undefined;

  const shape = getShapePath(element.shapeType ?? "line");

  const shapeProps: Record<string, string | number | undefined> = {
    fill: "none",
    stroke: strokeColor,
    strokeWidth,
    ...(strokeAttrs.strokeDasharray ? { strokeDasharray: strokeAttrs.strokeDasharray } : {}),
    ...(headMarkerId ? { markerStart: `url(#${headMarkerId})` } : {}),
    ...(tailMarkerId ? { markerEnd: `url(#${tailMarkerId})` } : {}),
  };

  return (
    <div style={outer} data-element-type="connector" data-element-id={element.id}>
      <svg
        style={{
          position: "absolute",
          inset: 0,
          width: "100%",
          height: "100%",
          overflow: "visible",
        }}
        viewBox="0 0 100 100"
        preserveAspectRatio="none"
        aria-hidden="true"
      >
        {(headMarkerId || tailMarkerId) && (
          <defs>
            {headMarkerId && headEnd && (
              <ArrowMarker
                id={headMarkerId}
                end={headEnd}
                color={strokeColor}
                orient="auto-start-reverse"
              />
            )}
            {tailMarkerId && tailEnd && (
              <ArrowMarker id={tailMarkerId} end={tailEnd} color={strokeColor} orient="auto" />
            )}
          </defs>
        )}
        {renderShapeElement(shape, shapeProps)}
      </svg>
    </div>
  );
}

// ── Arrow marker ──────────────────────────────────────────────────────────────

interface ArrowMarkerProps {
  id: string;
  end: ArrowEnd;
  color: string;
  orient: string;
}

function arrowScale(size: ArrowEnd["width"]): number {
  if (size === "sm") return 0.6;
  if (size === "lg") return 1.6;
  return 1;
}

function ArrowMarker({ id, end, color, orient }: ArrowMarkerProps) {
  const ws = arrowScale(end.width);
  const ls = arrowScale(end.length);
  const w = 10 * ws;
  const h = 7 * ws;
  const refX = ls * 9 * ws;

  const p = arrowPath(end.type, w, h);
  if (!p) return null;

  return (
    <marker
      id={id}
      markerWidth={w * ls}
      markerHeight={h}
      refX={refX}
      refY={h / 2}
      orient={orient}
      markerUnits="strokeWidth"
    >
      {p.type === "polygon" ? <polygon points={p.d} fill={color} /> : <path d={p.d} fill={color} />}
    </marker>
  );
}

function arrowPath(
  type: ArrowEnd["type"],
  w: number,
  h: number,
): { type: "polygon" | "path"; d: string } | null {
  const hw = h / 2;
  switch (type) {
    case "triangle":
    case "arrow":
      return { type: "polygon", d: `0 0, ${w} ${hw}, 0 ${h}` };
    case "stealth":
      return {
        type: "polygon",
        d: `0 0, ${w} ${hw}, 0 ${h}, ${w * 0.4} ${hw}`,
      };
    case "diamond":
      return {
        type: "polygon",
        d: `0 ${hw}, ${w / 2} 0, ${w} ${hw}, ${w / 2} ${h}`,
      };
    case "oval":
      return {
        type: "path",
        d: `M 0 ${hw} A ${w / 2} ${hw} 0 1 1 0 ${hw + 0.001} Z`,
      };
    default:
      return null;
  }
}
