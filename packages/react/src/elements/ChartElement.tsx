import type { ChartShape, ThemeColors } from "@pptx/parser";
import { elementStyle } from "../render/transform";

/**
 * Placeholder chart renderer.
 *
 * Chart XML parsing and rendering is complex and out of scope for the core
 * package. This renders a placeholder so the layout is preserved.
 *
 * Consumers can replace this via the `renderElement` prop on Presentation.Slide.
 */
export function ChartElement({ element }: { element: ChartShape; theme: ThemeColors }) {
  return (
    <div
      style={{
        ...elementStyle(element),
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        background: "#f4f4f5",
        border: "1px dashed #d1d5db",
        color: "#6b7280",
        fontSize: "10pt",
        fontFamily: "sans-serif",
      }}
      data-element-type="chart"
      data-element-id={element.id}
    >
      Chart
    </div>
  );
}
