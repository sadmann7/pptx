import { SafeXmlNode } from "../../ooxml/xml";
import { RenderContext } from "../context";
import { resolveColor, resolveColorToCss } from "../style";
import { extractChartLineStyle } from "./style";
import type { ChartFrameStyle } from "./type";

// Authoring tools write "let the slide show through" as a `solidFill` with
// `a:alpha` 0 rather than `noFill`; painting it would plate the chart, usually black.
function resolveBackgroundFill(fill: SafeXmlNode, ctx: RenderContext): string | undefined {
  if (resolveColor(fill, ctx).alpha <= 0) return undefined;
  return resolveColorToCss(fill, ctx);
}

export function extractBackgroundColors(
  chartXml: SafeXmlNode,
  chartNode: SafeXmlNode,
  ctx: RenderContext,
): { chartBg?: string; plotAreaBg?: string } {
  let chartBg: string | undefined;
  let plotAreaBg: string | undefined;

  const chartSpaceSpPr = chartXml.child("spPr");
  if (chartSpaceSpPr.exists()) {
    const noFill = chartSpaceSpPr.child("noFill");
    if (!noFill.exists()) {
      const fill = chartSpaceSpPr.child("solidFill");
      if (fill.exists()) {
        chartBg = resolveBackgroundFill(fill, ctx);
      } else {
        chartBg = "#ffffff";
      }
    }
  }

  const plotArea = chartNode.child("plotArea");
  if (plotArea.exists()) {
    const plotSpPr = plotArea.child("spPr");
    if (plotSpPr.exists()) {
      const noFill = plotSpPr.child("noFill");
      if (!noFill.exists()) {
        const fill = plotSpPr.child("solidFill");
        if (fill.exists()) {
          plotAreaBg = resolveBackgroundFill(fill, ctx);
        }
      }
    }
  }

  return { chartBg, plotAreaBg };
}

export function extractChartFrameStyle(
  chartXml: SafeXmlNode,
  ctx: RenderContext,
): ChartFrameStyle | undefined {
  const lineStyle = extractChartLineStyle(chartXml.child("spPr").child("ln"), ctx);
  if (!lineStyle) return undefined;

  return {
    borderColor: lineStyle.color,
    borderWidth: lineStyle.width,
    borderStyle: lineStyle.type,
  };
}
