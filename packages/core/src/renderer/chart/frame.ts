import { SafeXmlNode } from "../../ooxml/xml";
import { RenderContext } from "../context";
import { extractChartLineStyle, resolveChartColor, resolveChartFill } from "./style";
import type { ChartFrameStyle } from "./type";

/**
 * A background a deck paints at zero opacity, which authoring tools write as a
 * `solidFill` rather than a `noFill`. Treating it as a color would plate the
 * chart in whatever hue happens to carry the alpha, usually black.
 */
function isTransparentFill(fill: SafeXmlNode, ctx: RenderContext): boolean {
  const resolved = resolveChartFill(fill, ctx);
  return resolved !== undefined && resolved.alpha <= 0;
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
        chartBg = isTransparentFill(fill, ctx) ? undefined : resolveChartColor(fill, ctx);
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
          plotAreaBg = isTransparentFill(fill, ctx) ? undefined : resolveChartColor(fill, ctx);
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
