import { SafeXmlNode } from '../../parser/XmlParser';
import { RenderContext } from '../RenderContext';
import { extractChartLineStyle, resolveColorToHex } from './style';
import type { ChartFrameStyle } from './types';

export function extractBackgroundColors(
  chartXml: SafeXmlNode,
  chartNode: SafeXmlNode,
  ctx: RenderContext,
): { chartBg?: string; plotAreaBg?: string } {
  let chartBg: string | undefined;
  let plotAreaBg: string | undefined;

  const chartSpaceSpPr = chartXml.child('spPr');
  if (chartSpaceSpPr.exists()) {
    const noFill = chartSpaceSpPr.child('noFill');
    if (!noFill.exists()) {
      const fill = chartSpaceSpPr.child('solidFill');
      if (fill.exists()) {
        chartBg = resolveColorToHex(fill, ctx);
      } else {
        chartBg = '#ffffff';
      }
    }
  }

  const plotArea = chartNode.child('plotArea');
  if (plotArea.exists()) {
    const plotSpPr = plotArea.child('spPr');
    if (plotSpPr.exists()) {
      const noFill = plotSpPr.child('noFill');
      if (!noFill.exists()) {
        const fill = plotSpPr.child('solidFill');
        if (fill.exists()) {
          plotAreaBg = resolveColorToHex(fill, ctx);
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
  const lineStyle = extractChartLineStyle(chartXml.child('spPr').child('ln'), ctx);
  if (!lineStyle) return undefined;

  return {
    borderColor: lineStyle.color,
    borderWidth: lineStyle.width,
    borderStyle: lineStyle.type,
  };
}
