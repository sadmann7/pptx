import { SafeXmlNode } from "../../ooxml/xml";
import { RenderContext } from "../context";
import { parseOoxmlBoolElement } from "./boolean";
import { formatValue } from "./format";
import { extractChartLineStyle, resolveChartColor } from "./style";
import { extractTitleText, extractTitleTextStyle, extractTxPrStyle } from "./text";
import {
  type AxisInfo,
  type ChartLineStyle,
  DEFAULT_CHART_AXIS_LABEL_FONT_SIZE,
  DEFAULT_CHART_AXIS_LINE_COLOR,
  DEFAULT_CHART_FOREGROUND_COLOR,
  DEFAULT_MAJOR_GRIDLINE_STYLE,
} from "./type";

const DEFAULT_AXIS_INFO: AxisInfo = {
  deleted: false,
  tickLblPos: "nextTo",
  hasMajorGridlines: false,
  orientation: "minMax",
};

function themeHex(ctx: RenderContext, key: string): string | undefined {
  const value = ctx.theme.colorScheme.get(key);
  return value?.replace("#", "").toUpperCase();
}

function legacyOfficeImplicitAxisColor(ctx: RenderContext): string | undefined {
  // Office 2007/2010 default chart themes use black implicit axis/grid lines,
  // while newer Office themes use the lighter gray default.
  if (themeHex(ctx, "accent1") === "4F81BD" && themeHex(ctx, "accent2") === "C0504D") {
    return "#000000";
  }
  return undefined;
}

function extractAxisLabelColor(ax: SafeXmlNode, ctx: RenderContext): string | undefined {
  const txPr = ax.child("txPr");
  if (!txPr.exists()) return undefined;

  for (const p of txPr.children("p")) {
    const pPr = p.child("pPr");
    if (!pPr.exists()) continue;
    const defRPr = pPr.child("defRPr");
    if (!defRPr.exists()) continue;
    const fill = defRPr.child("solidFill");
    if (fill.exists()) {
      return resolveChartColor(fill, ctx);
    }
  }
  return undefined;
}

function extractAxisLineColor(ax: SafeXmlNode, ctx: RenderContext): string | undefined {
  const ln = ax.child("spPr").child("ln");
  if (!ln.exists()) return undefined;
  const fill = ln.child("solidFill");
  if (!fill.exists()) return undefined;
  return resolveChartColor(fill, ctx);
}

function isAxisLineHidden(ax: SafeXmlNode): boolean {
  return ax.child("spPr").child("ln").child("noFill").exists();
}

function extractMajorGridlineStyle(
  ax: SafeXmlNode,
  ctx: RenderContext,
): ChartLineStyle | undefined {
  const ln = ax.child("majorGridlines").child("spPr").child("ln");
  return extractChartLineStyle(ln, ctx);
}

function extractTitleRotation(title: SafeXmlNode): number | undefined {
  const bodyPr = title.child("tx").child("rich").child("bodyPr").exists()
    ? title.child("tx").child("rich").child("bodyPr")
    : title.child("txPr").child("bodyPr");
  const rot = bodyPr.numAttr("rot");
  if (rot === undefined) return undefined;
  const deg = rot / 60000;
  return Number(deg.toFixed(3));
}

function extractAxisTitle(
  ax: SafeXmlNode,
  ctx: RenderContext,
): Pick<AxisInfo, "title" | "titleStyle" | "titleRotation"> {
  const title = ax.child("title");
  if (!title.exists()) return {};

  const text = extractTitleText(title);
  if (!text) return {};

  return {
    title: text,
    titleStyle: extractTitleTextStyle(title, ctx),
    titleRotation: extractTitleRotation(title),
  };
}

function parseAxisNode(ax: SafeXmlNode, ctx: RenderContext): AxisInfo {
  if (!ax.exists()) return { ...DEFAULT_AXIS_INFO };
  const deleted = parseOoxmlBoolElement(ax.child("delete"));
  const tickLblPos = ax.child("tickLblPos").attr("val") || "nextTo";
  const crosses = ax.child("crosses").attr("val");
  const numFmtNode = ax.child("numFmt");
  const numFmt = numFmtNode.exists() ? numFmtNode.attr("formatCode") || undefined : undefined;
  const scaling = ax.child("scaling");
  const minNode = scaling.child("min");
  const maxNode = scaling.child("max");
  const min = minNode.exists() ? parseFloat(minNode.attr("val") || "") : undefined;
  const max = maxNode.exists() ? parseFloat(maxNode.attr("val") || "") : undefined;
  const hasMajorGridlines = ax.child("majorGridlines").exists();
  const majorTickMark = ax.child("majorTickMark").attr("val");
  const orientation = scaling.child("orientation").attr("val") || "minMax";
  const txStyle = extractTxPrStyle(ax, ctx);
  const labelColor = txStyle?.color ?? extractAxisLabelColor(ax, ctx);
  const labelFontSize = txStyle?.fontSize;
  const implicitAxisColor = legacyOfficeImplicitAxisColor(ctx);
  const lineHidden = isAxisLineHidden(ax);
  const lineColor = lineHidden ? undefined : (extractAxisLineColor(ax, ctx) ?? implicitAxisColor);
  const majorGridlineStyle = hasMajorGridlines
    ? (extractMajorGridlineStyle(ax, ctx) ??
      (implicitAxisColor
        ? { ...DEFAULT_MAJOR_GRIDLINE_STYLE, color: implicitAxisColor }
        : undefined))
    : undefined;
  const axisTitle = extractAxisTitle(ax, ctx);
  return {
    deleted,
    tickLblPos,
    crosses,
    numFmt: numFmt && numFmt !== "General" ? numFmt : undefined,
    min: min !== undefined && !isNaN(min) ? min : undefined,
    max: max !== undefined && !isNaN(max) ? max : undefined,
    hasMajorGridlines,
    majorTickMark,
    orientation,
    ...axisTitle,
    labelColor,
    labelFontSize,
    lineColor,
    lineHidden,
    majorGridlineStyle,
  };
}

export function getChartAxisIds(chartTypeNode?: SafeXmlNode): string[] {
  if (!chartTypeNode?.exists()) return [];
  return chartTypeNode
    .children("axId")
    .map((ax) => ax.attr("val"))
    .filter((id): id is string => id !== undefined && id !== "");
}

function findAxisById(
  plotArea: SafeXmlNode,
  axisNames: readonly string[],
  axisId: string | undefined,
): SafeXmlNode {
  if (axisId) {
    for (const axisName of axisNames) {
      const axes = plotArea.children(axisName);
      const matched = axes.find((axis) => axis.child("axId").attr("val") === axisId);
      if (matched) return matched;
    }
    return new SafeXmlNode(null);
  }

  for (const axisName of axisNames) {
    const axes = plotArea.children(axisName);
    if (axes[0]?.exists()) return axes[0];
  }
  return new SafeXmlNode(null);
}

export function parseAxes(
  plotArea: SafeXmlNode,
  ctx: RenderContext,
  chartTypeNode?: SafeXmlNode,
): { valueAxis: AxisInfo; categoryAxis: AxisInfo } {
  const axisIds = getChartAxisIds(chartTypeNode);
  const categoryAxisId = axisIds[0];
  const valueAxisId = axisIds[1];
  const valAx = findAxisById(plotArea, ["valAx"], valueAxisId);
  const catAx = findAxisById(plotArea, ["catAx", "dateAx"], categoryAxisId);
  return {
    valueAxis: parseAxisNode(valAx, ctx),
    categoryAxis: parseAxisNode(catAx, ctx),
  };
}

export function parseScatterAxes(
  plotArea: SafeXmlNode,
  ctx: RenderContext,
): { xAxis: AxisInfo; yAxis: AxisInfo } {
  const allValAx = plotArea.children("valAx");
  let xAxis: AxisInfo = { ...DEFAULT_AXIS_INFO };
  let yAxis: AxisInfo = { ...DEFAULT_AXIS_INFO };
  for (const ax of allValAx) {
    const axPos = ax.child("axPos").attr("val") ?? "";
    const info = parseAxisNode(ax, ctx);
    if (axPos === "b" || axPos === "t") {
      xAxis = info;
    } else if (axPos === "l" || axPos === "r") {
      yAxis = info;
    }
  }
  if (allValAx.length === 1) {
    yAxis = parseAxisNode(allValAx[0], ctx);
  }
  return { xAxis, yAxis };
}

/**
 * Room an axis title needs beyond its tick labels. ECharts' `containLabel`
 * grows the grid for labels but never for the axis name, so a grid that pays
 * only for labels pushes the title into the legend or past the chart's clip.
 */
export function getAxisTitleSpacePx(info: AxisInfo): number {
  if (!info.title || info.deleted) return 0;
  const fontSize = info.titleStyle?.fontSize ?? 10;
  return Math.round(fontSize * 1.8);
}

export function applyAxisInfo(
  axisDef: Record<string, unknown>,
  info: AxisInfo,
  kind: "value" | "category",
): void {
  if (info.deleted) {
    axisDef.axisLabel = { ...(axisDef.axisLabel as object), show: false };
    axisDef.axisLine = { show: false };
    axisDef.axisTick = { show: false };
    if (kind === "value") axisDef.splitLine = { show: false };
    return;
  }

  if (info.orientation === "maxMin") {
    axisDef.inverse = true;
  }

  if (info.crosses === "autoZero") {
    const existingLine = (axisDef.axisLine as Record<string, unknown>) || {};
    axisDef.axisLine = { ...existingLine, onZero: true };
  }

  if (info.title) {
    axisDef.name = info.title;
    axisDef.nameLocation = "middle";
    axisDef.nameGap = kind === "value" ? 42 : 24;
    if (info.titleRotation !== undefined) {
      axisDef.nameRotate = info.titleRotation;
    }

    const nameTextStyle: Record<string, unknown> = {};
    if (info.titleStyle?.color) nameTextStyle.color = info.titleStyle.color;
    if (info.titleStyle?.fontSize !== undefined) nameTextStyle.fontSize = info.titleStyle.fontSize;
    if (info.titleStyle?.fontFamily) nameTextStyle.fontFamily = info.titleStyle.fontFamily;
    if (info.titleStyle?.bold !== undefined) {
      nameTextStyle.fontWeight = info.titleStyle.bold ? "bold" : "normal";
    }
    if (Object.keys(nameTextStyle).length > 0) {
      axisDef.nameTextStyle = nameTextStyle;
    }
  }

  if (info.tickLblPos === "none") {
    axisDef.axisLabel = { ...(axisDef.axisLabel as object), show: false };
  }

  if (info.majorTickMark === "none" || info.lineHidden) {
    const existingTick = (axisDef.axisTick as Record<string, unknown>) || {};
    axisDef.axisTick = { ...existingTick, show: false };
  } else if (!info.deleted) {
    const existingTick = (axisDef.axisTick as Record<string, unknown>) || {};
    const existingLineStyle = (existingTick.lineStyle as Record<string, unknown>) || {};
    if (existingLineStyle.color === undefined) {
      axisDef.axisTick = {
        ...existingTick,
        lineStyle: { ...existingLineStyle, color: info.lineColor ?? DEFAULT_CHART_AXIS_LINE_COLOR },
      };
    }
  }

  if (kind === "value") {
    if (info.min !== undefined) axisDef.min = info.min;
    if (info.max !== undefined) axisDef.max = info.max;
  }

  if (kind === "value" && !info.deleted && info.tickLblPos !== "none") {
    const existingLabel = (axisDef.axisLabel as Record<string, unknown>) || {};
    if (!existingLabel.formatter) {
      const nf = info.numFmt;
      axisDef.axisLabel = {
        ...existingLabel,
        formatter: (val: number) => formatValue(val, nf),
      };
    }
  }

  if (!info.deleted && info.tickLblPos !== "none") {
    const existingLabel = (axisDef.axisLabel as Record<string, unknown>) || {};
    if (existingLabel.fontSize === undefined) {
      axisDef.axisLabel = {
        ...existingLabel,
        fontSize: DEFAULT_CHART_AXIS_LABEL_FONT_SIZE,
      };
    }
  }

  if (kind === "value") {
    if (!info.hasMajorGridlines) {
      axisDef.splitLine = { show: false };
    } else if (info.majorGridlineStyle) {
      const existingSplitLine = (axisDef.splitLine as Record<string, unknown>) || {};
      const existingLineStyle = (existingSplitLine.lineStyle as Record<string, unknown>) || {};
      axisDef.splitLine = {
        ...existingSplitLine,
        show: true,
        lineStyle: { ...existingLineStyle, ...info.majorGridlineStyle },
      };
    } else {
      const existingSplitLine = (axisDef.splitLine as Record<string, unknown>) || {};
      const existingLineStyle = (existingSplitLine.lineStyle as Record<string, unknown>) || {};
      axisDef.splitLine = {
        ...existingSplitLine,
        show: true,
        lineStyle: { ...DEFAULT_MAJOR_GRIDLINE_STYLE, ...existingLineStyle },
      };
    }
  }

  if (info.labelColor || !info.deleted) {
    const existingLabel = (axisDef.axisLabel as Record<string, unknown>) || {};
    const color =
      info.labelColor ??
      (existingLabel.color === undefined ? DEFAULT_CHART_FOREGROUND_COLOR : undefined);
    if (color) {
      axisDef.axisLabel = { ...existingLabel, color };
    }
  }
  if (info.labelFontSize !== undefined) {
    const existingLabel = (axisDef.axisLabel as Record<string, unknown>) || {};
    axisDef.axisLabel = { ...existingLabel, fontSize: info.labelFontSize };
  }

  if (info.lineHidden) {
    const existingLine = (axisDef.axisLine as Record<string, unknown>) || {};
    axisDef.axisLine = { ...existingLine, show: false };
  } else if (info.lineColor || !info.deleted) {
    const existingLine = (axisDef.axisLine as Record<string, unknown>) || {};
    const existingLineStyle = (existingLine.lineStyle as Record<string, unknown>) || {};
    const color =
      info.lineColor ??
      (existingLineStyle.color === undefined ? DEFAULT_CHART_AXIS_LINE_COLOR : undefined);
    if (color) {
      axisDef.axisLine = {
        ...existingLine,
        show: existingLine.show ?? true,
        lineStyle: { ...existingLineStyle, color },
      };
    }
  }
}
