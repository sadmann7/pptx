import type { ChartDataPoint, ChartSeries, ChartType, ParsedChart } from "../types";
import { attr, attrNum, get, parseXml, toArray } from "../xml";

// ─── Top-level entry ─────────────────────────────────────────────────────────

export function parseChartXml(xml: string): ParsedChart | undefined {
  if (!xml) return undefined;
  try {
    const doc = parseXml(xml) as Record<string, unknown>;
    const chartSpace = get(doc, "c:chartSpace") as Record<string, unknown> | undefined;
    if (!chartSpace) return undefined;

    const chart = get(chartSpace, "c:chart") as Record<string, unknown> | undefined;
    if (!chart) return undefined;

    const title = parseTitle(get(chart, "c:title"));
    const legendNode = get(chart, "c:legend");
    const showLegend = legendNode !== undefined;

    const plotArea = get(chart, "c:plotArea") as Record<string, unknown> | undefined;
    if (!plotArea) return undefined;

    if ("c:barChart" in plotArea) {
      return parseBarChart(plotArea, title, showLegend);
    }
    if ("c:pieChart" in plotArea) {
      return parsePieChart(plotArea, title, showLegend);
    }
    if ("c:lineChart" in plotArea) {
      return parseLineChart(plotArea, title, showLegend);
    }
    if ("c:areaChart" in plotArea) {
      return parseLineChart(plotArea, title, showLegend, "area");
    }

    return { type: "unknown", title, series: [], showLegend };
  } catch {
    return undefined;
  }
}

// ─── Title ───────────────────────────────────────────────────────────────────

function parseTitle(titleNode: unknown): string | undefined {
  if (!titleNode || typeof titleNode !== "object") return undefined;
  const tx = get(titleNode as Record<string, unknown>, "c:tx");
  const rich = get(tx as Record<string, unknown> | undefined, "c:rich");
  if (!rich) return undefined;
  // Collect text from all runs
  const parts: string[] = [];
  for (const p of toArray((rich as Record<string, unknown>)["a:p"]) as unknown[]) {
    for (const r of toArray((p as Record<string, unknown>)?.["a:r"]) as unknown[]) {
      const t = (r as Record<string, unknown>)?.["a:t"];
      if (typeof t === "string") parts.push(t);
    }
  }
  return parts.join("") || undefined;
}

// ─── Series helpers ──────────────────────────────────────────────────────────

function seriesName(serNode: Record<string, unknown>): string {
  const tx = get(serNode, "c:tx");
  if (!tx) return "";
  const strRef = get(tx as Record<string, unknown>, "c:strRef");
  const cache = get(strRef as Record<string, unknown> | undefined, "c:strCache");
  if (!cache) return "";
  const pts = toArray((cache as Record<string, unknown>)["c:pt"]);
  const first = pts[0] as Record<string, unknown> | undefined;
  const v = first?.["c:v"];
  return typeof v === "string" ? v : String(v ?? "");
}

function seriesColor(serNode: Record<string, unknown>): string | undefined {
  const spPr = get(serNode, "c:spPr") as Record<string, unknown> | undefined;
  const solidFill = get(spPr, "a:solidFill") as Record<string, unknown> | undefined;
  const srgb = get(solidFill, "a:srgbClr") as Record<string, unknown> | undefined;
  const hex = attr(srgb, "val");
  return hex ? `#${hex}` : undefined;
}

function parseCategories(serNode: Record<string, unknown>): string[] {
  const cat = get(serNode, "c:cat") as Record<string, unknown> | undefined;
  const ref = (get(cat, "c:strRef") ?? get(cat, "c:numRef")) as Record<string, unknown> | undefined;
  const cache = (get(ref, "c:strCache") ?? get(ref, "c:numCache")) as
    | Record<string, unknown>
    | undefined;
  if (!cache) return [];
  const pts = toArray((cache as Record<string, unknown>)["c:pt"]);
  const result: string[] = [];
  for (const pt of pts as unknown[]) {
    const v = (pt as Record<string, unknown>)?.["c:v"];
    result.push(typeof v === "string" ? v : String(v ?? ""));
  }
  return result;
}

function parseValues(serNode: Record<string, unknown>): number[] {
  const val = get(serNode, "c:val") as Record<string, unknown> | undefined;
  const numRef = get(val, "c:numRef") as Record<string, unknown> | undefined;
  const cache = get(numRef, "c:numCache") as Record<string, unknown> | undefined;
  if (!cache) return [];
  const pts = toArray((cache as Record<string, unknown>)["c:pt"]);
  const result: number[] = [];
  let ptIdx = 0;
  for (const pt of pts as unknown[]) {
    const n = pt as Record<string, unknown>;
    const idx = attrNum(n, "idx") ?? ptIdx;
    while (result.length < idx) result.push(0);
    const v = n["c:v"];
    result.push(typeof v === "number" ? v : parseFloat(String(v ?? "0")));
    ptIdx = idx + 1;
  }
  return result;
}

/** Per-point colors — used for pie charts where each slice has its own color. */
function parsePointColors(serNode: Record<string, unknown>, count: number): (string | undefined)[] {
  const colors: (string | undefined)[] = Array.from({ length: count });
  const dPts = toArray(serNode["c:dPt"]) as unknown[];
  for (const dPt of dPts) {
    const d = dPt as Record<string, unknown>;
    const idx = attrNum(d, "idx") ?? attrNum(get(d, "c:idx"), "val");
    if (idx == null || idx >= count) continue;
    const spPr = get(d, "c:spPr") as Record<string, unknown> | undefined;
    const solidFill = get(spPr, "a:solidFill") as Record<string, unknown> | undefined;
    const srgb = get(solidFill, "a:srgbClr") as Record<string, unknown> | undefined;
    const hex = attr(srgb, "val");
    if (hex) colors[idx] = `#${hex}`;
  }
  return colors;
}

// ─── Bar chart ───────────────────────────────────────────────────────────────

function parseBarChart(
  plotArea: Record<string, unknown>,
  title: string | undefined,
  showLegend: boolean,
): ParsedChart {
  const barChartNode = get(plotArea, "c:barChart") as Record<string, unknown>;
  const barDirAttr = attr(
    get(barChartNode, "c:barDir") as Record<string, unknown> | undefined,
    "val",
  );
  const barDirection: "horizontal" | "vertical" = barDirAttr === "bar" ? "horizontal" : "vertical";
  const groupingAttr = attr(
    get(barChartNode, "c:grouping") as Record<string, unknown> | undefined,
    "val",
  );
  const grouping =
    groupingAttr === "stacked"
      ? "stacked"
      : groupingAttr === "percentStacked"
        ? "percentStacked"
        : "clustered";

  const serNodes = toArray(barChartNode["c:ser"]) as Record<string, unknown>[];
  const series: ChartSeries[] = serNodes.map((s) => {
    const cats = parseCategories(s);
    const vals = parseValues(s);
    const data: ChartDataPoint[] = cats.map((label, i) => ({
      label,
      value: vals[i] ?? 0,
    }));
    return { name: seriesName(s), color: seriesColor(s), data };
  });

  const legendEntries = showLegend
    ? series.map((s) => ({ name: s.name, color: s.color ?? "#888" }))
    : undefined;

  return {
    type: "bar",
    title,
    series,
    barDirection,
    grouping,
    showLegend,
    ...(legendEntries ? { legendEntries } : {}),
  };
}

// ─── Pie chart ───────────────────────────────────────────────────────────────

const PIE_FALLBACK_COLORS = [
  "#5DA5DA",
  "#FAA43A",
  "#60BD68",
  "#F17CB0",
  "#B2912F",
  "#B276B2",
  "#DECF3F",
  "#F15854",
  "#4D4D4D",
];

function parsePieChart(
  plotArea: Record<string, unknown>,
  title: string | undefined,
  showLegend: boolean,
): ParsedChart {
  const pieChartNode = get(plotArea, "c:pieChart") as Record<string, unknown>;
  const serNodes = toArray(pieChartNode["c:ser"]) as Record<string, unknown>[];

  const series: ChartSeries[] = serNodes.map((s) => {
    const cats = parseCategories(s);
    const vals = parseValues(s);
    const pointColors = parsePointColors(s, cats.length);
    const data: ChartDataPoint[] = cats.map((label, i) => ({
      label,
      value: vals[i] ?? 0,
      color: pointColors[i] ?? PIE_FALLBACK_COLORS[i % PIE_FALLBACK_COLORS.length],
    }));
    return { name: seriesName(s), data };
  });

  const legendEntries = showLegend
    ? (series[0]?.data ?? []).map((d, i) => ({
        name: d.label,
        color: d.color ?? PIE_FALLBACK_COLORS[i % PIE_FALLBACK_COLORS.length]!,
      }))
    : undefined;

  return {
    type: "pie",
    title,
    series,
    showLegend,
    ...(legendEntries ? { legendEntries } : {}),
  };
}

// ─── Line / Area chart ───────────────────────────────────────────────────────

function parseLineChart(
  plotArea: Record<string, unknown>,
  title: string | undefined,
  showLegend: boolean,
  forceType: ChartType = "line",
): ParsedChart {
  const key = forceType === "area" ? "c:areaChart" : "c:lineChart";
  const chartNode = get(plotArea, key) as Record<string, unknown>;
  const serNodes = toArray(chartNode["c:ser"]) as Record<string, unknown>[];

  const series: ChartSeries[] = serNodes.map((s) => {
    const cats = parseCategories(s);
    const vals = parseValues(s);
    const data: ChartDataPoint[] = cats.map((label, i) => ({
      label,
      value: vals[i] ?? 0,
    }));
    return { name: seriesName(s), color: seriesColor(s), data };
  });

  const legendEntries = showLegend
    ? series.map((s) => ({ name: s.name, color: s.color ?? "#888" }))
    : undefined;

  return {
    type: forceType,
    title,
    series,
    showLegend,
    ...(legendEntries ? { legendEntries } : {}),
  };
}
