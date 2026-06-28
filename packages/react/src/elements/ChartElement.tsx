import React from "react";
import type { ChartShape, ParsedChart, ThemeColors } from "@pptx/parser";
import { elementStyle } from "../render/transform";

interface ChartElementProps {
  element: ChartShape;
  theme: ThemeColors;
}

export function ChartElement({ element }: ChartElementProps) {
  const outer: React.CSSProperties = {
    ...elementStyle(element),
    position: "absolute",
    overflow: "hidden",
  };

  if (!element.parsedChart) {
    return (
      <div
        style={{
          ...outer,
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

  return (
    <div style={outer} data-element-type="chart" data-element-id={element.id}>
      <ChartRenderer chart={element.parsedChart} />
    </div>
  );
}

// ─── Chart dispatcher ─────────────────────────────────────────────────────────

function ChartRenderer({ chart }: { chart: ParsedChart }) {
  const { title, showLegend, legendEntries } = chart;
  const legendH = showLegend && legendEntries?.length ? 24 : 0;
  const titleH = title ? 28 : 0;
  const padding = 8;

  return (
    <div
      style={{
        width: "100%",
        height: "100%",
        display: "flex",
        flexDirection: "column",
        fontFamily: "sans-serif",
        padding: `${padding}pt`,
        boxSizing: "border-box",
        gap: "4pt",
      }}
    >
      {title && (
        <div
          style={{
            fontSize: "10.5pt",
            fontWeight: 400,
            textAlign: "center",
            flexShrink: 0,
            height: `${titleH}pt`,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
          }}
        >
          {title}
        </div>
      )}

      <div style={{ flex: 1, minHeight: 0, position: "relative" }}>
        {(chart.type === "bar" || chart.type === "line") && <BarChartSVG chart={chart} />}
        {chart.type === "pie" && <PieChartSVG chart={chart} />}
        {chart.type === "unknown" && (
          <div
            style={{
              width: "100%",
              height: "100%",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              color: "#6b7280",
              fontSize: "9pt",
            }}
          >
            Chart
          </div>
        )}
      </div>

      {showLegend && legendEntries && legendEntries.length > 0 && (
        <div
          style={{
            display: "flex",
            flexWrap: "wrap",
            gap: "6pt",
            justifyContent: "center",
            flexShrink: 0,
            height: `${legendH}pt`,
            alignItems: "center",
          }}
        >
          {legendEntries.map((entry, i) => (
            <div key={i} style={{ display: "flex", alignItems: "center", gap: "3pt" }}>
              <div
                style={{
                  width: "8pt",
                  height: "8pt",
                  background: entry.color,
                  flexShrink: 0,
                }}
              />
              <span style={{ fontSize: "8pt", color: "#374151" }}>{entry.name}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ─── Bar chart ────────────────────────────────────────────────────────────────

const FALLBACK_COLORS = [
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

function BarChartSVG({ chart }: { chart: ParsedChart }) {
  const { series } = chart;
  if (!series.length) return null;

  // Collect all categories from the first series
  const categories = series[0]?.data.map((d) => d.label) ?? [];
  const nCats = categories.length;
  const nSeries = series.length;

  // Compute the max value for axis scaling
  const allVals = series.flatMap((s) => s.data.map((d) => d.value));
  const maxVal = Math.max(...allVals, 1);
  const niceMax = niceNumber(maxVal);

  // Layout constants (in SVG user units ≈ pt)
  const leftPad = 28;
  const rightPad = 8;
  const topPad = 6;
  const bottomPad = 28;
  const W = 300;
  const H = 180;
  const plotW = W - leftPad - rightPad;
  const plotH = H - topPad - bottomPad;

  const groupW = plotW / nCats;
  const barW = Math.max(groupW / (nSeries + 1), 2);
  const barGap = barW / 4;

  // Axis tick lines
  const tickCount = 5;
  const ticks = Array.from({ length: tickCount + 1 }, (_, i) =>
    Math.round((niceMax / tickCount) * i),
  );

  return (
    <svg
      viewBox={`0 0 ${W} ${H}`}
      preserveAspectRatio="xMidYMid meet"
      style={{ width: "100%", height: "100%" }}
      aria-hidden="true"
    >
      {/* Grid lines and Y axis ticks */}
      {ticks.map((tick) => {
        const y = topPad + plotH - (tick / niceMax) * plotH;
        return (
          <g key={tick}>
            <line
              x1={leftPad}
              y1={y}
              x2={leftPad + plotW}
              y2={y}
              stroke="#e5e7eb"
              strokeWidth="0.5"
            />
            <text x={leftPad - 3} y={y + 3} textAnchor="end" fontSize="7" fill="#6b7280">
              {tick}
            </text>
          </g>
        );
      })}

      {/* Bars */}
      {series.map((s, si) => {
        const color = s.color ?? FALLBACK_COLORS[si % FALLBACK_COLORS.length];
        return s.data.map((d, ci) => {
          const bH = Math.max((d.value / niceMax) * plotH, 0.5);
          const xOffset =
            leftPad +
            ci * groupW +
            groupW / 2 -
            (nSeries * (barW + barGap)) / 2 +
            si * (barW + barGap);
          const y = topPad + plotH - bH;
          return (
            <rect key={`${si}-${ci}`} x={xOffset} y={y} width={barW} height={bH} fill={color} />
          );
        });
      })}

      {/* Y axis left line */}
      <line
        x1={leftPad}
        y1={topPad}
        x2={leftPad}
        y2={topPad + plotH}
        stroke="#9ca3af"
        strokeWidth="0.5"
      />
      {/* X axis baseline */}
      <line
        x1={leftPad}
        y1={topPad + plotH}
        x2={leftPad + plotW}
        y2={topPad + plotH}
        stroke="#9ca3af"
        strokeWidth="0.5"
      />

      {/* Category labels */}
      {categories.map((cat, ci) => {
        const x = leftPad + ci * groupW + groupW / 2;
        return (
          <text
            key={ci}
            x={x}
            y={topPad + plotH + 10}
            textAnchor="middle"
            fontSize="7"
            fill="#374151"
          >
            {cat}
          </text>
        );
      })}
    </svg>
  );
}

/** Round up to a visually nice axis maximum. */
function niceNumber(val: number): number {
  if (val <= 0) return 10;
  const exp = Math.pow(10, Math.floor(Math.log10(val)));
  const frac = val / exp;
  let nice: number;
  if (frac <= 1) nice = 1;
  else if (frac <= 2) nice = 2;
  else if (frac <= 5) nice = 5;
  else nice = 10;
  return nice * exp;
}

// ─── Pie chart ────────────────────────────────────────────────────────────────

function PieChartSVG({ chart }: { chart: ParsedChart }) {
  const seriesData = chart.series[0]?.data ?? [];
  if (!seriesData.length) return null;

  const total = seriesData.reduce((s, d) => s + d.value, 0);
  if (total === 0) return null;

  const cx = 90;
  const cy = 90;
  const r = 75;

  let cumAngle = -Math.PI / 2; // start at 12 o'clock

  const slices = seriesData.map((d, i) => {
    const angle = (d.value / total) * 2 * Math.PI;
    const startA = cumAngle;
    cumAngle += angle;
    const endA = cumAngle;

    const x1 = cx + r * Math.cos(startA);
    const y1 = cy + r * Math.sin(startA);
    const x2 = cx + r * Math.cos(endA);
    const y2 = cy + r * Math.sin(endA);
    const largeArc = angle > Math.PI ? 1 : 0;

    const midA = startA + angle / 2;
    const labelR = r * 0.65;
    const lx = cx + labelR * Math.cos(midA);
    const ly = cy + labelR * Math.sin(midA);

    const pct = Math.round((d.value / total) * 100);
    const color = d.color ?? FALLBACK_COLORS[i % FALLBACK_COLORS.length];

    return { x1, y1, x2, y2, largeArc, lx, ly, pct, color, angle };
  });

  return (
    <svg
      viewBox="0 0 220 180"
      preserveAspectRatio="xMidYMid meet"
      style={{ width: "100%", height: "100%" }}
      aria-hidden="true"
    >
      {slices.map((s, i) => (
        <path
          key={i}
          d={`M ${cx} ${cy} L ${s.x1} ${s.y1} A ${r} ${r} 0 ${s.largeArc} 1 ${s.x2} ${s.y2} Z`}
          fill={s.color}
          stroke="#fff"
          strokeWidth="1"
        />
      ))}
      {slices.map((s, i) =>
        s.angle > 0.25 ? (
          <text
            key={i}
            x={s.lx}
            y={s.ly}
            textAnchor="middle"
            dominantBaseline="central"
            fontSize="8"
            fill="#fff"
            fontWeight="bold"
          >
            {s.pct}%
          </text>
        ) : null,
      )}
    </svg>
  );
}
