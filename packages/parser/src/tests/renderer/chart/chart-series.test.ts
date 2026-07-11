import { beforeAll, describe, expect, it } from "vitest";

import { parseExplosion, parseSeries } from "../../../renderer/chart/series";
import type { RenderContext } from "../../../renderer/context";
import { parseChartFragment, createChartTestContext, seriesXml } from "../../fixtures/chart-pptx";

let ctx: RenderContext;

beforeAll(async () => {
  ctx = await createChartTestContext();
});

describe("parseSeries", () => {
  it("extracts names, categories, and values from strRef/numRef caches", () => {
    const barChart = parseChartFragment(
      `<c:barChart>${seriesXml(0, "Revenue", ["Q1", "Q2", "Q3"], [10, 20, 30])}</c:barChart>`,
    ).child("barChart");
    const series = parseSeries(barChart, ctx);
    expect(series).toHaveLength(1);
    expect(series[0].name).toBe("Revenue");
    expect(series[0].categories).toEqual(["Q1", "Q2", "Q3"]);
    expect(series[0].values).toEqual([10, 20, 30]);
    expect(series[0].blankIndices?.size).toBe(0);
  });

  it("sorts series by c:order, not document order", () => {
    const barChart = parseChartFragment(`<c:barChart>
<c:ser><c:idx val="1"/><c:order val="1"/><c:tx><c:v>Second</c:v></c:tx></c:ser>
<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>First</c:v></c:tx></c:ser>
</c:barChart>`).child("barChart");
    const series = parseSeries(barChart, ctx);
    expect(series.map((s) => s.name)).toEqual(["First", "Second"]);
    expect(series.map((s) => s.order)).toEqual([0, 1]);
  });

  it("records blank value indices for gap points", () => {
    const barChart = parseChartFragment(
      `<c:barChart>${seriesXml(0, "Gappy", ["A", "B", "C"], [1, null, 3])}</c:barChart>`,
    ).child("barChart");
    const [s] = parseSeries(barChart, ctx);
    expect(s.values).toEqual([1, 0, 3]);
    expect([...(s.blankIndices ?? [])]).toEqual([1]);
  });

  it("reads the value format code from the numCache", () => {
    const barChart = parseChartFragment(`<c:barChart><c:ser>
<c:idx val="0"/><c:order val="0"/>
<c:val><c:numRef><c:numCache><c:formatCode>0.0%</c:formatCode><c:ptCount val="1"/><c:pt idx="0"><c:v>0.5</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser></c:barChart>`).child("barChart");
    expect(parseSeries(barChart, ctx)[0].formatCode).toBe("0.0%");
  });

  it("resolves an explicit series solid fill to hex", () => {
    const barChart = parseChartFragment(`<c:barChart><c:ser>
<c:idx val="0"/><c:order val="0"/>
<c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr>
</c:ser></c:barChart>`).child("barChart");
    expect(parseSeries(barChart, ctx)[0].colorHex).toBe("#FF0000");
  });

  it("parses marker symbol/size, smooth flag, and line width", () => {
    const lineChart = parseChartFragment(`<c:lineChart><c:ser>
<c:idx val="0"/><c:order val="0"/>
<c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln></c:spPr>
<c:marker><c:symbol val="diamond"/><c:size val="7"/></c:marker>
<c:smooth val="1"/>
</c:ser></c:lineChart>`).child("lineChart");
    const [s] = parseSeries(lineChart, ctx);
    expect(s.markerSymbol).toBe("diamond");
    // 7pt × 96/72 = 9.333px
    expect(s.markerSize).toBeCloseTo(9.333, 3);
    expect(s.smooth).toBe(true);
    // 28575 EMU / 12700 = 2.25px
    expect(s.lineWidth).toBe(2.25);
    expect(s.colorHex).toBe("#00FF00");
  });

  it("uses yVal for values and xVal for x coordinates on scatter series", () => {
    const scatterChart = parseChartFragment(`<c:scatterChart><c:ser>
<c:idx val="0"/><c:order val="0"/>
<c:xVal><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numCache></c:numRef></c:xVal>
<c:yVal><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>40</c:v></c:pt></c:numCache></c:numRef></c:yVal>
</c:ser></c:scatterChart>`).child("scatterChart");
    const [s] = parseSeries(scatterChart, ctx);
    expect(s.values).toEqual([10, 20, 40]);
    expect(s.xValues).toEqual([1, 2, 4]);
    // numeric xVal doubles as categories when no c:cat exists
    expect(s.categories).toEqual(["1", "2", "4"]);
  });

  it("parses bubble sizes and per-point styles", () => {
    const bubbleChart = parseChartFragment(`<c:bubbleChart><c:ser>
<c:idx val="0"/><c:order val="0"/>
<c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></c:spPr></c:dPt>
<c:xVal><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:xVal>
<c:yVal><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>6</c:v></c:pt></c:numCache></c:numRef></c:yVal>
<c:bubbleSize><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>100</c:v></c:pt><c:pt idx="1"><c:v>400</c:v></c:pt></c:numCache></c:numRef></c:bubbleSize>
</c:ser></c:bubbleChart>`).child("bubbleChart");
    const [s] = parseSeries(bubbleChart, ctx);
    expect(s.bubbleSizes).toEqual([100, 400]);
    expect(s.dataPointStyles?.[1]?.color).toBe("#0000FF");
    expect(s.dataPointColors?.[1]).toBe("#0000FF");
  });

  it("parses invertIfNegative explicitly set to off", () => {
    const barChart = parseChartFragment(`<c:barChart><c:ser>
<c:idx val="0"/><c:order val="0"/>
<c:invertIfNegative val="0"/>
</c:ser></c:barChart>`).child("barChart");
    expect(parseSeries(barChart, ctx)[0].invertIfNegative).toBe(false);
  });
});

describe("parseExplosion", () => {
  it("returns undefined when nothing is exploded", () => {
    const ser = parseChartFragment(`<c:ser><c:idx val="0"/></c:ser>`).child("ser");
    expect(parseExplosion(ser, 3)).toBeUndefined();
  });

  it("applies a series-level explosion to every point", () => {
    const ser = parseChartFragment(`<c:ser><c:explosion val="10"/></c:ser>`).child("ser");
    expect(parseExplosion(ser, 3)).toEqual([10, 10, 10]);
  });

  it("applies per-point dPt explosions on top of the series default", () => {
    const ser = parseChartFragment(`<c:ser>
<c:dPt><c:idx val="2"/><c:explosion val="25"/></c:dPt>
</c:ser>`).child("ser");
    expect(parseExplosion(ser, 3)).toEqual([0, 0, 25]);
  });
});
