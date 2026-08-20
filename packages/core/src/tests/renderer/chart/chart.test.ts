import type * as echarts from "echarts";
import { beforeAll, describe, expect, it } from "vitest";

import { buildPresentation } from "../../../model/presentation";
import { parseXml } from "../../../ooxml/xml";
import { readPptx } from "../../../ooxml/zip";
import { parseChartXml } from "../../../renderer/chart";
import { parseDataLabels, parsePointDataLabelOverrides } from "../../../renderer/chart/data-label";
import { buildDataTableElement, parseDataTable } from "../../../renderer/chart/data-table";
import type { SeriesData } from "../../../renderer/chart/type";
import type { RenderContext } from "../../../renderer/context";
import { renderSlide } from "../../../renderer/slide";
import {
  BAR_CHART_XML,
  axesXml,
  buildPptxWithChart,
  chartSpaceXml,
  createChartTestContext,
  parseChartFragment,
  seriesXml,
} from "../../fixtures/chart-pptx";

const OFFICE_ACCENTS = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"];

let ctx: RenderContext;

beforeAll(async () => {
  ctx = await createChartTestContext();
});

function parseOption(chartXmlString: string) {
  return parseChartXml(parseXml(chartXmlString), ctx);
}

type AnyRecord = Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any

describe("parseChartXml: bar chart", () => {
  let option: AnyRecord;

  beforeAll(() => {
    option = parseOption(BAR_CHART_XML) as AnyRecord;
    option = { ...option.option } as AnyRecord;
  });

  it("builds one ECharts bar series per c:ser", () => {
    const series = option.series as echarts.BarSeriesOption[];
    expect(series).toHaveLength(2);
    expect(series.map((s) => s.type)).toEqual(["bar", "bar"]);
    expect(series.map((s) => s.name)).toEqual(["Series A", "Series B"]);
    expect(series[0].data).toEqual([10, 20, 30]);
    expect(series[1].data).toEqual([5, 15, 25]);
  });

  it("derives bar gaps from the OOXML gapWidth default", () => {
    const series = option.series as AnyRecord[];
    expect(series[0].barGap).toBe("0%");
    // gapWidth 150 with 2 clustered series → 150/(200+150) ≈ 43%
    expect(series[0].barCategoryGap).toBe("43%");
  });

  it("uses the categories on the category axis", () => {
    expect(option.xAxis).toMatchObject({ type: "category", data: ["Q1", "Q2", "Q3"] });
  });

  it("applies a PowerPoint-like nice range to the value axis", () => {
    // data max 30 → interval 5, headroom to 35, floor at 0
    expect(option.yAxis).toMatchObject({ type: "value", min: 0, max: 35, interval: 5 });
  });

  it("extracts the rich text title with theme font and default color", () => {
    expect(option.title).toMatchObject({ text: "Sales 2024", left: "center" });
    expect(option.title.textStyle).toMatchObject({
      fontSize: 12,
      color: "#000000",
      fontWeight: "bold",
    });
    expect(option.title.textStyle.fontFamily).toContain("Calibri");
  });

  it("places the legend at the bottom with rect icons and series names", () => {
    expect(option.legend).toMatchObject({
      bottom: "5%",
      orient: "horizontal",
      icon: "rect",
      data: ["Series A", "Series B"],
    });
  });

  it("adopts the theme accent palette", () => {
    expect(option.color).toEqual(OFFICE_ACCENTS);
  });

  it("reserves grid space for title and bottom legend", () => {
    expect(option.grid).toMatchObject({ containLabel: true, top: 68, bottom: 35 });
  });
});

describe("parseChartXml: bar chart variants", () => {
  it("formats value axis labels using the axis numFmt", () => {
    const xml = chartSpaceXml(`<c:plotArea>
<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
${seriesXml(0, "S", ["A"], [0.5])}
<c:axId val="111111111"/><c:axId val="222222222"/>
</c:barChart>
${axesXml("111111111", "222222222", `<c:numFmt formatCode="0%" sourceLinked="0"/>`)}
</c:plotArea>`);
    const { option } = parseOption(xml) as AnyRecord;
    const formatter = option.yAxis.axisLabel.formatter as (v: number) => string;
    expect(formatter(0.5)).toBe("50%");
  });

  it("fills the plot area via grid without the default border for a noFill line", () => {
    const xml = chartSpaceXml(`<c:plotArea>
<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
${seriesXml(0, "S", ["A"], [1])}
</c:barChart>
<c:spPr><a:solidFill><a:srgbClr val="1C1C1C"/></a:solidFill><a:ln w="0"><a:noFill/></a:ln></c:spPr>
</c:plotArea>`);
    const { option } = parseOption(xml) as AnyRecord;
    expect(option.grid.show).toBe(true);
    expect(option.grid.backgroundColor).toBe("#1C1C1C");
    expect(option.grid.borderWidth).toBe(0);
  });

  it("draws the plot area border when the plot area declares a solid line", () => {
    const xml = chartSpaceXml(`<c:plotArea>
<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
${seriesXml(0, "S", ["A"], [1])}
</c:barChart>
<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr>
</c:plotArea>`);
    const { option } = parseOption(xml) as AnyRecord;
    expect(option.grid.borderColor).toBe("#FF0000");
    // 12700 EMU = 1pt = 1.333px
    expect(option.grid.borderWidth).toBeCloseTo(1.333, 3);
  });

  it("stacks series and widens bars for stacked grouping", () => {
    const xml = chartSpaceXml(`<c:plotArea>
<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/>
${seriesXml(0, "A", ["X"], [1])}
${seriesXml(1, "B", ["X"], [2])}
</c:barChart>
</c:plotArea>`);
    const { option } = parseOption(xml) as AnyRecord;
    expect(option.series[0].stack).toBe("total");
    // stacked bars share one slot → 150/(100+150) = 60%
    expect(option.series[0].barCategoryGap).toBe("60%");
  });

  it("normalizes values and forces a percent axis for percentStacked grouping", () => {
    const xml = chartSpaceXml(`<c:plotArea>
<c:barChart><c:barDir val="col"/><c:grouping val="percentStacked"/>
${seriesXml(0, "A", ["X", "Y"], [10, 20])}
${seriesXml(1, "B", ["X", "Y"], [30, 20])}
</c:barChart>
</c:plotArea>`);
    const { option } = parseOption(xml) as AnyRecord;
    expect(option.series[0].data).toEqual([0.25, 0.5]);
    expect(option.series[1].data).toEqual([0.75, 0.5]);
    expect(option.yAxis).toMatchObject({ min: 0, max: 1, interval: 0.1 });
    expect(option.yAxis.axisLabel.formatter(0.3)).toBe("30%");
  });

  it("renders negative bars inverted (white fill, black border) by default", () => {
    const xml = chartSpaceXml(`<c:plotArea>
<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
${seriesXml(0, "Mixed", ["A", "B"], [10, -5])}
</c:barChart>
</c:plotArea>`);
    const { option } = parseOption(xml) as AnyRecord;
    const negative = option.series[0].data[1];
    expect(negative).toMatchObject({
      value: -5,
      itemStyle: { color: "#FFFFFF", borderColor: "#000000", borderWidth: 1 },
    });
  });

  it("varies point colors for a single un-colored series", () => {
    const xml = chartSpaceXml(`<c:plotArea>
<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
${seriesXml(0, "Solo", ["A", "B"], [1, 2])}
</c:barChart>
</c:plotArea>`);
    const { option } = parseOption(xml) as AnyRecord;
    // darkened accent palette cycles per point
    expect(option.series[0].data[0]).toMatchObject({ itemStyle: { color: "#3c64ac" } });
  });

  it("swaps axes for horizontal bar charts", () => {
    const xml = chartSpaceXml(`<c:plotArea>
<c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/>
${seriesXml(0, "S", ["A", "B"], [1, 2])}
</c:barChart>
</c:plotArea>`);
    const { option } = parseOption(xml) as AnyRecord;
    expect(option.yAxis).toMatchObject({ type: "category", data: ["A", "B"] });
    expect(option.xAxis.type).toBe("value");
  });

  it("clusters horizontal bars upward, first series at the bottom, legend to match", () => {
    const xml = chartSpaceXml(`<c:plotArea>
<c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/>
${seriesXml(0, "First", ["A", "B"], [1, 2])}
${seriesXml(1, "Second", ["A", "B"], [3, 4])}
</c:barChart>
</c:plotArea>
<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>`);
    const { option } = parseOption(xml) as AnyRecord;
    // ECharts draws series[0] at the top of each cluster; PowerPoint puts it at
    // the bottom, so the drawn order is the reverse of the OOXML order.
    expect(option.series.map((s: AnyRecord) => s.name)).toEqual(["Second", "First"]);
    expect(option.legend.data).toEqual(["Second", "First"]);
  });

  it("keeps series order for vertical and stacked bars", () => {
    const upright = chartSpaceXml(`<c:plotArea>
<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
${seriesXml(0, "First", ["A"], [1])}
${seriesXml(1, "Second", ["A"], [2])}
</c:barChart>
</c:plotArea>
<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>`);
    // Stacked bars stack in series order in both, so reversing would break them.
    const stacked = chartSpaceXml(`<c:plotArea>
<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/>
${seriesXml(0, "First", ["A"], [1])}
${seriesXml(1, "Second", ["A"], [2])}
</c:barChart>
</c:plotArea>
<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>`);
    for (const xml of [upright, stacked]) {
      const { option } = parseOption(xml) as AnyRecord;
      expect(option.series.map((s: AnyRecord) => s.name)).toEqual(["First", "Second"]);
      expect(option.legend.data).toEqual(["First", "Second"]);
    }
  });

  it("pins each series' palette colour so reordering cannot reassign it", () => {
    const xml = chartSpaceXml(`<c:plotArea>
<c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/>
${seriesXml(0, "First", ["A"], [1])}
${seriesXml(1, "Second", ["A"], [2])}
</c:barChart>
</c:plotArea>`);
    const { option } = parseOption(xml) as AnyRecord;
    const colours = Object.fromEntries(
      option.series.map((s: AnyRecord) => [s.name, s.itemStyle?.color]),
    );
    expect(colours).toEqual({ First: OFFICE_ACCENTS[0], Second: OFFICE_ACCENTS[1] });
  });
});

describe("parseChartXml: titles", () => {
  it("synthesizes the Office auto-title from a single series name", () => {
    const xml = chartSpaceXml(`<c:autoTitleDeleted val="0"/>
<c:plotArea>
<c:barChart>${seriesXml(0, "Only Series", ["A"], [1])}</c:barChart>
</c:plotArea>`);
    const { option } = parseOption(xml) as AnyRecord;
    expect(option.title.text).toBe("Only Series");
  });

  it("suppresses the title when autoTitleDeleted is set", () => {
    const xml =
      chartSpaceXml(`<c:title><c:tx><c:rich><a:p><a:r><a:t>Zombie</a:t></a:r></a:p></c:rich></c:tx></c:title>
<c:autoTitleDeleted val="1"/>
<c:plotArea>
<c:barChart>${seriesXml(0, "S", ["A"], [1])}</c:barChart>
</c:plotArea>`);
    const { option } = parseOption(xml) as AnyRecord;
    expect(option.title).toBeUndefined();
  });

  it("positions manually laid out titles", () => {
    const xml = chartSpaceXml(`<c:title>
<c:tx><c:rich><a:p><a:r><a:t>Placed</a:t></a:r></a:p></c:rich></c:tx>
<c:layout><c:manualLayout><c:x val="0.1"/><c:y val="0.05"/></c:manualLayout></c:layout>
</c:title>
<c:autoTitleDeleted val="0"/>
<c:plotArea>
<c:barChart>${seriesXml(0, "S", ["A"], [1])}</c:barChart>
</c:plotArea>`);
    const { option } = parseOption(xml) as AnyRecord;
    expect(option.title).toMatchObject({ text: "Placed", left: "10%", top: "5%" });
  });
});

describe("parseChartXml: line chart", () => {
  it("builds smooth marker-less line series", () => {
    const xml = chartSpaceXml(`<c:plotArea>
<c:lineChart><c:grouping val="standard"/>
<c:ser>
<c:idx val="0"/><c:order val="0"/>
<c:tx><c:v>Trend</c:v></c:tx>
<c:marker><c:symbol val="none"/></c:marker>
<c:cat><c:strRef><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>J</c:v></c:pt><c:pt idx="1"><c:v>F</c:v></c:pt><c:pt idx="2"><c:v>M</c:v></c:pt></c:strCache></c:strRef></c:cat>
<c:val><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val>
<c:smooth val="1"/>
</c:ser>
</c:lineChart>
</c:plotArea>`);
    const { option } = parseOption(xml) as AnyRecord;
    const line = option.series[0];
    expect(line.type).toBe("line");
    expect(line.smooth).toBe(true);
    expect(line.showSymbol).toBe(false);
    expect(line.data).toEqual([1, 3, 2]);
    expect(line.lineStyle.width).toBe(3);
  });

  it("maps blank points according to dispBlanksAs", () => {
    const seriesFragment = seriesXml(0, "Gappy", ["A", "B", "C"], [1, null, 3]);
    const gap = parseOption(
      chartSpaceXml(`<c:plotArea><c:lineChart>${seriesFragment}</c:lineChart></c:plotArea>`),
    ) as AnyRecord;
    expect(gap.option.series[0].data).toEqual([1, null, 3]);
    expect(gap.option.series[0].connectNulls).toBe(false);

    const zero = parseOption(
      chartSpaceXml(
        `<c:plotArea><c:lineChart>${seriesFragment}</c:lineChart></c:plotArea><c:dispBlanksAs val="zero"/>`,
      ),
    ) as AnyRecord;
    expect(zero.option.series[0].data).toEqual([1, 0, 3]);

    const span = parseOption(
      chartSpaceXml(
        `<c:plotArea><c:lineChart>${seriesFragment}</c:lineChart></c:plotArea><c:dispBlanksAs val="span"/>`,
      ),
    ) as AnyRecord;
    expect(span.option.series[0].data).toEqual([1, null, 3]);
    expect(span.option.series[0].connectNulls).toBe(true);
  });

  it("uses line path icons in the legend", () => {
    const xml = chartSpaceXml(`<c:plotArea>
<c:lineChart>${seriesXml(0, "Trend", ["A"], [1])}</c:lineChart>
</c:plotArea>
<c:legend><c:legendPos val="r"/></c:legend>`);
    const { option } = parseOption(xml) as AnyRecord;
    const entry = option.legend.data[0];
    expect(entry.name).toBe("Trend");
    expect(entry.icon).toMatch(/^path:\/\//);
  });
});

describe("parseChartXml: pie and doughnut", () => {
  const PIE_XML = chartSpaceXml(`<c:plotArea>
<c:pieChart>
<c:varyColors val="1"/>
${seriesXml(0, "Share", ["Alpha", "Beta", "Gamma"], [50, 30, 20], `<c:dLbls><c:dLblPos val="outEnd"/><c:showVal val="1"/><c:showPercent val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showLeaderLines val="0"/></c:dLbls>`)}
</c:pieChart>
</c:plotArea>
<c:legend><c:legendPos val="r"/></c:legend>`);

  it("builds pie data from categories and values", () => {
    const { option } = parseOption(PIE_XML) as AnyRecord;
    const pie = option.series[0];
    expect(pie.type).toBe("pie");
    expect(pie.data).toEqual([
      { name: "Alpha", value: 50 },
      { name: "Beta", value: 30 },
      { name: "Gamma", value: 20 },
    ]);
    expect(option.legend.data).toEqual(["Alpha", "Beta", "Gamma"]);
  });

  it("shifts the pie left of a right-hand legend", () => {
    const { option } = parseOption(PIE_XML) as AnyRecord;
    expect(option.series[0].center).toEqual(["38%", "55%"]);
    expect(option.series[0].radius).toBe("82%");
  });

  it("builds an outside data label showing value and percent", () => {
    const { option } = parseOption(PIE_XML) as AnyRecord;
    const label = option.series[0].label;
    expect(label.show).toBe(true);
    expect(label.position).toBe("outside");
    expect(label.formatter({ value: 50, name: "Alpha", percent: 50 })).toBe("50 50%");
    expect(option.series[0].labelLine.show).toBe(false);
  });

  it("computes doughnut ring radii from holeSize", () => {
    const xml = chartSpaceXml(`<c:plotArea>
<c:doughnutChart>
${seriesXml(0, "Rings", ["A", "B"], [1, 2])}
<c:holeSize val="40"/>
</c:doughnutChart>
</c:plotArea>`);
    const { option } = parseOption(xml) as AnyRecord;
    expect(option.series[0].type).toBe("pie");
    // outer 82%, inner = round(82 × 40%) = 33%
    expect(option.series[0].radius).toEqual(["33%", "82%"]);
    expect(option.series[0].center).toEqual(["50%", "55%"]);
  });

  it("marks exploded slices as selected", () => {
    const xml = chartSpaceXml(`<c:plotArea>
<c:pieChart>
<c:ser>
<c:idx val="0"/><c:order val="0"/>
<c:tx><c:v>S</c:v></c:tx>
<c:dPt><c:idx val="1"/><c:explosion val="20"/></c:dPt>
<c:cat><c:strRef><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:strRef></c:cat>
<c:val><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser>
</c:pieChart>
</c:plotArea>`);
    const { option } = parseOption(xml) as AnyRecord;
    const exploded = option.series[0].data[1];
    expect(exploded.selected).toBe(true);
    expect(exploded.selectedOffset).toBe(10);
    expect(option.series[0].selectedMode).toBe("multiple");
  });
});

describe("parseChartXml: scatter chart", () => {
  const scatterSeries = `<c:ser>
<c:idx val="0"/><c:order val="0"/>
<c:tx><c:v>Points</c:v></c:tx>
<c:xVal><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numCache></c:numRef></c:xVal>
<c:yVal><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>40</c:v></c:pt></c:numCache></c:numRef></c:yVal>
</c:ser>`;

  it("renders marker-only scatter styles as scatter series with x/y pairs", () => {
    const xml = chartSpaceXml(`<c:plotArea>
<c:scatterChart><c:scatterStyle val="none"/>${scatterSeries}</c:scatterChart>
</c:plotArea>`);
    const { option } = parseOption(xml) as AnyRecord;
    expect(option.series[0].type).toBe("scatter");
    expect(option.series[0].data).toEqual([
      [1, 10],
      [2, 20],
      [4, 40],
    ]);
    expect(option.xAxis.type).toBe("value");
    expect(option.yAxis.type).toBe("value");
  });

  it("renders lineMarker scatter styles as line series with default symbols", () => {
    const xml = chartSpaceXml(`<c:plotArea>
<c:scatterChart><c:scatterStyle val="lineMarker"/>${scatterSeries}</c:scatterChart>
</c:plotArea>`);
    const { option } = parseOption(xml) as AnyRecord;
    expect(option.series[0].type).toBe("line");
    expect(option.series[0].symbol).toBe("diamond");
    expect(option.series[0].showSymbol).toBe(true);
  });
});

describe("parseChartXml: combo charts", () => {
  it("merges a secondary line chart onto a bar chart with a second value axis", () => {
    const xml = chartSpaceXml(`<c:plotArea>
<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
${seriesXml(0, "Bars A", ["X", "Y"], [10, 20])}
${seriesXml(1, "Bars B", ["X", "Y"], [5, 15])}
<c:axId val="1"/><c:axId val="2"/>
</c:barChart>
<c:lineChart><c:grouping val="standard"/>
${seriesXml(2, "Line", ["X", "Y"], [0.1, 0.9])}
<c:axId val="1"/><c:axId val="3"/>
</c:lineChart>
<c:catAx><c:axId val="1"/><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
<c:valAx><c:axId val="2"/><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
<c:valAx><c:axId val="3"/><c:delete val="0"/><c:axPos val="r"/><c:crossAx val="1"/></c:valAx>
</c:plotArea>
<c:legend><c:legendPos val="b"/></c:legend>`);
    const { option } = parseOption(xml) as AnyRecord;

    expect(option.series.map((s: AnyRecord) => s.type)).toEqual(["bar", "bar", "line"]);
    expect(Array.isArray(option.yAxis)).toBe(true);
    expect(option.yAxis).toHaveLength(2);
    expect(option.series[2].yAxisIndex).toBe(1);
    const legendNames = option.legend.data.map((d: AnyRecord | string) =>
      typeof d === "string" ? d : d.name,
    );
    expect(legendNames).toEqual(["Bars A", "Bars B", "Line"]);
  });
});

describe("parseChartXml: fallbacks", () => {
  it("reports an unsupported chart when the plotArea has no known chart type", () => {
    const { option } = parseOption(
      chartSpaceXml("<c:plotArea><c:layout/></c:plotArea>"),
    ) as AnyRecord;
    expect(option.title.text).toBe("Unsupported chart type");
  });

  it("reports an unsupported chart when the plotArea is missing", () => {
    const { option } = parseOption(chartSpaceXml("")) as AnyRecord;
    expect(option.title.text).toBe("Unsupported chart");
  });
});

describe("parseChartXml: determinism", () => {
  // Thumbnails, screenshot tests and video renders all capture a single frame.
  it("disables the ECharts entrance animation", () => {
    const { option } = parseOption(BAR_CHART_XML) as AnyRecord;
    expect(option.animation).toBe(false);
  });
});

describe("data labels", () => {
  it("parses shared dLbls config", () => {
    const node = parseChartFragment(`<c:ser><c:dLbls>
<c:dLblPos val="ctr"/>
<c:showVal val="1"/><c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showLeaderLines val="1"/>
<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1100" b="1"><a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
</c:dLbls></c:ser>`).child("ser");
    expect(parseDataLabels(node, ctx)).toMatchObject({
      showVal: true,
      showCatName: true,
      showSerName: false,
      showPercent: false,
      showLeaderLines: true,
      position: "ctr",
      color: "#FF00FF",
      fontSize: 11,
      bold: true,
    });
  });

  it("returns undefined when nothing is shown", () => {
    const node = parseChartFragment(
      `<c:ser><c:dLbls><c:showVal val="0"/><c:showCatName val="0"/></c:dLbls></c:ser>`,
    ).child("ser");
    expect(parseDataLabels(node, ctx)).toBeUndefined();
  });

  it("parses per-point overrides including deletions", () => {
    const dLbls = parseChartFragment(`<c:dLbls>
<c:dLbl><c:idx val="0"/><c:delete val="1"/></c:dLbl>
<c:dLbl><c:idx val="2"/><c:showVal val="1"/><c:dLblPos val="outEnd"/></c:dLbl>
<c:showVal val="1"/>
</c:dLbls>`).child("dLbls");
    const overrides = parsePointDataLabelOverrides(dLbls, ctx);
    expect(overrides.get(0)).toMatchObject({ deleted: true, showVal: false });
    expect(overrides.get(2)).toMatchObject({ showVal: true, position: "outEnd" });
    expect(overrides.has(1)).toBe(false);
  });
});

describe("data table", () => {
  it("exposes dTable info from parseChartXml", () => {
    const xml = chartSpaceXml(`<c:plotArea>
<c:barChart><c:barDir val="col"/>
${seriesXml(0, "S", ["A", "B"], [1, 2])}
</c:barChart>
<c:dTable><c:showKeys val="1"/></c:dTable>
</c:plotArea>`);
    const result = parseOption(xml);
    expect(result.dataTable).toBeDefined();
    expect(result.dataTable!.showKeys).toBe(true);
    expect(result.dataTable!.seriesArr[0].name).toBe("S");
  });

  it("parseDataTable returns undefined without c:dTable", () => {
    const plotArea = parseChartFragment("<c:plotArea/>").child("plotArea");
    expect(parseDataTable(plotArea)).toBeUndefined();
  });

  it("buildDataTableElement lays out categories, series rows, and blanks", () => {
    const seriesArr: SeriesData[] = [
      {
        name: "Revenue",
        order: 0,
        categories: ["Q1", "Q2", "Q3"],
        values: [1000, 0, 3000],
        formatCode: "#,##0",
        blankIndices: new Set([1]),
      },
    ];
    const table = buildDataTableElement({ seriesArr, showKeys: true }, ["#FF0000"]);

    const headers = [...table.querySelectorAll("thead th")].map((th) => th.textContent);
    expect(headers).toEqual(["", "Q1", "Q2", "Q3"]);

    const cells = [...table.querySelectorAll("tbody td")].map((td) => td.textContent);
    expect(cells).toEqual(["Revenue", "1,000", "", "3,000"]);

    const key = table.querySelector("tbody td span") as HTMLElement;
    expect(key.style.backgroundColor).toBeTruthy();
  });
});

describe("renderChart via renderSlide (happy-dom)", () => {
  // NOTE: full ECharts initialization is intentionally NOT exercised here.
  // The chart renderer is hard-wired to the echarts CanvasRenderer, and
  // happy-dom's <canvas> has no 2D context (getContext returns null), which
  // makes zrender's async paint loop throw unhandled errors. renderChart
  // defers init until the container is connected and sized, so rendering a
  // detached slide deterministically produces the DOM scaffolding only.
  it("positions the chart wrapper and builds the custom legend overlay", async () => {
    const buffer = await buildPptxWithChart(BAR_CHART_XML);
    const presentation = buildPresentation(await readPptx(buffer));
    const handle = renderSlide(presentation, presentation.slides[0]);
    await handle.ready;

    const legendOverlay = handle.element.querySelector(
      ".pptx-chart-custom-legend",
    ) as HTMLElement | null;
    expect(legendOverlay).not.toBeNull();
    const legendLabels = [...legendOverlay!.querySelectorAll("span")].map((s) => s.textContent);
    expect(legendLabels).toEqual(["Series A", "Series B"]);

    const wrapper = legendOverlay!.parentElement as HTMLElement;
    expect(wrapper.style.left).toBe("96px");
    expect(wrapper.style.top).toBe("96px");
    expect(wrapper.style.width).toBe("640px");
    expect(wrapper.style.height).toBe("480px");

    // No ECharts instance is created in the detached happy-dom container.
    expect(handle.element.querySelector("canvas")).toBeNull();
    handle.dispose();
  });

  it("renders a placeholder when the chart part is missing from the package", async () => {
    const buffer = await buildPptxWithChart(BAR_CHART_XML, { omitChartPart: true });
    const presentation = buildPresentation(await readPptx(buffer));
    const handle = renderSlide(presentation, presentation.slides[0]);
    await handle.ready;

    expect(handle.element.textContent).toContain("Chart not found");
    handle.dispose();
  });
});
