import { beforeAll, describe, expect, it } from "vitest";

import {
  applyAxisInfo,
  getAxisTitleSpacePx,
  getChartAxisIds,
  parseAxes,
  parseScatterAxes,
} from "../../../renderer/chart/axis";
import { parseOoxmlBoolElement } from "../../../renderer/chart/boolean";
import type { RenderContext } from "../../../renderer/context";
import { parseChartFragment, createChartTestContext } from "../../fixtures/chart-pptx";

let ctx: RenderContext;

beforeAll(async () => {
  ctx = await createChartTestContext();
});

describe("parseOoxmlBoolElement", () => {
  it("is false for a missing element", () => {
    expect(parseOoxmlBoolElement(parseChartFragment("").child("delete"))).toBe(false);
  });

  it("defaults to true when the element exists without val", () => {
    expect(parseOoxmlBoolElement(parseChartFragment("<c:delete/>").child("delete"))).toBe(true);
  });

  it("honors explicit val attributes in all OOXML spellings", () => {
    for (const [val, expected] of [
      ["0", false],
      ["1", true],
      ["false", false],
      ["true", true],
      ["off", false],
      ["on", true],
    ] as const) {
      const node = parseChartFragment(`<c:delete val="${val}"/>`).child("delete");
      expect(parseOoxmlBoolElement(node)).toBe(expected);
    }
  });
});

describe("parseAxes", () => {
  it("extracts scaling min/max, format code, gridlines, and tick label position", () => {
    const plotArea = parseChartFragment(`<c:plotArea>
<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:tickLblPos val="low"/></c:catAx>
<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="maxMin"/><c:min val="5"/><c:max val="50"/></c:scaling><c:delete val="0"/><c:numFmt formatCode="0.0%"/><c:majorGridlines/><c:majorTickMark val="none"/></c:valAx>
</c:plotArea>`).child("plotArea");
    const { valueAxis, categoryAxis } = parseAxes(plotArea, ctx);

    expect(valueAxis.min).toBe(5);
    expect(valueAxis.max).toBe(50);
    expect(valueAxis.orientation).toBe("maxMin");
    expect(valueAxis.numFmt).toBe("0.0%");
    expect(valueAxis.hasMajorGridlines).toBe(true);
    expect(valueAxis.majorTickMark).toBe("none");
    expect(valueAxis.deleted).toBe(false);

    expect(categoryAxis.tickLblPos).toBe("low");
    expect(categoryAxis.orientation).toBe("minMax");
    expect(categoryAxis.hasMajorGridlines).toBe(false);
  });

  it("treats a General numFmt as no format code", () => {
    const plotArea = parseChartFragment(`<c:plotArea>
<c:valAx><c:axId val="2"/><c:numFmt formatCode="General"/></c:valAx>
</c:plotArea>`).child("plotArea");
    expect(parseAxes(plotArea, ctx).valueAxis.numFmt).toBeUndefined();
  });

  it("flags deleted axes", () => {
    const plotArea = parseChartFragment(`<c:plotArea>
<c:valAx><c:axId val="2"/><c:delete val="1"/></c:valAx>
</c:plotArea>`).child("plotArea");
    expect(parseAxes(plotArea, ctx).valueAxis.deleted).toBe(true);
  });

  it("marks axis lines hidden when spPr declares a noFill line", () => {
    const plotArea = parseChartFragment(`<c:plotArea>
<c:valAx><c:axId val="2"/><c:spPr><a:ln w="0"><a:noFill/></a:ln></c:spPr></c:valAx>
<c:catAx><c:axId val="1"/><c:spPr><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr></c:catAx>
</c:plotArea>`).child("plotArea");
    const { valueAxis, categoryAxis } = parseAxes(plotArea, ctx);
    expect(valueAxis.lineHidden).toBe(true);
    expect(valueAxis.lineColor).toBeUndefined();
    expect(categoryAxis.lineHidden).toBe(false);
    expect(categoryAxis.lineColor).toBe("#FF0000");
  });

  it("hides the axis line and ticks for a noFill axis line", () => {
    const plotArea = parseChartFragment(`<c:plotArea>
<c:valAx><c:axId val="2"/><c:spPr><a:ln w="0"><a:noFill/></a:ln></c:spPr></c:valAx>
</c:plotArea>`).child("plotArea");
    const { valueAxis } = parseAxes(plotArea, ctx);
    const axisDef: Record<string, unknown> = {};
    applyAxisInfo(axisDef, valueAxis, "value");
    expect(axisDef.axisLine).toMatchObject({ show: false });
    expect(axisDef.axisTick).toMatchObject({ show: false });
  });

  it("returns defaults when the axis node is absent", () => {
    const plotArea = parseChartFragment("<c:plotArea/>").child("plotArea");
    const { valueAxis } = parseAxes(plotArea, ctx);
    expect(valueAxis).toMatchObject({
      deleted: false,
      tickLblPos: "nextTo",
      hasMajorGridlines: false,
      orientation: "minMax",
    });
  });

  it("selects axes by axId when the chart type node declares them", () => {
    const root = parseChartFragment(`<c:plotArea>
<c:barChart><c:axId val="10"/><c:axId val="20"/></c:barChart>
<c:valAx><c:axId val="99"/><c:scaling><c:max val="1"/></c:scaling></c:valAx>
<c:valAx><c:axId val="20"/><c:scaling><c:max val="200"/></c:scaling></c:valAx>
<c:catAx><c:axId val="10"/><c:tickLblPos val="none"/></c:catAx>
</c:plotArea>`);
    const plotArea = root.child("plotArea");
    const chartTypeNode = plotArea.child("barChart");

    expect(getChartAxisIds(chartTypeNode)).toEqual(["10", "20"]);
    const { valueAxis, categoryAxis } = parseAxes(plotArea, ctx, chartTypeNode);
    expect(valueAxis.max).toBe(200);
    expect(categoryAxis.tickLblPos).toBe("none");
  });

  it("extracts axis titles with rotation", () => {
    const plotArea = parseChartFragment(`<c:plotArea>
<c:valAx><c:axId val="2"/>
<c:title><c:tx><c:rich><a:bodyPr rot="-5400000"/><a:p><a:r><a:t>Revenue</a:t></a:r></a:p></c:rich></c:tx></c:title>
</c:valAx>
</c:plotArea>`).child("plotArea");
    const { valueAxis } = parseAxes(plotArea, ctx);
    expect(valueAxis.title).toBe("Revenue");
    expect(valueAxis.titleRotation).toBe(-90);
  });
});

describe("parseScatterAxes", () => {
  it("assigns valAx nodes to x/y by axPos", () => {
    const plotArea = parseChartFragment(`<c:plotArea>
<c:valAx><c:axId val="1"/><c:axPos val="b"/><c:scaling><c:max val="10"/></c:scaling></c:valAx>
<c:valAx><c:axId val="2"/><c:axPos val="l"/><c:scaling><c:max val="99"/></c:scaling></c:valAx>
</c:plotArea>`).child("plotArea");
    const { xAxis, yAxis } = parseScatterAxes(plotArea, ctx);
    expect(xAxis.max).toBe(10);
    expect(yAxis.max).toBe(99);
  });

  it("treats a single valAx as the y axis", () => {
    const plotArea = parseChartFragment(`<c:plotArea>
<c:valAx><c:axId val="1"/><c:scaling><c:max val="42"/></c:scaling></c:valAx>
</c:plotArea>`).child("plotArea");
    const { xAxis, yAxis } = parseScatterAxes(plotArea, ctx);
    expect(yAxis.max).toBe(42);
    expect(xAxis.max).toBeUndefined();
  });
});

describe("applyAxisInfo", () => {
  it("hides everything for deleted value axes", () => {
    const def: Record<string, unknown> = { type: "value" };
    applyAxisInfo(
      def,
      { deleted: true, tickLblPos: "nextTo", hasMajorGridlines: false, orientation: "minMax" },
      "value",
    );
    expect(def.axisLabel).toMatchObject({ show: false });
    expect(def.axisLine).toEqual({ show: false });
    expect(def.axisTick).toEqual({ show: false });
    expect(def.splitLine).toEqual({ show: false });
  });

  it("inverts maxMin axes and applies min/max to value axes", () => {
    const def: Record<string, unknown> = { type: "value" };
    applyAxisInfo(
      def,
      {
        deleted: false,
        tickLblPos: "nextTo",
        hasMajorGridlines: true,
        orientation: "maxMin",
        min: 5,
        max: 50,
      },
      "value",
    );
    expect(def.inverse).toBe(true);
    expect(def.min).toBe(5);
    expect(def.max).toBe(50);
    expect(def.splitLine).toMatchObject({
      show: true,
      lineStyle: { color: "#898989", width: 1, type: "solid" },
    });
  });

  it("hides tick labels when tickLblPos is none", () => {
    const def: Record<string, unknown> = { type: "category" };
    applyAxisInfo(
      def,
      { deleted: false, tickLblPos: "none", hasMajorGridlines: false, orientation: "minMax" },
      "category",
    );
    expect(def.axisLabel).toMatchObject({ show: false });
    // axis line itself remains visible
    expect(def.axisLine).toMatchObject({ show: true });
  });

  it("installs a value-axis label formatter honoring the axis numFmt", () => {
    const def: Record<string, unknown> = { type: "value" };
    applyAxisInfo(
      def,
      {
        deleted: false,
        tickLblPos: "nextTo",
        hasMajorGridlines: false,
        orientation: "minMax",
        numFmt: "0%",
      },
      "value",
    );
    const label = def.axisLabel as { formatter: (v: number) => string };
    expect(label.formatter(0.5)).toBe("50%");
  });

  it("marks the axis line onZero for crosses=autoZero", () => {
    const def: Record<string, unknown> = { type: "category" };
    applyAxisInfo(
      def,
      {
        deleted: false,
        tickLblPos: "nextTo",
        hasMajorGridlines: false,
        orientation: "minMax",
        crosses: "autoZero",
      },
      "category",
    );
    expect(def.axisLine).toMatchObject({ onZero: true });
  });

  it("applies axis titles with style to the name fields", () => {
    const def: Record<string, unknown> = { type: "value" };
    applyAxisInfo(
      def,
      {
        deleted: false,
        tickLblPos: "nextTo",
        hasMajorGridlines: false,
        orientation: "minMax",
        title: "Revenue",
        titleStyle: { color: "#FF0000", fontSize: 12, bold: true },
        titleRotation: -90,
      },
      "value",
    );
    expect(def.name).toBe("Revenue");
    expect(def.nameLocation).toBe("middle");
    expect(def.nameGap).toBe(42);
    expect(def.nameRotate).toBe(-90);
    expect(def.nameTextStyle).toEqual({ color: "#FF0000", fontSize: 12, fontWeight: "bold" });
  });
});

describe("getAxisTitleSpacePx", () => {
  const axis = {
    deleted: false,
    tickLblPos: "nextTo",
    hasMajorGridlines: false,
    orientation: "minMax",
  };

  it("is zero without a title", () => {
    expect(getAxisTitleSpacePx(axis)).toBe(0);
  });

  it("is zero for a deleted axis", () => {
    expect(getAxisTitleSpacePx({ ...axis, deleted: true, title: "Revenue" })).toBe(0);
  });

  it("scales with the title font size", () => {
    expect(getAxisTitleSpacePx({ ...axis, title: "Revenue" })).toBe(18);
    expect(getAxisTitleSpacePx({ ...axis, title: "Revenue", titleStyle: { fontSize: 13 } })).toBe(
      23,
    );
  });

  it("grows a line height per extra line", () => {
    expect(getAxisTitleSpacePx({ ...axis, title: "Revenue\nUSD" })).toBe(31);
    expect(getAxisTitleSpacePx({ ...axis, title: "Revenue\nUSD\nnet" })).toBe(43);
  });
});
