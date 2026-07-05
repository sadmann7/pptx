import { beforeAll, describe, expect, it } from "vitest";

import { parseXml } from "../ooxml/xml-parser";
import { extractBackgroundColors, extractChartFrameStyle } from "../renderer/chart/frame";
import {
  buildChartPalette,
  createChartRenderContext,
  getVaryColorPointPalette,
  parseChartStyleId,
} from "../renderer/chart/palette";
import type { RenderContext } from "../renderer/render-context";
import { parseChartFragment, chartSpaceXml, createChartTestContext } from "./chart-pptx";

let ctx: RenderContext;

beforeAll(async () => {
  ctx = await createChartTestContext();
});

const OFFICE_ACCENTS = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"];

describe("buildChartPalette", () => {
  it("falls back to the theme accent colors", () => {
    const chartXml = parseXml(chartSpaceXml("<c:plotArea/>"));
    expect(buildChartPalette(chartXml, ctx)).toEqual(OFFICE_ACCENTS);
  });

  it("prefers a chart colors part palette when registered for the chart path", () => {
    const chartXml = parseXml(chartSpaceXml("<c:plotArea/>"));
    const colorStyle = parseChartFragment(
      `<c:colorStyle><a:srgbClr val="111111"/><a:srgbClr val="222222"/></c:colorStyle>`,
    ).child("colorStyle");
    const ctxWithColors: RenderContext = {
      ...ctx,
      presentation: {
        ...ctx.presentation,
        chartColorStyles: new Map([["ppt/charts/chart1.xml", colorStyle]]),
      },
    };
    expect(buildChartPalette(chartXml, ctxWithColors, "ppt/charts/chart1.xml")).toEqual([
      "#111111",
      "#222222",
    ]);
    // Without a matching path the theme accents win
    expect(buildChartPalette(chartXml, ctxWithColors, "ppt/charts/other.xml")).toEqual(
      OFFICE_ACCENTS,
    );
  });
});

describe("parseChartStyleId", () => {
  it("reads the c:style val", () => {
    const chartXml = parseXml(chartSpaceXml("<c:plotArea/>", `<c:style val="201"/>`));
    expect(parseChartStyleId(chartXml)).toBe(201);
  });

  it("returns undefined when absent", () => {
    expect(parseChartStyleId(parseXml(chartSpaceXml("<c:plotArea/>")))).toBeUndefined();
  });
});

describe("getVaryColorPointPalette", () => {
  it("returns the raw accents when darkening is disabled", () => {
    expect(getVaryColorPointPalette(ctx, { darken: false })).toEqual(OFFICE_ACCENTS);
  });

  it("darkens each accent by 12% by default", () => {
    const palette = getVaryColorPointPalette(ctx);
    // #4472C4 → r 68×0.88=60, g 114×0.88=100, b 196×0.88=172
    expect(palette[0]).toBe("#3c64ac");
    expect(palette).toHaveLength(6);
  });
});

describe("createChartRenderContext", () => {
  it("returns the original context without a color map override", () => {
    const chartXml = parseXml(chartSpaceXml("<c:plotArea/>"));
    expect(createChartRenderContext(chartXml, ctx)).toBe(ctx);
  });

  it("applies clrMapOvr overrideClrMapping attributes to the layout", () => {
    const chartXml = parseXml(
      chartSpaceXml(
        "<c:plotArea/>",
        `<c:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1"/></c:clrMapOvr>`,
      ),
    );
    const overridden = createChartRenderContext(chartXml, ctx);
    expect(overridden).not.toBe(ctx);
    expect(overridden.layout.colorMapOverride?.get("bg1")).toBe("dk1");
    expect(overridden.layout.colorMapOverride?.get("tx1")).toBe("lt1");
  });
});

describe("extractBackgroundColors", () => {
  function backgrounds(chartSpaceExtras: string, plotAreaExtras = "") {
    const chartXml = parseXml(
      chartSpaceXml(`<c:plotArea><c:layout/>${plotAreaExtras}</c:plotArea>`, chartSpaceExtras),
    );
    return extractBackgroundColors(chartXml, chartXml.child("chart"), ctx);
  }

  it("resolves explicit chart space and plot area solid fills", () => {
    const { chartBg, plotAreaBg } = backgrounds(
      `<c:spPr><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill></c:spPr>`,
      `<c:spPr><a:solidFill><a:srgbClr val="EEEEEE"/></a:solidFill></c:spPr>`,
    );
    expect(chartBg).toBe("#FFFF00");
    expect(plotAreaBg).toBe("#EEEEEE");
  });

  it("defaults the chart background to white when spPr has no fill info", () => {
    expect(backgrounds("<c:spPr/>").chartBg).toBe("#ffffff");
  });

  it("keeps backgrounds transparent for noFill and missing spPr", () => {
    expect(backgrounds("<c:spPr><a:noFill/></c:spPr>").chartBg).toBeUndefined();
    const { chartBg, plotAreaBg } = backgrounds("");
    expect(chartBg).toBeUndefined();
    expect(plotAreaBg).toBeUndefined();
  });
});

describe("extractChartFrameStyle", () => {
  it("maps the chart space outline to border styles", () => {
    const chartXml = parseXml(
      chartSpaceXml(
        "<c:plotArea/>",
        `<c:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="336699"/></a:solidFill></a:ln></c:spPr>`,
      ),
    );
    // 25400 EMU = 2 pt ≈ 2.667px
    expect(extractChartFrameStyle(chartXml, ctx)).toEqual({
      borderColor: "#336699",
      borderWidth: 25400 / 9525,
      borderStyle: "solid",
    });
  });

  it("returns undefined for noFill outlines or no outline", () => {
    const noLine = parseXml(chartSpaceXml("<c:plotArea/>"));
    expect(extractChartFrameStyle(noLine, ctx)).toBeUndefined();
    const noFill = parseXml(
      chartSpaceXml("<c:plotArea/>", `<c:spPr><a:ln w="12700"><a:noFill/></a:ln></c:spPr>`),
    );
    expect(extractChartFrameStyle(noFill, ctx)).toBeUndefined();
  });
});
