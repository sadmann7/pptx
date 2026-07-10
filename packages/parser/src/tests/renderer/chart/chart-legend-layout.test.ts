import { beforeAll, describe, expect, it } from "vitest";

import { numToPct } from "../../../renderer/chart/layout";
import {
  buildLegendOption,
  extractLegendInfo,
  getGridBottomPx,
  getGridTopPx,
  getLegendOptionObject,
  getLegendPlacement,
  legendIsAtTop,
  lineLegendIconPath,
  pickSeriesStringColor,
} from "../../../renderer/chart/legend";
import type { RenderContext } from "../../../renderer/render-context";
import { parseChartFragment, createChartTestContext } from "../../helpers/chart-pptx";

let ctx: RenderContext;

beforeAll(async () => {
  ctx = await createChartTestContext();
});

function legendInfoFor(xml: string) {
  return extractLegendInfo(parseChartFragment(xml).child("chart"), ctx);
}

describe("numToPct", () => {
  it("converts fractions to percent strings with 2-decimal precision", () => {
    expect(numToPct(0.5)).toBe("50%");
    expect(numToPct(0)).toBe("0%");
    expect(numToPct(1)).toBe("100%");
    expect(numToPct(0.123)).toBe("12.3%");
    expect(numToPct(0.12345)).toBe("12.35%");
  });
});

describe("extractLegendInfo", () => {
  it("returns undefined when no legend exists", () => {
    expect(legendInfoFor("<c:chart/>")).toBeUndefined();
  });

  it("maps each legendPos to the matching ECharts anchor", () => {
    const bottom = legendInfoFor(`<c:chart><c:legend><c:legendPos val="b"/></c:legend></c:chart>`);
    expect(bottom?.position).toBe("b");
    expect(bottom?.option).toMatchObject({ bottom: "5%", orient: "horizontal" });

    const top = legendInfoFor(`<c:chart><c:legend><c:legendPos val="t"/></c:legend></c:chart>`);
    expect(top?.option).toMatchObject({ top: "14%", orient: "horizontal" });

    const left = legendInfoFor(`<c:chart><c:legend><c:legendPos val="l"/></c:legend></c:chart>`);
    expect(left?.option).toMatchObject({ left: "2%", top: "middle", orient: "vertical" });

    const right = legendInfoFor(`<c:chart><c:legend><c:legendPos val="r"/></c:legend></c:chart>`);
    expect(right?.option).toMatchObject({ right: "2%", top: "middle", orient: "vertical" });

    const topRight = legendInfoFor(
      `<c:chart><c:legend><c:legendPos val="tr"/></c:legend></c:chart>`,
    );
    expect(topRight?.option).toMatchObject({ top: "14%", right: "2%", orient: "vertical" });
  });

  it("defaults to right placement for missing or unknown positions", () => {
    const noPos = legendInfoFor(`<c:chart><c:legend/></c:chart>`);
    expect(noPos?.position).toBe("r");
    const unknown = legendInfoFor(
      `<c:chart><c:legend><c:legendPos val="zz"/></c:legend></c:chart>`,
    );
    expect(unknown?.position).toBe("r");
  });

  it("parses the overlay flag", () => {
    const overlaid = legendInfoFor(
      `<c:chart><c:legend><c:legendPos val="r"/><c:overlay val="1"/></c:legend></c:chart>`,
    );
    expect(overlaid?.overlay).toBe(true);
    expect(getLegendPlacement(overlaid)).toBe("none");
  });

  it("parses c:layout manualLayout into percent insets", () => {
    const info = legendInfoFor(`<c:chart><c:legend><c:legendPos val="r"/>
<c:layout><c:manualLayout><c:x val="0.7"/><c:y val="0.2"/><c:w val="0.25"/><c:h val="0.6"/></c:manualLayout></c:layout>
</c:legend></c:chart>`);
    expect(info?.manualLayout).toEqual({
      left: "70%",
      top: "20%",
      width: "25%",
      height: "60%",
    });
  });

  it("extracts legend text style from txPr", () => {
    const info = legendInfoFor(`<c:chart><c:legend><c:legendPos val="b"/>
<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="900" b="1"><a:solidFill><a:srgbClr val="336699"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
</c:legend></c:chart>`);
    expect(info?.textStyle).toMatchObject({
      color: "#336699",
      fontSize: 9,
      fontWeight: "bold",
    });
  });
});

describe("legend placement helpers", () => {
  it("computes placement from the option anchors", () => {
    expect(
      getLegendPlacement(
        legendInfoFor(`<c:chart><c:legend><c:legendPos val="b"/></c:legend></c:chart>`),
      ),
    ).toBe("bottom");
    expect(
      getLegendPlacement(
        legendInfoFor(`<c:chart><c:legend><c:legendPos val="t"/></c:legend></c:chart>`),
      ),
    ).toBe("top");
    expect(
      getLegendPlacement(
        legendInfoFor(`<c:chart><c:legend><c:legendPos val="l"/></c:legend></c:chart>`),
      ),
    ).toBe("left");
    expect(getLegendPlacement(undefined)).toBe("none");
  });

  it("treats t and tr as top for grid sizing", () => {
    const top = legendInfoFor(`<c:chart><c:legend><c:legendPos val="t"/></c:legend></c:chart>`);
    const topRight = legendInfoFor(
      `<c:chart><c:legend><c:legendPos val="tr"/></c:legend></c:chart>`,
    );
    const right = legendInfoFor(`<c:chart><c:legend><c:legendPos val="r"/></c:legend></c:chart>`);
    expect(legendIsAtTop(top)).toBe(true);
    expect(legendIsAtTop(topRight)).toBe(true);
    expect(legendIsAtTop(right)).toBe(false);
    expect(getGridTopPx(true, top)).toBe(52);
    expect(getGridTopPx(false, right)).toBe(20);
  });

  it("reserves bottom grid space only for bottom legends", () => {
    const bottom = legendInfoFor(`<c:chart><c:legend><c:legendPos val="b"/></c:legend></c:chart>`);
    const right = legendInfoFor(`<c:chart><c:legend><c:legendPos val="r"/></c:legend></c:chart>`);
    expect(getGridBottomPx(bottom)).toBe(35);
    expect(getGridBottomPx(right)).toBe(20);
    expect(getGridBottomPx(undefined)).toBe(20);
  });
});

describe("buildLegendOption", () => {
  it("hides the legend when no legend option exists", () => {
    expect(buildLegendOption(undefined, undefined, undefined, ["A"], { fontSize: 10 })).toEqual({
      show: false,
    });
  });

  it("uses rect icons and font-sized items for plain string entries", () => {
    const info = legendInfoFor(`<c:chart><c:legend><c:legendPos val="b"/></c:legend></c:chart>`)!;
    const legend = getLegendOptionObject(
      buildLegendOption(info.option, info, undefined, ["A", "B"], { fontSize: 10 }),
    )!;
    expect(legend.icon).toBe("rect");
    expect(legend.itemWidth).toBe(10);
    expect(legend.itemHeight).toBe(10);
    expect(legend.data).toEqual(["A", "B"]);
  });

  it("widens items when entries use line path icons", () => {
    const info = legendInfoFor(`<c:chart><c:legend><c:legendPos val="b"/></c:legend></c:chart>`)!;
    const legend = getLegendOptionObject(
      buildLegendOption(info.option, info, undefined, [{ name: "A", icon: lineLegendIconPath() }], {
        fontSize: 10,
      }),
    )!;
    expect(legend.icon).toBeUndefined();
    expect(legend.itemWidth).toBe(24);
    expect(legend.itemHeight).toBe(9);
  });

  it("prefers manual layout insets over the computed top", () => {
    const info = legendInfoFor(`<c:chart><c:legend><c:legendPos val="t"/>
<c:layout><c:manualLayout><c:x val="0.1"/><c:y val="0.05"/></c:manualLayout></c:layout>
</c:legend></c:chart>`)!;
    const legend = getLegendOptionObject(
      buildLegendOption(info.option, info, 26, ["A"], { fontSize: 10 }),
    )!;
    expect(legend.left).toBe("10%");
    expect(legend.top).toBe("5%");
  });
});

describe("pickSeriesStringColor", () => {
  it("passes through strings and falls back for gradients", () => {
    expect(pickSeriesStringColor("#123456", "#000000")).toBe("#123456");
    expect(pickSeriesStringColor({ type: "linear" }, "#000000")).toBe("#000000");
    expect(pickSeriesStringColor(undefined, "#00B050")).toBe("#00B050");
  });
});
