import { beforeAll, describe, expect, it } from "vitest";

import {
  chartTextStyleToEChartsTextStyle,
  extractTitleRichText,
  extractTitleText,
  extractTitleTextStyle,
  extractTxPrStyle,
  getChartThemeFontFamily,
} from "../renderer/chart/text";
import type { RenderContext } from "../renderer/render-context";
import { parseChartFragment, createChartTestContext } from "./chart-pptx";

let ctx: RenderContext;

beforeAll(async () => {
  ctx = await createChartTestContext();
});

describe("extractTitleText", () => {
  it("joins rich text runs and converts breaks to newlines", () => {
    const title = parseChartFragment(`<c:title><c:tx><c:rich>
<a:bodyPr/>
<a:p><a:r><a:t>Hello</a:t></a:r><a:br/><a:r><a:t>World</a:t></a:r></a:p>
</c:rich></c:tx></c:title>`).child("title");
    expect(extractTitleText(title)).toBe("Hello\nWorld");
  });

  it("joins multiple paragraphs with newlines", () => {
    const title = parseChartFragment(`<c:title><c:tx><c:rich>
<a:p><a:r><a:t>Line 1</a:t></a:r></a:p>
<a:p><a:r><a:t>Line 2</a:t></a:r></a:p>
</c:rich></c:tx></c:title>`).child("title");
    expect(extractTitleText(title)).toBe("Line 1\nLine 2");
  });

  it("falls back to strRef caches", () => {
    const title = parseChartFragment(`<c:title><c:tx><c:strRef><c:f>Sheet1!$A$1</c:f>
<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Cached Title</c:v></c:pt></c:strCache>
</c:strRef></c:tx></c:title>`).child("title");
    expect(extractTitleText(title)).toBe("Cached Title");
  });

  it("returns undefined without a tx element", () => {
    expect(extractTitleText(parseChartFragment("<c:title/>").child("title"))).toBeUndefined();
  });
});

describe("extractTitleRichText", () => {
  it("maps styled runs into ECharts rich text tokens", () => {
    const title = parseChartFragment(`<c:title><c:tx><c:rich>
<a:p>
<a:r><a:rPr lang="en-US" sz="1800" b="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Big</a:t></a:r>
<a:r><a:t> plain</a:t></a:r>
</a:p>
</c:rich></c:tx></c:title>`).child("title");
    const rich = extractTitleRichText(title, ctx);
    expect(rich).toBeDefined();
    expect(rich!.text).toBe("{r0|Big} plain");
    expect(rich!.rich.r0).toMatchObject({
      color: "#FF0000",
      fontSize: 18,
      fontWeight: "bold",
    });
  });

  it("escapes rich-text control characters in run content", () => {
    const title = parseChartFragment(`<c:title><c:tx><c:rich>
<a:p><a:r><a:rPr sz="1400"/><a:t>A{B|C}</a:t></a:r></a:p>
</c:rich></c:tx></c:title>`).child("title");
    const rich = extractTitleRichText(title, ctx);
    expect(rich!.text).toBe("{r0|A\\{B\\|C\\}}");
  });

  it("returns undefined when no run carries style", () => {
    const title = parseChartFragment(`<c:title><c:tx><c:rich>
<a:p><a:r><a:t>plain</a:t></a:r></a:p>
</c:rich></c:tx></c:title>`).child("title");
    expect(extractTitleRichText(title, ctx)).toBeUndefined();
  });
});

describe("extractTxPrStyle / extractTitleTextStyle", () => {
  it("extracts color, size, boldness, and font family from defRPr", () => {
    const legend = parseChartFragment(`<c:legend>
<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1200" b="1"><a:solidFill><a:srgbClr val="00B050"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr>
</c:legend>`).child("legend");
    const style = extractTxPrStyle(legend, ctx);
    expect(style).toMatchObject({ color: "#00B050", fontSize: 12, bold: true });
    expect(style?.fontFamily).toContain("Arial");
  });

  it("returns undefined without txPr", () => {
    expect(
      extractTxPrStyle(parseChartFragment("<c:legend/>").child("legend"), ctx),
    ).toBeUndefined();
  });

  it("reads title style from the rich text paragraph properties", () => {
    const title = parseChartFragment(`<c:title><c:tx><c:rich>
<a:bodyPr/>
<a:p><a:pPr><a:defRPr sz="1600"/></a:pPr><a:r><a:t>T</a:t></a:r></a:p>
</c:rich></c:tx></c:title>`).child("title");
    expect(extractTitleTextStyle(title, ctx)).toMatchObject({ fontSize: 16 });
  });
});

describe("chartTextStyleToEChartsTextStyle", () => {
  it("maps bold to fontWeight and passes through visual props", () => {
    expect(
      chartTextStyleToEChartsTextStyle({ color: "#112233", fontSize: 11, bold: false }),
    ).toEqual({ color: "#112233", fontSize: 11, fontWeight: "normal" });
  });

  it("returns undefined for undefined or empty styles", () => {
    expect(chartTextStyleToEChartsTextStyle(undefined)).toBeUndefined();
    expect(chartTextStyleToEChartsTextStyle({})).toBeUndefined();
  });
});

describe("getChartThemeFontFamily", () => {
  it("resolves the minor latin theme font", () => {
    // The minimal fixture theme declares Calibri as the minor latin font.
    expect(getChartThemeFontFamily(ctx)).toContain("Calibri");
  });
});
