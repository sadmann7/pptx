/**
 * Shape effects (a:effectLst outerShdw) and flip transforms on plain shapes.
 */
import { describe, expect, it } from "vitest";

import { buildPresentation } from "../../model/presentation";
import { parseZip } from "../../ooxml/zip";
import { renderSlide } from "../../renderer/slide";
import { buildPptxWithShapes } from "../fixtures/minimal-pptx";

async function renderShapes(shapesXml: string): Promise<HTMLElement> {
  const buffer = await buildPptxWithShapes(shapesXml);
  const files = await parseZip(buffer);
  const presentation = buildPresentation(files);
  return renderSlide(presentation, presentation.slides[0]).element;
}

function rectShape(spPrInner: string, xfrmAttrs = ""): string {
  return `<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm${xfrmAttrs}><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
${spPrInner}
</p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>`;
}

function mainPath(element: HTMLElement): SVGPathElement {
  const path = element.querySelector("svg path");
  if (!path) throw new Error("no svg path rendered");
  return path as SVGPathElement;
}

describe("outer shadow effects", () => {
  it("renders a:outerShdw as an SVG feDropShadow filter on the shape path", async () => {
    const element = await renderShapes(
      rectShape(`<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
<a:effectLst>
<a:outerShdw blurRad="76200" dist="38100" dir="0">
<a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr>
</a:outerShdw>
</a:effectLst>`),
    );

    const dropShadow = element.querySelector("svg defs filter feDropShadow")!;
    expect(dropShadow).not.toBeNull();
    // dist 38100 EMU = 4px, dir 0 -> dx=4 dy=0; blurRad 76200 EMU = 8px -> stdDeviation 4.
    expect(dropShadow.getAttribute("dx")).toBe("4.0");
    expect(dropShadow.getAttribute("dy")).toBe("0.0");
    expect(dropShadow.getAttribute("stdDeviation")).toBe("4.00");
    expect(dropShadow.getAttribute("flood-color")).toBe("rgb(0,0,0)");
    expect(dropShadow.getAttribute("flood-opacity")).toBe("0.5000");
    expect(mainPath(element).getAttribute("filter")).toMatch(/^url\(#shape-shadow/);
  });

  it("projects the shadow offset along the dir angle", async () => {
    const element = await renderShapes(
      rectShape(`<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
<a:effectLst>
<a:outerShdw blurRad="0" dist="38100" dir="5400000">
<a:srgbClr val="000000"/>
</a:outerShdw>
</a:effectLst>`),
    );
    const dropShadow = element.querySelector("svg defs filter feDropShadow")!;
    // dir 5400000 = 90 degrees -> shadow drops straight down.
    expect(dropShadow.getAttribute("dx")).toBe("0.0");
    expect(dropShadow.getAttribute("dy")).toBe("4.0");
  });

  it("falls back to a CSS drop-shadow filter for line-like shapes", async () => {
    const element = await renderShapes(`<p:cxnSp>
<p:nvCxnSpPr><p:cNvPr id="30" name="Connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
<p:spPr>
<a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
<a:effectLst>
<a:outerShdw blurRad="19050" dist="38100" dir="0">
<a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr>
</a:outerShdw>
</a:effectLst>
</p:spPr>
</p:cxnSp>`);

    const wrapper = element.querySelector("svg")?.closest("div") as HTMLElement;
    expect(wrapper.style.filter).toContain("drop-shadow(4.0px 0.0px 2.0px");
    expect(wrapper.style.filter).toContain("rgba(0,0,0,0.500)");
  });
});

describe("shape flip transforms", () => {
  it("applies flipH as scaleX(-1) on the shape wrapper", async () => {
    const element = await renderShapes(
      rectShape(`<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>`, ` flipH="1"`),
    );
    const wrapper = element.querySelector("svg")?.closest("div") as HTMLElement;
    expect(wrapper.style.transform).toBe("scaleX(-1)");
  });

  it("applies flipV as scaleY(-1) on the shape wrapper", async () => {
    const element = await renderShapes(
      rectShape(`<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>`, ` flipV="1"`),
    );
    const wrapper = element.querySelector("svg")?.closest("div") as HTMLElement;
    expect(wrapper.style.transform).toBe("scaleY(-1)");
  });

  it("combines rotation with both flips in transform order", async () => {
    const element = await renderShapes(
      rectShape(
        `<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>`,
        ` rot="5400000" flipH="1" flipV="1"`,
      ),
    );
    const wrapper = element.querySelector("svg")?.closest("div") as HTMLElement;
    expect(wrapper.style.transform).toBe("rotate(90deg) scaleX(-1) scaleY(-1)");
  });
});
