import { describe, expect, it } from "vitest";

import { buildPresentation } from "../model/presentation";
import { parseZip } from "../parser/zip-parser";
import { renderSlide } from "../renderer/slide-renderer";
import { buildPptxWithShapes } from "./minimal-pptx";

async function renderShape(spPrInner: string): Promise<HTMLElement> {
  const buffer = await buildPptxWithShapes(`<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
${spPrInner}
</p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>`);

  const files = await parseZip(buffer);
  const presentation = buildPresentation(files);
  return renderSlide(presentation, presentation.slides[0]).element;
}

function mainPath(element: HTMLElement): SVGPathElement {
  const path = element.querySelector("svg path");
  if (!path) throw new Error("no svg path rendered");
  return path as SVGPathElement;
}

describe("shape fills", () => {
  it("renders solidFill srgbClr as the path fill", async () => {
    const element = await renderShape(
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
       <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>`,
    );
    expect(mainPath(element).getAttribute("fill")?.toLowerCase()).toBe("#4472c4");
  });

  it("resolves schemeClr through the theme color scheme", async () => {
    const element = await renderShape(
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
       <a:solidFill><a:schemeClr val="accent2"/></a:solidFill>`,
    );
    // Theme accent2 = ED7D31 in the fixture.
    expect(mainPath(element).getAttribute("fill")?.toLowerCase()).toBe("#ed7d31");
  });

  it("applies alpha as rgba fill", async () => {
    const element = await renderShape(
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
       <a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill>`,
    );
    const fill = mainPath(element).getAttribute("fill") ?? "";
    expect(fill).toMatch(/rgba\(\s*255\s*,\s*0\s*,\s*0\s*,\s*0?\.50*\s*\)/);
  });

  it("renders noFill as fill=none", async () => {
    const element = await renderShape(`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>`);
    expect(mainPath(element).getAttribute("fill")).toBe("none");
  });

  it("renders gradFill as an SVG gradient reference", async () => {
    const element = await renderShape(
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
       <a:gradFill>
         <a:gsLst>
           <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
           <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
         </a:gsLst>
         <a:lin ang="5400000" scaled="1"/>
       </a:gradFill>`,
    );
    const fill = mainPath(element).getAttribute("fill") ?? "";
    expect(fill).toMatch(/^url\(#/);
    const gradient = element.querySelector("svg defs linearGradient");
    expect(gradient).not.toBeNull();
    const stops = gradient!.querySelectorAll("stop");
    expect(stops.length).toBe(2);
  });
});

describe("shape strokes", () => {
  it("applies a:ln solidFill and width to the path stroke", async () => {
    const element = await renderShape(
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
       <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
       <a:ln w="25400"><a:solidFill><a:srgbClr val="1F1F1F"/></a:solidFill></a:ln>`,
    );
    const path = mainPath(element);
    expect(path.getAttribute("stroke")?.toLowerCase()).toBe("#1f1f1f");
    // 25400 EMU = 2pt = 2.6667px
    expect(Number(path.getAttribute("stroke-width"))).toBeCloseTo(2.6667, 3);
  });

  it("draws dashed strokes via stroke-dasharray", async () => {
    const element = await renderShape(
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
       <a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="dash"/></a:ln>`,
    );
    expect(mainPath(element).getAttribute("stroke-dasharray")).toBeTruthy();
  });
});

describe("shape placement", () => {
  it("positions the wrapper from xfrm off/ext", async () => {
    const element = await renderShape(
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
       <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>`,
    );
    const wrapper = element.querySelector("svg")?.closest("div");
    expect(wrapper).not.toBeNull();
    expect(wrapper!.style.left).toBe("96px");
    expect(wrapper!.style.top).toBe("96px");
    expect(wrapper!.style.width).toBe("192px");
  });

  it("applies rotation from xfrm@rot", async () => {
    const buffer = await buildPptxWithShapes(`<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Rotated"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm rot="2700000"><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
</p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>`);
    const files = await parseZip(buffer);
    const presentation = buildPresentation(files);
    const element = renderSlide(presentation, presentation.slides[0]).element;
    const wrapper = element.querySelector("svg")?.closest("div");
    // 2700000 / 60000 = 45 degrees
    expect(wrapper!.style.transform).toContain("rotate(45deg)");
  });
});
