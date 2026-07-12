/**
 * Theme style references (`p:style` > fillRef/lnRef/effectRef) resolution
 * through the real parse/render pipeline.
 */
import { describe, expect, it } from "vitest";

import { buildPresentation } from "../../model/presentation";
import { readPptx } from "../../ooxml/zip";
import { renderSlide } from "../../renderer/slide";
import { buildPptxWithShapes } from "../fixtures/minimal-pptx";
import { buildRichPptx } from "../fixtures/rich-pptx";

async function renderShapes(shapesXml: string): Promise<HTMLElement> {
  const buffer = await buildPptxWithShapes(shapesXml);
  const files = await readPptx(buffer);
  const presentation = buildPresentation(files);
  return renderSlide(presentation, presentation.slides[0]).element;
}

function styledShape(styleXml: string, spPrInner = ""): string {
  return `<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Styled"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
${spPrInner}
</p:spPr>
${styleXml}
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>`;
}

function mainPath(element: HTMLElement): SVGPathElement {
  const path = element.querySelector("svg path");
  if (!path) throw new Error("no svg path rendered");
  return path as SVGPathElement;
}

describe("theme fill references (a:fillRef)", () => {
  it("resolves the fill from theme fillStyleLst with phClr = fillRef color", async () => {
    const element = await renderShapes(
      styledShape(`<p:style>
<a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef>
<a:fillRef idx="2"><a:schemeClr val="accent1"/></a:fillRef>
<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
</p:style>`),
    );
    // fillStyleLst[1] = solidFill schemeClr phClr; phClr = accent1 = 4472C4.
    expect(mainPath(element).getAttribute("fill")?.toLowerCase()).toBe("#4472c4");
  });

  it("applies color modifiers on the fillRef color before substituting phClr", async () => {
    const element = await renderShapes(
      styledShape(`<p:style>
<a:lnRef idx="0"><a:schemeClr val="lt1"/></a:lnRef>
<a:fillRef idx="1"><a:schemeClr val="lt1"><a:lumMod val="50000"/></a:schemeClr></a:fillRef>
<a:effectRef idx="0"><a:schemeClr val="lt1"/></a:effectRef>
<a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef>
</p:style>`),
    );
    // lt1 = FFFFFF, lumMod 50% halves HSL lightness -> #808080.
    expect(mainPath(element).getAttribute("fill")?.toLowerCase()).toBe("#808080");
  });

  it("prefers an explicit spPr fill over the fillRef", async () => {
    const element = await renderShapes(
      styledShape(
        `<p:style>
<a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef>
<a:fillRef idx="2"><a:schemeClr val="accent1"/></a:fillRef>
<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
</p:style>`,
        `<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>`,
      ),
    );
    expect(mainPath(element).getAttribute("fill")?.toLowerCase()).toBe("#ff0000");
  });
});

describe("theme line references (a:lnRef)", () => {
  it("resolves stroke color and width from theme lnStyleLst with phClr = lnRef color", async () => {
    const element = await renderShapes(
      styledShape(`<p:style>
<a:lnRef idx="1"><a:schemeClr val="accent2"/></a:lnRef>
<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
</p:style>`),
    );
    const path = mainPath(element);
    // lnStyleLst[0] = ln w=6350 solidFill phClr; phClr = accent2 = ED7D31.
    expect(path.getAttribute("stroke")?.toLowerCase()).toBe("#ed7d31");
    // 6350 EMU = 0.6667px at 96dpi
    expect(Number(path.getAttribute("stroke-width"))).toBeCloseTo(0.6667, 3);
  });

  it("uses the theme line width for the referenced index", async () => {
    const element = await renderShapes(
      styledShape(`<p:style>
<a:lnRef idx="3"><a:schemeClr val="accent1"/></a:lnRef>
<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
</p:style>`),
    );
    // lnStyleLst[2] = ln w=19050 EMU = 2px.
    expect(Number(mainPath(element).getAttribute("stroke-width"))).toBeCloseTo(2, 3);
  });

  it("explicit a:ln solidFill wins over the lnRef color", async () => {
    const element = await renderShapes(
      styledShape(
        `<p:style>
<a:lnRef idx="1"><a:schemeClr val="accent2"/></a:lnRef>
<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
</p:style>`,
        `<a:ln w="25400"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln>`,
      ),
    );
    const path = mainPath(element);
    expect(path.getAttribute("stroke")?.toLowerCase()).toBe("#00ff00");
    expect(Number(path.getAttribute("stroke-width"))).toBeCloseTo(2.6667, 3);
  });
});

describe("theme effect references (a:effectRef)", () => {
  it("applies an outer shadow from the theme effectStyleLst", async () => {
    // Custom theme whose 3rd effect style carries an outer shadow.
    const fmtScheme = `<a:fmtScheme name="Office">
<a:fillStyleLst>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
</a:fillStyleLst>
<a:lnStyleLst>
<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
</a:lnStyleLst>
<a:effectStyleLst>
<a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst>
<a:outerShdw blurRad="76200" dist="38100" dir="0">
<a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr>
</a:outerShdw>
</a:effectLst></a:effectStyle>
</a:effectStyleLst>
<a:bgFillStyleLst>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
</a:bgFillStyleLst>
</a:fmtScheme>`;

    const buffer = await buildRichPptx({
      fmtSchemeXml: fmtScheme,
      shapesXml: styledShape(`<p:style>
<a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef>
<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
<a:effectRef idx="3"><a:schemeClr val="accent1"/></a:effectRef>
<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
</p:style>`),
    });
    const files = await readPptx(buffer);
    const presentation = buildPresentation(files);
    const element = renderSlide(presentation, presentation.slides[0]).element;

    const dropShadow = element.querySelector("svg defs filter feDropShadow");
    expect(dropShadow).not.toBeNull();
    // dist 38100 EMU = 4px at dir 0 -> dx=4, dy=0; blurRad 76200 EMU = 8px -> stdDeviation 4.
    expect(dropShadow!.getAttribute("dx")).toBe("4.0");
    expect(dropShadow!.getAttribute("dy")).toBe("0.0");
    expect(dropShadow!.getAttribute("stdDeviation")).toBe("4.00");
    expect(dropShadow!.getAttribute("flood-opacity")).toBe("0.4000");

    const path = mainPath(element);
    expect(path.getAttribute("filter")).toMatch(/^url\(#shape-shadow/);
  });
});
