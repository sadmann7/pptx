/**
 * Slide background rendering: slide/layout/master precedence, theme bgRef
 * references, gradient/pattern/image backgrounds.
 */
import { describe, expect, it } from "vitest";

import { buildPresentation } from "../../model/presentation";
import { parseZip } from "../../ooxml/zip-parser";
import { renderSlide } from "../../renderer/slide-renderer";
import { buildRichPptx, RichPptxOptions, tinyPngBytes } from "../helpers/rich-pptx";

async function renderWith(options: RichPptxOptions): Promise<HTMLElement> {
  const buffer = await buildRichPptx(options);
  const files = await parseZip(buffer);
  const presentation = buildPresentation(files);
  return renderSlide(presentation, presentation.slides[0]).element;
}

/** Normalize a CSS color (hex or rgb()) to a comparable lowercase hex string. */
function normalizeColor(value: string): string {
  const rgbMatch = value.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
  if (rgbMatch) {
    const [r, g, b] = [rgbMatch[1], rgbMatch[2], rgbMatch[3]].map(Number);
    return `#${[r, g, b].map((c) => c.toString(16).padStart(2, "0")).join("")}`;
  }
  return value.trim().toLowerCase();
}

const solidBg = (hex: string) =>
  `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`;

describe("background inheritance precedence", () => {
  it("uses the slide's own bgPr solidFill when present", async () => {
    const element = await renderWith({ slideBg: solidBg("FF00AA") });
    expect(normalizeColor(element.style.backgroundColor)).toBe("#ff00aa");
  });

  it("falls back to the layout background when the slide has none", async () => {
    const element = await renderWith({
      layoutBg: `<p:bg><p:bgPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`,
    });
    expect(normalizeColor(element.style.backgroundColor)).toBe("#4472c4");
  });

  it("falls back to the master background (lt1 = white) when slide and layout have none", async () => {
    const element = await renderWith({});
    expect(normalizeColor(element.style.backgroundColor)).toBe("#ffffff");
  });

  it("slide background wins over layout background", async () => {
    const element = await renderWith({
      slideBg: solidBg("112233"),
      layoutBg: solidBg("445566"),
    });
    expect(normalizeColor(element.style.backgroundColor)).toBe("#112233");
  });

  it("renders white when no background exists anywhere in the chain", async () => {
    const element = await renderWith({ masterBg: "" });
    expect(normalizeColor(element.style.backgroundColor)).toBe("#ffffff");
  });
});

describe("background fill variants (bgPr)", () => {
  it("composites a semi-transparent solid background on white", async () => {
    const element = await renderWith({
      slideBg: `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:solidFill><a:effectLst/></p:bgPr></p:bg>`,
    });
    // 50% black over white = rgb(128,128,128); output stays opaque.
    expect(normalizeColor(element.style.backgroundColor)).toBe("#808080");
  });

  it("renders bgPr noFill as opaque white", async () => {
    const element = await renderWith({
      slideBg: `<p:bg><p:bgPr><a:noFill/><a:effectLst/></p:bgPr></p:bg>`,
    });
    expect(normalizeColor(element.style.backgroundColor)).toBe("#ffffff");
  });

  it("renders a linear gradient background as CSS linear-gradient", async () => {
    const element = await renderWith({
      slideBg: `<p:bg><p:bgPr>
<a:gradFill>
<a:gsLst>
<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
</a:gsLst>
<a:lin ang="5400000" scaled="1"/>
</a:gradFill>
<a:effectLst/></p:bgPr></p:bg>`,
    });
    expect(element.style.background).toContain("linear-gradient(");
  });

  it("renders a radial (path) gradient background as an SVG overlay", async () => {
    const element = await renderWith({
      slideBg: `<p:bg><p:bgPr>
<a:gradFill>
<a:gsLst>
<a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>
<a:gs pos="100000"><a:srgbClr val="4472C4"/></a:gs>
</a:gsLst>
<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>
</a:gradFill>
<a:effectLst/></p:bgPr></p:bg>`,
    });
    const svg = element.querySelector('svg[data-pptx-background-gradient="true"]');
    expect(svg).not.toBeNull();
    const radial = svg!.querySelector("radialGradient");
    expect(radial).not.toBeNull();
    // 1280x720 slide: gradient centered with radius to the farthest corner.
    expect(Number(radial!.getAttribute("cx"))).toBeCloseTo(640, 3);
    expect(Number(radial!.getAttribute("cy"))).toBeCloseTo(360, 3);
    const rect = svg!.querySelector("rect");
    expect(rect!.getAttribute("fill")).toMatch(/^url\(#bg-grad-/);
  });

  it("renders a pattern background as tiled CSS gradient layers", async () => {
    const element = await renderWith({
      slideBg: `<p:bg><p:bgPr>
<a:pattFill prst="ltUpDiag">
<a:fgClr><a:srgbClr val="FF0000"/></a:fgClr>
<a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr>
</a:pattFill>
<a:effectLst/></p:bgPr></p:bg>`,
    });
    expect(element.style.backgroundImage).toContain("repeating-linear-gradient(-45deg");
    expect(element.style.backgroundSize).toBe("8px 8px");
    expect(element.style.backgroundRepeat).toBe("repeat");
    expect(normalizeColor(element.style.backgroundColor)).toBe("#ffffff");
  });

  it("renders a blipFill background image via a blob URL", async () => {
    const element = await renderWith({
      slideBg: `<p:bg><p:bgPr>
<a:blipFill><a:blip r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>
<a:effectLst/></p:bgPr></p:bg>`,
      extraSlideRelsXml: `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>`,
      binaryParts: { "ppt/media/image1.png": tinyPngBytes() },
    });
    expect(element.style.backgroundImage).toMatch(/^url\("blob:/);
    expect(element.style.backgroundSize).toBe("100% 100%");
    expect(element.style.backgroundRepeat).toBe("no-repeat");
  });
});

describe("theme background references (bgRef)", () => {
  it("resolves idx >= 1001 against bgFillStyleLst with phClr from the bgRef color", async () => {
    const element = await renderWith({
      slideBg: `<p:bg><p:bgRef idx="1001"><a:schemeClr val="accent2"/></p:bgRef></p:bg>`,
    });
    expect(normalizeColor(element.style.backgroundColor)).toBe("#ed7d31");
  });

  it("falls back to fillStyleLst for low bgRef indices", async () => {
    const element = await renderWith({
      slideBg: `<p:bg><p:bgRef idx="2"><a:schemeClr val="accent1"/></p:bgRef></p:bg>`,
    });
    expect(normalizeColor(element.style.backgroundColor)).toBe("#4472c4");
  });

  it("uses the bgRef color directly when the index has no theme entry", async () => {
    const element = await renderWith({
      slideBg: `<p:bg><p:bgRef idx="1099"><a:srgbClr val="00FF00"/></p:bgRef></p:bg>`,
    });
    expect(normalizeColor(element.style.backgroundColor)).toBe("#00ff00");
  });
});
