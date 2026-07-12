/**
 * Gradient fills (linear angle mapping, modifier stops, radial paths) and
 * pattern fills rendered as SVG by the shape renderer.
 */
import { describe, expect, it } from "vitest";

import { buildPresentation } from "../../model/presentation";
import { readPptx } from "../../ooxml/zip";
import { renderSlide } from "../../renderer/slide";
import { buildPptxWithShapes } from "../fixtures/minimal-pptx";

async function renderShape(spPrInner: string): Promise<HTMLElement> {
  const buffer = await buildPptxWithShapes(`<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
${spPrInner}
</p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>`);

  const files = await readPptx(buffer);
  const presentation = buildPresentation(files);
  return renderSlide(presentation, presentation.slides[0]).element;
}

function mainPath(element: HTMLElement): SVGPathElement {
  const path = element.querySelector("svg path");
  if (!path) throw new Error("no svg path rendered");
  return path as SVGPathElement;
}

function hexChannels(hex: string): [number, number, number] {
  const clean = hex.replace("#", "");
  return [
    parseInt(clean.slice(0, 2), 16),
    parseInt(clean.slice(2, 4), 16),
    parseInt(clean.slice(4, 6), 16),
  ];
}

describe("linear gradient fills", () => {
  it("resolves schemeClr stops with lumMod modifiers", async () => {
    const element = await renderShape(
      `<a:gradFill>
         <a:gsLst>
           <a:gs pos="0"><a:schemeClr val="lt1"><a:lumMod val="50000"/></a:schemeClr></a:gs>
           <a:gs pos="100000"><a:schemeClr val="accent1"/></a:gs>
         </a:gsLst>
         <a:lin ang="0" scaled="1"/>
       </a:gradFill>`,
    );
    const stops = [...element.querySelectorAll("svg defs linearGradient stop")];
    expect(stops).toHaveLength(2);
    // lt1 (FFFFFF) with lumMod 50% -> HSL lightness halved -> #808080.
    expect(stops[0].getAttribute("stop-color")?.toLowerCase()).toBe("#808080");
    expect(stops[0].getAttribute("offset")).toBe("0%");
    expect(stops[1].getAttribute("stop-color")?.toLowerCase()).toBe("#4472c4");
    expect(stops[1].getAttribute("offset")).toBe("100%");
  });

  it("lightens a stop with a tint modifier", async () => {
    const element = await renderShape(
      `<a:gradFill>
         <a:gsLst>
           <a:gs pos="0"><a:schemeClr val="accent1"><a:tint val="50000"/></a:schemeClr></a:gs>
           <a:gs pos="100000"><a:schemeClr val="accent1"/></a:gs>
         </a:gsLst>
         <a:lin ang="0" scaled="1"/>
       </a:gradFill>`,
    );
    const stops = [...element.querySelectorAll("svg defs linearGradient stop")];
    const tinted = hexChannels(stops[0].getAttribute("stop-color") ?? "#000000");
    const base = hexChannels("#4472C4");
    // Tint mixes toward white: every channel must strictly increase.
    for (let i = 0; i < 3; i++) {
      expect(tinted[i]).toBeGreaterThan(base[i]);
      expect(tinted[i]).toBeLessThanOrEqual(255);
    }
  });

  it("sorts out-of-order stops by position", async () => {
    const element = await renderShape(
      `<a:gradFill>
         <a:gsLst>
           <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
           <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
           <a:gs pos="50000"><a:srgbClr val="00FF00"/></a:gs>
         </a:gsLst>
         <a:lin ang="0" scaled="1"/>
       </a:gradFill>`,
    );
    const stops = [...element.querySelectorAll("svg defs linearGradient stop")];
    expect(stops.map((s) => s.getAttribute("offset"))).toEqual(["0%", "50%", "100%"]);
    expect(stops.map((s) => s.getAttribute("stop-color")?.toLowerCase())).toEqual([
      "#ff0000",
      "#00ff00",
      "#0000ff",
    ]);
  });

  it("maps a:lin ang=0 to a left-to-right SVG gradient axis", async () => {
    const element = await renderShape(
      `<a:gradFill>
         <a:gsLst>
           <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
           <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
         </a:gsLst>
         <a:lin ang="0" scaled="1"/>
       </a:gradFill>`,
    );
    const gradient = element.querySelector("svg defs linearGradient")!;
    expect(gradient.getAttribute("gradientUnits")).toBe("userSpaceOnUse");
    // Shape is 192x96px: axis runs (0,48) -> (192,48).
    expect(Number(gradient.getAttribute("x1"))).toBeCloseTo(0, 5);
    expect(Number(gradient.getAttribute("y1"))).toBeCloseTo(48, 5);
    expect(Number(gradient.getAttribute("x2"))).toBeCloseTo(192, 5);
    expect(Number(gradient.getAttribute("y2"))).toBeCloseTo(48, 5);
  });

  it("maps a:lin ang=5400000 (90 deg) to a top-to-bottom SVG gradient axis", async () => {
    const element = await renderShape(
      `<a:gradFill>
         <a:gsLst>
           <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
           <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
         </a:gsLst>
         <a:lin ang="5400000" scaled="1"/>
       </a:gradFill>`,
    );
    const gradient = element.querySelector("svg defs linearGradient")!;
    expect(Number(gradient.getAttribute("x1"))).toBeCloseTo(96, 5);
    expect(Number(gradient.getAttribute("y1"))).toBeCloseTo(0, 5);
    expect(Number(gradient.getAttribute("x2"))).toBeCloseTo(96, 5);
    expect(Number(gradient.getAttribute("y2"))).toBeCloseTo(96, 5);
  });
});

describe("radial (path) gradient fills", () => {
  it("renders a:path path=circle as a centered SVG radialGradient", async () => {
    const element = await renderShape(
      `<a:gradFill>
         <a:gsLst>
           <a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>
           <a:gs pos="100000"><a:srgbClr val="4472C4"/></a:gs>
         </a:gsLst>
         <a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>
       </a:gradFill>`,
    );
    const path = mainPath(element);
    expect(path.getAttribute("fill")).toMatch(/^url\(#/);

    const gradient = element.querySelector("svg defs radialGradient")!;
    expect(gradient).not.toBeNull();
    expect(gradient.getAttribute("gradientUnits")).toBe("userSpaceOnUse");
    // Center of the 192x96 shape; radius reaches the farthest corner.
    expect(Number(gradient.getAttribute("cx"))).toBeCloseTo(96, 5);
    expect(Number(gradient.getAttribute("cy"))).toBeCloseTo(48, 5);
    expect(Number(gradient.getAttribute("r"))).toBeCloseTo(Math.hypot(96, 48), 3);

    const stops = [...gradient.querySelectorAll("stop")];
    expect(stops).toHaveLength(2);
    expect(stops[0].getAttribute("stop-color")?.toLowerCase()).toBe("#ffffff");
    expect(stops[1].getAttribute("stop-color")?.toLowerCase()).toBe("#4472c4");
  });

  it("offsets the radial center from a:fillToRect", async () => {
    // fillToRect l=100000 t=100000 collapses the focus to the bottom-right corner.
    const element = await renderShape(
      `<a:gradFill>
         <a:gsLst>
           <a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>
           <a:gs pos="100000"><a:srgbClr val="4472C4"/></a:gs>
         </a:gsLst>
         <a:path path="circle"><a:fillToRect l="100000" t="100000"/></a:path>
       </a:gradFill>`,
    );
    const gradient = element.querySelector("svg defs radialGradient")!;
    expect(Number(gradient.getAttribute("cx"))).toBeCloseTo(192, 5);
    expect(Number(gradient.getAttribute("cy"))).toBeCloseTo(96, 5);
  });

  it("renders a:path path=rect as blended horizontal+vertical linear gradients", async () => {
    const element = await renderShape(
      `<a:gradFill>
         <a:gsLst>
           <a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>
           <a:gs pos="100000"><a:srgbClr val="4472C4"/></a:gs>
         </a:gsLst>
         <a:path path="rect"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>
       </a:gradFill>`,
    );
    // The main path defers filling to a blend group of two linear gradients.
    const gradients = element.querySelectorAll("svg defs linearGradient");
    expect(gradients.length).toBe(2);
    const blended = element.querySelectorAll('svg g[style*="isolation"] path');
    expect(blended.length).toBe(2);
    // The top-level shape path itself carries no fill of its own.
    const topLevelPath = element.querySelector("svg > path");
    expect(topLevelPath).not.toBeNull();
    expect(topLevelPath!.getAttribute("fill")).toBe("none");
  });
});

describe("pattern fills (a:pattFill)", () => {
  it("renders prst=ltUpDiag as an SVG pattern with fg/bg colors", async () => {
    const element = await renderShape(
      `<a:pattFill prst="ltUpDiag">
         <a:fgClr><a:srgbClr val="FF0000"/></a:fgClr>
         <a:bgClr><a:srgbClr val="FFFF00"/></a:bgClr>
       </a:pattFill>`,
    );
    const path = mainPath(element);
    expect(path.getAttribute("fill")).toMatch(/^url\(#shape-pattern-/);

    const pattern = element.querySelector("svg defs pattern")!;
    expect(pattern).not.toBeNull();
    expect(pattern.getAttribute("patternUnits")).toBe("userSpaceOnUse");

    const bgRect = pattern.querySelector("rect")!;
    expect(bgRect.getAttribute("fill")?.toLowerCase()).toBe("#ffff00");

    const fgLine = pattern.querySelector("line")!;
    expect(fgLine).not.toBeNull();
    expect(fgLine.getAttribute("stroke")?.toLowerCase()).toBe("#ff0000");
  });

  it("resolves pattern fg/bg through the theme color scheme", async () => {
    const element = await renderShape(
      `<a:pattFill prst="pct50">
         <a:fgClr><a:schemeClr val="accent1"/></a:fgClr>
         <a:bgClr><a:schemeClr val="accent2"/></a:bgClr>
       </a:pattFill>`,
    );
    const pattern = element.querySelector("svg defs pattern")!;
    expect(pattern.querySelector("rect")!.getAttribute("fill")?.toLowerCase()).toBe("#ed7d31");
    expect(pattern.querySelector("circle")!.getAttribute("fill")?.toLowerCase()).toBe("#4472c4");
  });
});
