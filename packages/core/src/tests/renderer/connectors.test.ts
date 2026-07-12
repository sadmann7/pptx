/**
 * Connector shapes (p:cxnSp): straight connector geometry, stroke,
 * arrowhead markers (headEnd/tailEnd), and flips.
 */
import { describe, expect, it } from "vitest";

import { buildPresentation } from "../../model/presentation";
import { readPptx } from "../../ooxml/zip";
import { renderSlide } from "../../renderer/slide";
import { buildPptxWithShapes } from "../fixtures/minimal-pptx";

async function renderConnector(spPrInner: string, extraAfterSpPr = ""): Promise<HTMLElement> {
  const buffer = await buildPptxWithShapes(`<p:cxnSp>
<p:nvCxnSpPr><p:cNvPr id="30" name="Connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
<p:spPr>
${spPrInner}
</p:spPr>
${extraAfterSpPr}
</p:cxnSp>`);

  const files = await readPptx(buffer);
  const presentation = buildPresentation(files);
  return renderSlide(presentation, presentation.slides[0]).element;
}

function mainPath(element: HTMLElement): SVGPathElement {
  const path = element.querySelector("svg path");
  if (!path) throw new Error("no svg path rendered");
  return path as SVGPathElement;
}

const XFRM = `<a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>`;

describe("connector strokes and geometry", () => {
  it("renders straightConnector1 as a stroked, unfilled diagonal path", async () => {
    const element = await renderConnector(
      `${XFRM}
<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
<a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>`,
    );
    const path = mainPath(element);
    expect(path.getAttribute("fill")).toBe("none");
    expect(path.getAttribute("stroke")?.toLowerCase()).toBe("#ff0000");
    expect(Number(path.getAttribute("stroke-width"))).toBeCloseTo(2.6667, 3);
    // 192x96px box: connector runs corner to corner.
    expect(path.getAttribute("d")).toBe("M0,0 L192,96");
  });

  it("keeps connectors stroke-only even when a style fillRef is present", async () => {
    const element = await renderConnector(
      `${XFRM}
<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>`,
      `<p:style>
<a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>
<a:fillRef idx="2"><a:schemeClr val="accent1"/></a:fillRef>
<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
</p:style>`,
    );
    expect(mainPath(element).getAttribute("fill")).toBe("none");
  });

  it("draws a cxnSp with prst=line as an endpoint-to-endpoint connector", async () => {
    // For connection shapes the 'line' preset must run (0,0)->(w,h), not a midline.
    const element = await renderConnector(
      `${XFRM}
<a:prstGeom prst="line"><a:avLst/></a:prstGeom>
<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>`,
    );
    expect(mainPath(element).getAttribute("d")).toBe("M0,0 L192,96");
  });

  it("flips the path coordinates for flipV instead of using a CSS transform", async () => {
    const element = await renderConnector(
      `<a:xfrm flipV="1"><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>`,
    );
    const path = mainPath(element);
    // Line now runs bottom-left to top-right.
    expect(path.getAttribute("d")).toBe("M0,96 L192,0");
    const wrapper = element.querySelector("svg")?.closest("div") as HTMLElement;
    expect(wrapper.style.transform).not.toContain("scale");
  });
});

describe("connector arrowheads (headEnd/tailEnd)", () => {
  it("renders triangle/arrow line ends as SVG markers", async () => {
    const element = await renderConnector(
      `${XFRM}
<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
<a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
<a:headEnd type="triangle"/><a:tailEnd type="arrow"/>
</a:ln>`,
    );
    const path = mainPath(element);
    const markerStart = path.getAttribute("marker-start");
    const markerEnd = path.getAttribute("marker-end");
    expect(markerStart).toMatch(/^url\(#arrow-marker-/);
    expect(markerEnd).toMatch(/^url\(#arrow-marker-/);
    expect(markerStart).not.toBe(markerEnd);

    const markers = [...element.querySelectorAll("svg defs marker")];
    expect(markers).toHaveLength(2);
    for (const marker of markers) {
      expect(marker.getAttribute("markerUnits")).toBe("userSpaceOnUse");
      expect(marker.getAttribute("orient")).toBe("auto");
    }

    // Head arrow points backward, tail arrow points forward; both use stroke color.
    const headId = markerStart!.slice(5, -1);
    const head = markers.find((m) => m.getAttribute("id") === headId)!;
    expect(head.querySelector("polygon")!.getAttribute("points")).toBe("0,5 10,0 10,10");
    expect(head.querySelector("polygon")!.getAttribute("fill")?.toLowerCase()).toBe("#ff0000");

    const tailId = markerEnd!.slice(5, -1);
    const tail = markers.find((m) => m.getAttribute("id") === tailId)!;
    expect(tail.querySelector("polygon")!.getAttribute("points")).toBe("10,5 0,0 0,10");
  });

  it("renders only marker-end when just tailEnd is set", async () => {
    const element = await renderConnector(
      `${XFRM}
<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>
<a:tailEnd type="triangle"/>
</a:ln>`,
    );
    const path = mainPath(element);
    expect(path.getAttribute("marker-start")).toBeNull();
    expect(path.getAttribute("marker-end")).toMatch(/^url\(#arrow-marker-/);
    expect(element.querySelectorAll("svg defs marker")).toHaveLength(1);
  });

  it("renders stealth and oval line ends with their marker geometry", async () => {
    const element = await renderConnector(
      `${XFRM}
<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
<a:ln w="12700"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>
<a:headEnd type="stealth"/><a:tailEnd type="oval"/>
</a:ln>`,
    );
    const markers = [...element.querySelectorAll("svg defs marker")];
    expect(markers).toHaveLength(2);
    const stealth = markers.find((m) => m.querySelector("path"));
    const oval = markers.find((m) => m.querySelector("circle"));
    expect(stealth).toBeDefined();
    expect(oval).toBeDefined();
    expect(oval!.querySelector("circle")!.getAttribute("fill")?.toLowerCase()).toBe("#0000ff");
  });
});
