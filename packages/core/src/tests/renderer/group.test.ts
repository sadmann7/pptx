/**
 * Group rendering: child coordinate-space remapping, nested groups,
 * group flips/rotation, grpFill inheritance, and group-level effects.
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

/** Find the wrapper div of the shape whose SVG path has the given fill. */
function shapeWrapperByFill(element: HTMLElement, hexLower: string): HTMLElement {
  const paths = [...element.querySelectorAll("svg path")];
  const match = paths.find((p) => p.getAttribute("fill")?.toLowerCase() === hexLower);
  if (!match) throw new Error(`no path with fill ${hexLower}`);
  const wrapper = match.closest("div");
  if (!wrapper) throw new Error("shape has no wrapper div");
  return wrapper as HTMLElement;
}

function simpleRect(id: number, offX: number, offY: number, cx: number, cy: number, hex: string) {
  return `<p:sp>
<p:nvSpPr><p:cNvPr id="${id}" name="Rect${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="${offX}" y="${offY}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>
</p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>`;
}

function group(inner: string, xfrmAttrs = "", grpSpPrExtra = ""): string {
  return `<p:grpSp>
<p:nvGrpSpPr><p:cNvPr id="10" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr>
<a:xfrm${xfrmAttrs}>
<a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/>
<a:chOff x="0" y="0"/><a:chExt cx="914400" cy="457200"/>
</a:xfrm>
${grpSpPrExtra}
</p:grpSpPr>
${inner}
</p:grpSp>`;
}

describe("group child coordinate remapping", () => {
  it("scales and positions children when chExt differs from ext", async () => {
    // Group: 96,96 px offset, 192x96 px box; child space 96x48 -> scale 2x/2x.
    // Child at (48,0) 48x48 in child space -> (96,0) 192x96-local, 96x96 px.
    const element = await renderShapes(group(simpleRect(11, 457200, 0, 457200, 457200, "ED7D31")));

    const shapeWrapper = shapeWrapperByFill(element, "#ed7d31");
    expect(shapeWrapper.style.left).toBe("96px");
    expect(shapeWrapper.style.top).toBe("0px");
    expect(shapeWrapper.style.width).toBe("96px");
    expect(shapeWrapper.style.height).toBe("96px");

    const groupWrapper = shapeWrapper.parentElement as HTMLElement;
    expect(groupWrapper.style.position).toBe("absolute");
    expect(groupWrapper.style.left).toBe("96px");
    expect(groupWrapper.style.top).toBe("96px");
    expect(groupWrapper.style.width).toBe("192px");
    expect(groupWrapper.style.height).toBe("96px");
  });

  it("scales children of nested groups through both coordinate spaces", async () => {
    // Outer group 192x96 with child space 96x48 (scale 2).
    // Inner group occupies (0,0) 48x48 of child space -> rendered at (0,0) 96x96.
    // Inner child space is 48x48 -> inner scale 2; innermost rect (24,0) 24x24
    // -> (48,0) 48x48 inside the inner group wrapper.
    const innerGroup = `<p:grpSp>
<p:nvGrpSpPr><p:cNvPr id="20" name="Inner Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr>
<a:xfrm>
<a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/>
<a:chOff x="0" y="0"/><a:chExt cx="457200" cy="457200"/>
</a:xfrm>
</p:grpSpPr>
${simpleRect(21, 228600, 0, 228600, 228600, "00AA00")}
</p:grpSp>`;

    const element = await renderShapes(group(innerGroup));

    const shapeWrapper = shapeWrapperByFill(element, "#00aa00");
    expect(shapeWrapper.style.left).toBe("48px");
    expect(shapeWrapper.style.top).toBe("0px");
    expect(shapeWrapper.style.width).toBe("48px");
    expect(shapeWrapper.style.height).toBe("48px");

    const innerGroupWrapper = shapeWrapper.parentElement as HTMLElement;
    expect(innerGroupWrapper.style.left).toBe("0px");
    expect(innerGroupWrapper.style.width).toBe("96px");
    expect(innerGroupWrapper.style.height).toBe("96px");

    const outerGroupWrapper = innerGroupWrapper.parentElement as HTMLElement;
    expect(outerGroupWrapper.style.left).toBe("96px");
    expect(outerGroupWrapper.style.top).toBe("96px");
    expect(outerGroupWrapper.style.width).toBe("192px");
  });
});

describe("group flips and rotation", () => {
  it("mirrors child positions and toggles child flipH for flipH groups", async () => {
    // Child at (0,0) 48x48 in child space -> (0,0) 96x96 after scaling.
    // flipH group: x -> groupW - x - w = 192 - 0 - 96 = 96.
    const element = await renderShapes(
      group(simpleRect(11, 0, 0, 457200, 457200, "ED7D31"), ` flipH="1"`),
    );

    const shapeWrapper = shapeWrapperByFill(element, "#ed7d31");
    expect(shapeWrapper.style.left).toBe("96px");
    // The flip is propagated to the child instead of mirroring the group DOM.
    expect(shapeWrapper.style.transform).toContain("scaleX(-1)");
    const groupWrapper = shapeWrapper.parentElement as HTMLElement;
    expect(groupWrapper.style.transform).not.toContain("scaleX");
  });

  it("mirrors child positions vertically for flipV groups", async () => {
    // Child at (0,0) 48x24 in child space -> (0,0) 96x48; flipV: y -> 96-0-48 = 48.
    const element = await renderShapes(
      group(simpleRect(11, 0, 0, 457200, 228600, "ED7D31"), ` flipV="1"`),
    );

    const shapeWrapper = shapeWrapperByFill(element, "#ed7d31");
    expect(shapeWrapper.style.top).toBe("48px");
    expect(shapeWrapper.style.transform).toContain("scaleY(-1)");
  });

  it("applies group rotation on the group wrapper", async () => {
    const element = await renderShapes(
      group(simpleRect(11, 0, 0, 457200, 457200, "ED7D31"), ` rot="2700000"`),
    );

    const groupWrapper = shapeWrapperByFill(element, "#ed7d31").parentElement as HTMLElement;
    expect(groupWrapper.style.transform).toContain("rotate(45deg)");
    expect(groupWrapper.style.transformOrigin).toBe("center center");
  });
});

describe("group fill inheritance (a:grpFill)", () => {
  it("children with a:grpFill inherit the group's solid fill", async () => {
    const child = `<p:sp>
<p:nvSpPr><p:cNvPr id="11" name="Inheriting"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:grpFill/>
</p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>`;
    const element = await renderShapes(
      group(child, "", `<a:solidFill><a:srgbClr val="00FF77"/></a:solidFill>`),
    );

    const paths = [...element.querySelectorAll("svg path")];
    const inherited = paths.find((p) => p.getAttribute("fill")?.toLowerCase() === "#00ff77");
    expect(inherited).toBeDefined();
  });
});

describe("group effects", () => {
  it("applies grpSpPr outerShdw as a drop-shadow filter on the group wrapper", async () => {
    const effect = `<a:effectLst>
<a:outerShdw blurRad="19050" dist="38100" dir="0">
<a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr>
</a:outerShdw>
</a:effectLst>`;
    const element = await renderShapes(
      group(simpleRect(11, 0, 0, 457200, 457200, "ED7D31"), "", effect),
    );

    const groupWrapper = shapeWrapperByFill(element, "#ed7d31").parentElement as HTMLElement;
    // dist 38100 EMU = 4px at dir 0; blurRad 19050 EMU = 2px.
    expect(groupWrapper.style.filter).toContain("drop-shadow(4.0px 0.0px 2.0px");
    expect(groupWrapper.style.filter).toContain("rgba(0,0,0,0.500)");
  });
});
