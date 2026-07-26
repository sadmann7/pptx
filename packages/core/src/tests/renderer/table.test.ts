import { describe, expect, it } from "vitest";

import { buildPresentation } from "../../model/presentation";
import { readPptx } from "../../ooxml/zip";
import { renderSlide } from "../../renderer/slide";
import { buildPptxWithShapes } from "../fixtures/minimal-pptx";

function tableCell(text: string, tcPrXml = "<a:tcPr/>"): string {
  return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></a:txBody>${tcPrXml}</a:tc>`;
}

/** A cell with a solid fill and a 1pt border on all four sides. */
function borderedCell(text: string, fillHex: string, lineHex = "112233"): string {
  const ln = (name: string) =>
    `<a:${name} w="12700"><a:solidFill><a:srgbClr val="${lineHex}"/></a:solidFill></a:${name}>`;
  return tableCell(
    text,
    `<a:tcPr>${ln("lnL")}${ln("lnR")}${ln("lnT")}${ln("lnB")}<a:solidFill><a:srgbClr val="${fillHex}"/></a:solidFill></a:tcPr>`,
  );
}

/** Renders a slide with one graphicFrame table and returns the table's wrapper div. */
async function renderTableFrame(
  extXml: string,
  gridColsXml: string,
  rowsXml: string,
): Promise<HTMLElement> {
  const buffer = await buildPptxWithShapes(`<p:graphicFrame>
<p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="914400" y="914400"/>${extXml}</p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
<a:tbl>
<a:tblPr firstRow="1" bandRow="1"/>
<a:tblGrid>${gridColsXml}</a:tblGrid>
${rowsXml}
</a:tbl>
</a:graphicData></a:graphic>
</p:graphicFrame>`);

  const files = await readPptx(buffer);
  const presentation = buildPresentation(files);
  const handle = renderSlide(presentation, presentation.slides[0]);
  const table = handle.element.querySelector("table");
  if (!table?.parentElement) throw new Error("table wrapper not found");
  return table.parentElement;
}

describe("table sizing", () => {
  it("sizes the table from the grid, ignoring a stale graphicFrame extent", async () => {
    // Regression: Google Slides exports write a dummy 3000000x3000000 extent
    // on table graphicFrames; sizing from it crushed tables into a small
    // square. The grid says 2x4572000 EMU wide (960px), 2x457200 EMU tall (96px).
    const wrapper = await renderTableFrame(
      `<a:ext cx="3000000" cy="3000000"/>`,
      `<a:gridCol w="4572000"/><a:gridCol w="4572000"/>`,
      `<a:tr h="457200">${tableCell("A")}${tableCell("B")}</a:tr>
       <a:tr h="457200">${tableCell("C")}${tableCell("D")}</a:tr>`,
    );

    expect(wrapper.style.width).toBe("960px");
    expect(wrapper.style.height).toBe("96px");
  });

  it("resolves the node size from the grid so overlays match the render", async () => {
    const buffer = await buildPptxWithShapes(`<p:graphicFrame>
<p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3000000" cy="3000000"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
<a:tbl>
<a:tblPr/>
<a:tblGrid><a:gridCol w="4572000"/><a:gridCol w="4572000"/></a:tblGrid>
<a:tr h="457200">${tableCell("A")}${tableCell("B")}</a:tr>
<a:tr h="457200">${tableCell("C")}${tableCell("D")}</a:tr>
</a:tbl>
</a:graphicData></a:graphic>
</p:graphicFrame>`);

    const presentation = buildPresentation(await readPptx(buffer));
    const node = presentation.slides[0].nodes[0];
    expect(node.nodeType).toBe("table");
    expect(node.size).toEqual({ w: 960, h: 96 });
  });

  it("keeps proportional column widths from the grid", async () => {
    const wrapper = await renderTableFrame(
      `<a:ext cx="3000000" cy="3000000"/>`,
      `<a:gridCol w="2286000"/><a:gridCol w="6858000"/>`,
      `<a:tr h="457200">${tableCell("A")}${tableCell("B")}</a:tr>`,
    );

    const cols = [...wrapper.querySelectorAll("col")];
    expect(cols).toHaveLength(2);
    expect(cols[0].style.width).toBe("25%");
    expect(cols[1].style.width).toBe("75%");
  });

  it("falls back to the frame extent when the grid declares no widths", async () => {
    // 3810000 EMU = 400px, 1905000 EMU = 200px.
    const wrapper = await renderTableFrame(
      `<a:ext cx="3810000" cy="1905000"/>`,
      `<a:gridCol/><a:gridCol/>`,
      `<a:tr>${tableCell("A")}${tableCell("B")}</a:tr>`,
    );

    expect(wrapper.style.width).toBe("400px");
    expect(wrapper.style.height).toBe("200px");
  });

  it("paints each shared edge once so borders never straddle two fills", async () => {
    // Collapsed borders are composited against whatever sits behind the table,
    // so at fractional zoom a shared edge picks up the slide background and
    // reads as a different colour depending on the fills it separates.
    const wrapper = await renderTableFrame(
      `<a:ext cx="3000000" cy="3000000"/>`,
      `<a:gridCol w="4572000"/><a:gridCol w="4572000"/>`,
      `<a:tr h="457200">${borderedCell("A", "0F0F0F")}${borderedCell("B", "0F0F0F")}</a:tr>
       <a:tr h="457200">${borderedCell("C", "EFE9D9")}${borderedCell("D", "EFE9D9")}</a:tr>`,
    );

    const table = wrapper.querySelector("table")!;
    expect(table.style.borderCollapse).toBe("separate");
    expect(table.style.borderSpacing).toMatch(/^0(px)?$/);

    const [a, b, c, d] = [...wrapper.querySelectorAll("td")];
    // The upper/left cell owns each interior edge.
    expect(a.style.borderRight).toContain("#112233");
    expect(a.style.borderBottom).toContain("#112233");
    expect(b.style.borderLeft).toBe("");
    expect(c.style.borderTop).toBe("");
    expect(d.style.borderTop).toBe("");
    expect(d.style.borderLeft).toBe("");
  });

  it("draws a uniform outline on the table instead of on the boundary cells", async () => {
    const wrapper = await renderTableFrame(
      `<a:ext cx="3000000" cy="3000000"/>`,
      `<a:gridCol w="4572000"/><a:gridCol w="4572000"/>`,
      `<a:tr h="457200">${borderedCell("A", "0F0F0F")}${borderedCell("B", "0F0F0F")}</a:tr>
       <a:tr h="457200">${borderedCell("C", "EFE9D9")}${borderedCell("D", "EFE9D9")}</a:tr>`,
    );

    const table = wrapper.querySelector("table")!;
    for (const side of ["borderTop", "borderBottom", "borderLeft", "borderRight"] as const) {
      expect(table.style[side]).toContain("#112233");
    }
    expect(table.style.boxSizing).toBe("border-box");

    const [a, , , d] = [...wrapper.querySelectorAll("td")];
    expect(a.style.borderTop).toBe("");
    expect(a.style.borderLeft).toBe("");
    expect(d.style.borderRight).toBe("");
    expect(d.style.borderBottom).toBe("");
  });

  it("leaves an outline edge on its cells when they disagree", async () => {
    const wrapper = await renderTableFrame(
      `<a:ext cx="3000000" cy="3000000"/>`,
      `<a:gridCol w="4572000"/><a:gridCol w="4572000"/>`,
      `<a:tr h="457200">${borderedCell("A", "0F0F0F", "112233")}${borderedCell("B", "0F0F0F", "445566")}</a:tr>`,
    );

    const table = wrapper.querySelector("table")!;
    expect(table.style.borderTop).toBe("");

    const [a, b] = [...wrapper.querySelectorAll("td")];
    expect(a.style.borderTop).toContain("#112233");
    expect(b.style.borderTop).toContain("#445566");
  });

  it("backs a uniformly filled row with its own colour, and a mixed row with none", async () => {
    const wrapper = await renderTableFrame(
      `<a:ext cx="3000000" cy="3000000"/>`,
      `<a:gridCol w="4572000"/><a:gridCol w="4572000"/>`,
      `<a:tr h="457200">${borderedCell("A", "0F0F0F")}${borderedCell("B", "0F0F0F")}</a:tr>
       <a:tr h="457200">${borderedCell("C", "EFE9D9")}${borderedCell("D", "0F0F0F")}</a:tr>`,
    );

    const [uniform, mixed] = [...wrapper.querySelectorAll("tr")];
    expect(uniform.style.backgroundColor.toUpperCase()).toBe("#0F0F0F");
    expect(mixed.style.backgroundColor).toBe("");
  });

  it("does not back a row whose cells leave the grid partly uncovered", async () => {
    // The open band belongs to a vertically merged cell; tinting it would show
    // through that cell when it has no fill of its own.
    const wrapper = await renderTableFrame(
      `<a:ext cx="3000000" cy="3000000"/>`,
      `<a:gridCol w="4572000"/><a:gridCol w="4572000"/>`,
      `<a:tr h="457200"><a:tc rowSpan="2">${""}<a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc>${borderedCell("B", "0F0F0F")}</a:tr>
       <a:tr h="457200"><a:tc vMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc>${borderedCell("D", "0F0F0F")}</a:tr>`,
    );

    const rows = [...wrapper.querySelectorAll("tr")];
    expect(rows[1].style.backgroundColor).toBe("");
  });

  it("does not clip rows that grow beyond declared heights", async () => {
    const wrapper = await renderTableFrame(
      `<a:ext cx="3000000" cy="3000000"/>`,
      `<a:gridCol w="4572000"/>`,
      `<a:tr h="9525">${tableCell("tall content that needs more than 1px of height")}</a:tr>`,
    );

    expect(wrapper.style.overflow).toBe("visible");
  });
});
