import { describe, expect, it } from "vitest";

import { buildPresentation } from "../model/presentation";
import { parseZip } from "../parser/zip-parser";
import { renderSlide } from "../renderer/slide-renderer";
import { buildPptxWithShapes } from "./minimal-pptx";

function tableCell(text: string): string {
  return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`;
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

  const files = await parseZip(buffer);
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

  it("does not clip rows that grow beyond declared heights", async () => {
    const wrapper = await renderTableFrame(
      `<a:ext cx="3000000" cy="3000000"/>`,
      `<a:gridCol w="4572000"/>`,
      `<a:tr h="9525">${tableCell("tall content that needs more than 1px of height")}</a:tr>`,
    );

    expect(wrapper.style.overflow).toBe("visible");
  });
});
