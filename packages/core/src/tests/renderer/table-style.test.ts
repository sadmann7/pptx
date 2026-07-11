import { describe, expect, it } from "vitest";

import { buildPresentation } from "../../model/presentation";
import { readPptx } from "../../ooxml/zip";
import { renderSlide } from "../../renderer/slide";
import {
  getAllPredefinedStyleIds,
  getPredefinedTableStyle,
  PREDEFINED_STYLE_COUNT,
} from "../../renderer/table-style";
import { buildCustomPptx } from "../fixtures/fixture-extras";

const MEDIUM_STYLE_2_ACCENT1 = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
const THEMED_STYLE_2_ACCENT1 = "{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}";

describe("getPredefinedTableStyle", () => {
  it("returns a parsed a:tblStyle node for a known GUID", () => {
    const style = getPredefinedTableStyle(MEDIUM_STYLE_2_ACCENT1);
    expect(style).toBeDefined();
    expect(style!.exists()).toBe(true);
    expect(style!.attr("styleId")).toBe(MEDIUM_STYLE_2_ACCENT1);
    expect(style!.attr("styleName")).toBe("Medium-Style-2");
    expect(style!.child("wholeTbl").exists()).toBe(true);
    expect(style!.child("firstRow").exists()).toBe(true);
    expect(style!.child("band1H").exists()).toBe(true);
  });

  it("returns undefined for unknown GUIDs", () => {
    expect(getPredefinedTableStyle("{00000000-0000-0000-0000-000000000000}")).toBeUndefined();
    expect(getPredefinedTableStyle("")).toBeUndefined();
    expect(getPredefinedTableStyle("not-a-guid")).toBeUndefined();
  });

  it("caches the parsed node per GUID", () => {
    const first = getPredefinedTableStyle(MEDIUM_STYLE_2_ACCENT1);
    const second = getPredefinedTableStyle(MEDIUM_STYLE_2_ACCENT1);
    expect(second).toBe(first);
  });

  it("generates a valid style for every one of the 74 known GUIDs", () => {
    const ids = getAllPredefinedStyleIds();
    expect(ids).toHaveLength(74);
    expect(PREDEFINED_STYLE_COUNT).toBe(74);
    for (const id of ids) {
      const style = getPredefinedTableStyle(id);
      expect(style, `style ${id} should parse`).toBeDefined();
      expect(style!.attr("styleId")).toBe(id);
      // Every generated style has at least the banding sections.
      expect(style!.child("band1H").exists()).toBe(true);
    }
  });

  it("includes the accent tblBg for Themed-Style-2 accents", () => {
    const style = getPredefinedTableStyle(THEMED_STYLE_2_ACCENT1);
    expect(style!.child("tblBg").exists()).toBe(true);
    expect(style!.child("tblBg").child("fillRef").attr("idx")).toBe("1");
  });
});

// ---------------------------------------------------------------------------
// Full pipeline: built-in style referenced by GUID from a slide table
// ---------------------------------------------------------------------------

function tableCell(text: string): string {
  return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`;
}

function styledTableFrame(styleId: string): string {
  return `<p:graphicFrame>
<p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="914400" y="914400"/><a:ext cx="9144000" cy="1371600"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
<a:tbl>
<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>${styleId}</a:tableStyleId></a:tblPr>
<a:tblGrid><a:gridCol w="4572000"/><a:gridCol w="4572000"/></a:tblGrid>
<a:tr h="457200">${tableCell("Header A")}${tableCell("Header B")}</a:tr>
<a:tr h="457200">${tableCell("Band 1")}${tableCell("Band 1b")}</a:tr>
<a:tr h="457200">${tableCell("Plain 2")}${tableCell("Plain 2b")}</a:tr>
</a:tbl>
</a:graphicData></a:graphic>
</p:graphicFrame>`;
}

/**
 * Empty tblStyleLst part. Real PowerPoint packages always ship one; the
 * predefined-style fallback is only reachable when the part exists but does
 * not define the referenced GUID (see the TODO(spec?) test below).
 */
const EMPTY_TABLE_STYLES = `<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="${MEDIUM_STYLE_2_ACCENT1}"/>`;

/** Render a 3x2 table styled by a built-in GUID and return the rendered rows. */
async function renderStyledTable(
  styleId: string,
  options: { includeTableStylesPart?: boolean } = {},
): Promise<{
  table: HTMLTableElement;
  rows: HTMLTableRowElement[];
}> {
  const buffer = await buildCustomPptx({
    slides: [styledTableFrame(styleId)],
    extraFiles:
      options.includeTableStylesPart === false ? {} : { "ppt/tableStyles.xml": EMPTY_TABLE_STYLES },
  });

  const files = await readPptx(buffer);
  const presentation = buildPresentation(files);
  const handle = renderSlide(presentation, presentation.slides[0]);
  const table = handle.element.querySelector("table");
  if (!table) throw new Error("table not found");
  return { table, rows: [...table.querySelectorAll("tr")] };
}

describe("built-in table styles through the render pipeline", () => {
  it("styles the header row with the Medium-Style-2 accent1 fill and light text", async () => {
    const { rows } = await renderStyledTable(MEDIUM_STYLE_2_ACCENT1);
    expect(rows).toHaveLength(3);

    const headerCells = [...rows[0].querySelectorAll("td")];
    expect(headerCells).toHaveLength(2);
    for (const td of headerCells) {
      // firstRow: fill = accent1 (theme: 4472C4)
      expect(td.style.backgroundColor.toUpperCase()).toBe("#4472C4");
    }

    // firstRow tcTxStyle: b="on", color lt1 (theme: FFFFFF) applied to run spans.
    const headerSpans = [...rows[0].querySelectorAll("span")].filter((s) =>
      s.textContent?.includes("Header"),
    );
    expect(headerSpans.length).toBeGreaterThan(0);
    for (const span of headerSpans) {
      expect(span.style.color.toUpperCase()).toBe("#FFFFFF");
      expect(span.style.fontWeight).toBe("bold");
    }
  });

  it("applies banding fills that differ between banded and plain rows", async () => {
    const { rows } = await renderStyledTable(MEDIUM_STYLE_2_ACCENT1);

    const headerBg = rows[0].querySelector("td")!.style.backgroundColor;
    const band1Bg = rows[1].querySelector("td")!.style.backgroundColor; // band1H: accent1 tint 40%
    const plainBg = rows[2].querySelector("td")!.style.backgroundColor; // wholeTbl: accent1 tint 20%

    expect(band1Bg).not.toBe("");
    expect(plainBg).not.toBe("");
    expect(band1Bg).not.toBe(headerBg);
    expect(band1Bg).not.toBe(plainBg);
  });

  it("applies the Themed-Style-2 table background via its theme fillRef", async () => {
    const { table } = await renderStyledTable(THEMED_STYLE_2_ACCENT1);
    // tblBg fillRef idx=1 resolves through the theme fillStyleLst (phClr → accent1).
    const bg = `${table.style.backgroundColor} ${table.style.background}`.toUpperCase();
    expect(bg).toContain("#4472C4");
  });

  it("leaves cells unstyled for an unknown style GUID", async () => {
    const { rows } = await renderStyledTable("{00000000-0000-0000-0000-000000000000}");
    const td = rows[0].querySelector("td")!;
    expect(td.style.backgroundColor).toBe("");
  });

  // TODO(spec?): findTableStyle bails out when the package has no
  // ppt/tableStyles.xml part at all (`!ctx.presentation.tableStyles`), so the
  // predefined built-in fallback is unreachable for such packages even though
  // the slide references a known built-in GUID. Real PowerPoint files always
  // ship a tableStyles part, but a minimal producer that omits it silently
  // loses built-in styling. Pinning the current behavior here.
  it("does not apply built-in styles when the tableStyles part is missing entirely", async () => {
    const { rows } = await renderStyledTable(MEDIUM_STYLE_2_ACCENT1, {
      includeTableStylesPart: false,
    });
    const td = rows[0].querySelector("td")!;
    expect(td.style.backgroundColor).toBe("");
  });
});
