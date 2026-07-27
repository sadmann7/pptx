/**
 * Structural comparison specs (port of upstream's Layer 1).
 *
 * Instead of comparing pixels, these assert the parsed/serialized presentation
 * structure against what the fixture generator is known to produce
 * (scripts/generate-fixtures.ts). Catches parser regressions (dropped nodes,
 * wrong transforms, lost text) independently of rendering.
 *
 * Units: fixtures are authored in EMU; the model converts at 96px/inch
 * (914400 EMU = 96px). The 12192000 x 6858000 slide size becomes 1280 x 720.
 */
import { expect, test } from "@playwright/test";

import { getStructure, openSlide } from "./utils";

test.describe("basic deck structure", () => {
  test("parses dimensions, slides, and shape transforms", async ({ page }) => {
    await openSlide(page, "basic.pptx");
    const structure = await getStructure(page);

    expect(structure.width).toBe(1280);
    expect(structure.height).toBe(720);
    expect(structure.slideCount).toBe(3);

    const texts = structure.slides.map((slide) => slide.nodes[0]?.textBody?.totalText);
    expect(texts).toEqual(["Slide one", "Slide two", "Slide three"]);

    for (const slide of structure.slides) {
      expect(slide.nodes).toHaveLength(1);
      const rect = slide.nodes[0];
      expect(rect.nodeType).toBe("shape");
      expect(rect.presetGeometry).toBe("rect");
      // off 914400,914400 ext 6096000x1828800 EMU
      expect(rect.position.x).toBeCloseTo(96, 3);
      expect(rect.position.y).toBeCloseTo(96, 3);
      expect(rect.size.w).toBeCloseTo(640, 3);
      expect(rect.size.h).toBeCloseTo(192, 3);
      expect(rect.rotation).toBe(0);
      expect(rect.flipH).toBe(false);
      expect(rect.flipV).toBe(false);
    }
  });
});

test.describe("BOM deck structure", () => {
  // Structural proof the BOM'd relationship parts were parsed: if DOMParser
  // rejected them, the deck would have no slides or no nodes at all.
  test("parses slides and text through BOM-prefixed rels", async ({ page }) => {
    await openSlide(page, "bom-rels.pptx");
    const structure = await getStructure(page);

    expect(structure.slideCount).toBe(1);
    expect(structure.slides[0].nodes).toHaveLength(1);
    expect(structure.slides[0].nodes[0].textBody?.totalText).toBe("BOM deck renders");
  });
});

test.describe("tables and groups deck structure", () => {
  // Regression: table columns were sized from the frame ext instead of the
  // per-column w values in tblGrid.
  test("parses table grid column widths and cell text", async ({ page }) => {
    await openSlide(page, "tables-groups.pptx", 0);
    const structure = await getStructure(page);

    const [table] = structure.slides[0].nodes;
    expect(table.nodeType).toBe("table");
    // tblGrid: 4572000/3048000/1524000 EMU, not the even 320/320/320 split
    // a frame-ext-based fallback would produce.
    expect(table.columns).toEqual([480, 320, 160]);

    expect(table.rows).toHaveLength(2);
    expect(table.rows?.[0].cells.map((cell) => cell.text)).toEqual(["Stage", "Matches", "Days"]);
    expect(table.rows?.[1].cells.map((cell) => cell.text)).toEqual(["Group", "72", "16"]);
  });

  test("parses group children through chOff/chExt coordinate space", async ({ page }) => {
    await openSlide(page, "tables-groups.pptx", 1);
    const structure = await getStructure(page);

    const [group] = structure.slides[1].nodes;
    expect(group.nodeType).toBe("group");
    expect(group.position).toEqual({ x: 96, y: 96 });
    expect(group.size).toEqual({ w: 384, h: 192 });

    // Children keep child-space coordinates in the model; the 2x chExt scale
    // is applied at render time.
    expect(group.children).toHaveLength(2);
    const [first, second] = group.children ?? [];
    expect(first.nodeType).toBe("shape");
    expect(first.position).toEqual({ x: 0, y: 0 });
    expect(first.size).toEqual({ w: 96, h: 96 });
    expect(second.position).toEqual({ x: 96, y: 0 });
  });
});

test.describe("nested charts deck structure", () => {
  test("resolves chart nodes to their nested part paths", async ({ page }) => {
    await openSlide(page, "nested-charts.pptx");
    const structure = await getStructure(page);

    expect(structure.slideCount).toBe(2);

    const chartPaths = structure.slides.map((slide) => {
      expect(slide.nodes).toHaveLength(1);
      const chart = slide.nodes[0];
      expect(chart.nodeType).toBe("chart");
      // xfrm: off 914400,914400 ext 9144000x4572000 EMU
      expect(chart.position.x).toBeCloseTo(96, 3);
      expect(chart.position.y).toBeCloseTo(96, 3);
      expect(chart.size.w).toBeCloseTo(960, 3);
      expect(chart.size.h).toBeCloseTo(480, 3);
      return chart.chartPath;
    });

    expect(chartPaths).toEqual(["ppt/slides/charts/chart1.xml", "ppt/slides/charts/chart2.xml"]);
  });
});
