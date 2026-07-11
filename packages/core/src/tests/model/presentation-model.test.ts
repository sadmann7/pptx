import { beforeAll, describe, expect, it } from "vitest";

import { buildPresentation, materializeAllSlides } from "../../model/presentation";
import type { PptxFiles } from "../../ooxml/zip";
import { readPptx, RECOMMENDED_PPTX_READ_LIMITS } from "../../ooxml/zip";
import { buildPptxWithShapes } from "../fixtures/minimal-pptx";

const SHAPE = `<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Blue Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="914400" y="457200"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom>
<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
</p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Quarterly revenue</a:t></a:r></a:p></p:txBody>
</p:sp>`;

let buffer: ArrayBuffer;
let files: PptxFiles;

beforeAll(async () => {
  buffer = await buildPptxWithShapes(SHAPE);
  files = await readPptx(buffer);
});

describe("readPptx", () => {
  it("extracts all package parts", () => {
    expect(files.presentation).toContain("<p:presentation");
    expect(files.slides.size).toBe(1);
    expect(files.slideLayouts.size).toBe(1);
    expect([...files.slides.keys()][0]).toBe("ppt/slides/slide1.xml");
  });

  it("rejects non-zip input", async () => {
    await expect(readPptx(new ArrayBuffer(32))).rejects.toThrow();
  });

  it("enforces entry-count limits", async () => {
    await expect(readPptx(buffer, { maxEntries: 2 })).rejects.toThrow(/entries|limit/i);
  });

  it("enforces uncompressed-size limits", async () => {
    await expect(readPptx(buffer, { maxEntryUncompressedBytes: 16 })).rejects.toThrow();
    // Recommended limits comfortably admit the fixture.
    await expect(readPptx(buffer, RECOMMENDED_PPTX_READ_LIMITS)).resolves.toBeDefined();
  });
});

describe("buildPresentation", () => {
  it("resolves slide size in pixels from sldSz", () => {
    const pres = buildPresentation(files);
    expect(pres.width).toBeCloseTo(1280, 6);
    expect(pres.height).toBeCloseTo(720, 6);
  });

  it("links slide → layout → master → theme", () => {
    const pres = buildPresentation(files);
    expect(pres.slides).toHaveLength(1);
    const layoutPath = pres.slideToLayout.get(pres.slides[0].index);
    expect(layoutPath).toBe("ppt/slideLayouts/slideLayout1.xml");
    const masterPath = pres.layoutToMaster.get(layoutPath!);
    expect(masterPath).toBe("ppt/slideMasters/slideMaster1.xml");
    const themePath = pres.masterToTheme.get(masterPath!);
    expect(themePath).toBe("ppt/theme/theme1.xml");
    expect(pres.themes.get(themePath!)?.colorScheme.get("accent1")).toBe("4472C4");
  });

  it("materializes slide nodes with parsed geometry and adjustments", () => {
    const pres = buildPresentation(files);
    materializeAllSlides(pres);
    const nodes = pres.slides[0].nodes;
    expect(nodes.length).toBeGreaterThan(0);
    const shape = nodes.find((n) => n.nodeType === "shape");
    expect(shape).toBeDefined();
    if (shape?.nodeType !== "shape") throw new Error("unreachable");
    expect(shape.presetGeometry).toBe("roundRect");
    expect(shape.adjustments.get("adj")).toBe(25000);
    // 914400 EMU = 96px, 457200 EMU = 48px
    expect(shape.position).toEqual({ x: 96, y: 48 });
    expect(shape.size).toEqual({ w: 192, h: 96 });
  });
});
