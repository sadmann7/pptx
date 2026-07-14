import { describe, expect, it } from "vitest";

import { buildPresentation } from "../../model/presentation";
import { readPptx, RECOMMENDED_PPTX_READ_LIMITS } from "../../ooxml/zip";
import { renderSlide } from "../../renderer/slide";
import { buildCustomPptx, fakePngBytes } from "../fixtures/fixture-extras";

const PNG = fakePngBytes(64);

function pptxWithMedia() {
  return buildCustomPptx({
    extraFiles: {
      "ppt/media/image1.png": PNG,
      "ppt/media/image2.png": fakePngBytes(32),
    },
  });
}

describe("readPptx media and auxiliary parts", () => {
  it("eagerly decodes media bytes into the media map", async () => {
    const files = await readPptx(await pptxWithMedia());
    expect(files.media.size).toBe(2);
    expect(files.media.get("ppt/media/image1.png")).toEqual(PNG);
    expect(files.mediaResolver).toBeUndefined();
  });

  it("adds a decoded alias for percent-encoded media entry names", async () => {
    const buffer = await buildCustomPptx({
      extraFiles: { "ppt/media/my%20image.png": PNG },
    });
    const files = await readPptx(buffer);
    expect(files.media.get("ppt/media/my%20image.png")).toEqual(PNG);
    expect(files.media.get("ppt/media/my image.png")).toEqual(PNG);
  });

  it("extracts ppt/tableStyles.xml into files.tableStyles", async () => {
    const tableStylesXml = `<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{X}"><a:tblStyle styleId="{X}" styleName="Custom"/></a:tblStyleLst>`;
    const files = await readPptx(
      await buildCustomPptx({ extraFiles: { "ppt/tableStyles.xml": tableStylesXml } }),
    );
    expect(files.tableStyles).toBe(tableStylesXml);
  });

  it("extracts embedded font parts under ppt/fonts/*.fntdata", async () => {
    const fontBytes = new Uint8Array([0x00, 0x01, 0x00, 0x00, 1, 2, 3]);
    const files = await readPptx(
      await buildCustomPptx({ extraFiles: { "ppt/fonts/font1.fntdata": fontBytes } }),
    );
    expect(files.fonts.get("ppt/fonts/font1.fntdata")).toEqual(fontBytes);
  });

  it("stores masters, layouts, and rels keyed by normalized path", async () => {
    const files = await readPptx(await buildCustomPptx());
    expect([...files.slideMasters.keys()]).toEqual(["ppt/slideMasters/slideMaster1.xml"]);
    expect([...files.slideMasterRels.keys()]).toEqual([
      "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    ]);
    expect([...files.themes.keys()]).toEqual(["ppt/theme/theme1.xml"]);
  });

  it("categorizes chart parts nested outside ppt/charts (e.g. ppt/slides/charts)", async () => {
    const chartXml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>`;
    const relsXml = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;
    const files = await readPptx(
      await buildCustomPptx({
        extraFiles: {
          "ppt/slides/charts/chart1.xml": chartXml,
          "ppt/slides/charts/_rels/chart1.xml.rels": relsXml,
          "ppt/slides/charts/style1.xml": "<cs:chartStyle/>",
          "ppt/slides/charts/colors1.xml": "<cs:colorStyle/>",
        },
      }),
    );
    expect(files.charts.get("ppt/slides/charts/chart1.xml")).toBe(chartXml);
    expect(files.chartRels?.get("ppt/slides/charts/_rels/chart1.xml.rels")).toBe(relsXml);
    expect(files.chartStyles.has("ppt/slides/charts/style1.xml")).toBe(true);
    expect(files.chartColors.has("ppt/slides/charts/colors1.xml")).toBe(true);
    // Slide categorization must not swallow the nested chart parts.
    expect(files.slides.has("ppt/slides/charts/chart1.xml")).toBe(false);
  });

  it("leaves uncategorized parts (e.g. notesSlides) out of the result", async () => {
    const files = await readPptx(
      await buildCustomPptx({
        extraFiles: { "ppt/notesSlides/notesSlide1.xml": "<p:notes/>" },
      }),
    );
    // readPptx has no notesSlides category; the entry is only size-counted.
    expect(files.slides.size).toBe(1);
    expect(Object.keys(files)).not.toContain("notesSlides");
  });

  it("passes malformed [Content_Types].xml through without failing the pipeline", async () => {
    const buffer = await buildCustomPptx({
      slides: [""],
      contentTypesXml: "<<< this is not xml >>>",
    });
    const files = await readPptx(buffer);
    // Categorization is path-based, so a broken content-types part is inert.
    expect(files.contentTypes).toBe("<<< this is not xml >>>");
    const presentation = buildPresentation(files);
    expect(presentation.slides).toHaveLength(1);
    const handle = renderSlide(presentation, presentation.slides[0]);
    expect(handle.element).toBeInstanceOf(HTMLElement);
  });
});

describe("readPptx lazyMedia", () => {
  it("leaves media empty and exposes a resolver with totals", async () => {
    const files = await readPptx(await pptxWithMedia(), { lazyMedia: true });
    expect(files.media.size).toBe(0);
    expect(files.mediaResolver).toBeDefined();
    expect(files.mediaResolver?.totalCount).toBe(2);
    expect(files.mediaResolver?.totalBytes).toBe(64 + 32);
    expect(files.mediaResolver?.loadedCount).toBe(0);
    expect(files.mediaResolver?.loadedBytes).toBe(0);
  });

  it("decodes entries on demand and populates the shared media map", async () => {
    const files = await readPptx(await pptxWithMedia(), { lazyMedia: true });
    const resolved = await files.mediaResolver?.resolve("../media/image1.png");
    expect(resolved?.mediaPath).toBe("ppt/media/image1.png");
    expect(resolved?.data).toEqual(PNG);

    // The resolved entry is now visible to eager lookups too.
    expect(files.media.get("ppt/media/image1.png")).toEqual(PNG);
    expect(files.mediaResolver?.loadedCount).toBe(1);
    expect(files.mediaResolver?.loadedBytes).toBe(64);
  });

  it("serves repeat and concurrent resolves from the cache", async () => {
    const files = await readPptx(await pptxWithMedia(), { lazyMedia: true });
    const [a, b] = await Promise.all([
      files.mediaResolver!.resolve("../media/image1.png"),
      files.mediaResolver!.resolve("../media/image1.png"),
    ]);
    expect(a?.data).toEqual(PNG);
    expect(b?.data).toEqual(PNG);
    expect(files.mediaResolver?.loadedCount).toBe(1);

    const again = await files.mediaResolver!.resolve("../media/image1.png");
    expect(again?.data).toEqual(PNG);
    expect(files.mediaResolver?.loadedCount).toBe(1);
  });

  it("resolves undefined for unknown targets", async () => {
    const files = await readPptx(await pptxWithMedia(), { lazyMedia: true });
    await expect(files.mediaResolver!.resolve("../media/missing.png")).resolves.toBeUndefined();
  });

  it("otherwise categorizes XML parts exactly like eager readPptx", async () => {
    const buffer = await pptxWithMedia();
    const eager = await readPptx(buffer);
    const lazy = await readPptx(buffer, { lazyMedia: true });
    expect([...lazy.slides.keys()]).toEqual([...eager.slides.keys()]);
    expect(lazy.presentation).toBe(eager.presentation);
    expect(lazy.contentTypes).toBe(eager.contentTypes);
  });
});

describe("readPptx limits", () => {
  it("enforces maxMediaBytes in eager mode", async () => {
    const buffer = await buildCustomPptx({
      extraFiles: { "ppt/media/big.png": fakePngBytes(1024) },
    });
    await expect(readPptx(buffer, { limits: { maxMediaBytes: 512 } })).rejects.toThrow(
      /media bytes/,
    );
  });

  it("enforces maxMediaBytes up front in lazy mode via zip directory sizes", async () => {
    const buffer = await buildCustomPptx({
      extraFiles: { "ppt/media/big.png": fakePngBytes(1024) },
    });
    await expect(
      readPptx(buffer, { limits: { maxMediaBytes: 512 }, lazyMedia: true }),
    ).rejects.toThrow(/media bytes/);
  });

  it("counts media across multiple entries against maxMediaBytes", async () => {
    const buffer = await buildCustomPptx({
      extraFiles: {
        "ppt/media/a.png": fakePngBytes(300),
        "ppt/media/b.png": fakePngBytes(300),
      },
    });
    // Each entry fits alone but not together.
    await expect(readPptx(buffer, { limits: { maxMediaBytes: 400 } })).rejects.toThrow(
      /media bytes/,
    );
    await expect(readPptx(buffer, { limits: { maxMediaBytes: 600 } })).resolves.toBeDefined();
  });

  it("admits the fixture under RECOMMENDED_PPTX_READ_LIMITS in both modes", async () => {
    const buffer = await pptxWithMedia();
    await expect(readPptx(buffer, { limits: RECOMMENDED_PPTX_READ_LIMITS })).resolves.toBeDefined();
    const lazy = await readPptx(buffer, { limits: RECOMMENDED_PPTX_READ_LIMITS, lazyMedia: true });
    const resolved = await lazy.mediaResolver?.resolve("../media/image1.png");
    expect(resolved?.data).toEqual(PNG);
  });

  it("rejects invalid maxConcurrency values", async () => {
    const buffer = await buildCustomPptx();
    await expect(readPptx(buffer, { limits: { maxConcurrency: 0 } })).rejects.toThrow(
      /maxConcurrency/,
    );
    await expect(readPptx(buffer, { limits: { maxConcurrency: 1.5 } })).rejects.toThrow(
      /maxConcurrency/,
    );
    await expect(readPptx(buffer, { limits: { maxConcurrency: 1 } })).resolves.toBeDefined();
  });

  it("enforces maxEntryUncompressedBytes against media entries", async () => {
    const buffer = await buildCustomPptx({
      extraFiles: { "ppt/media/big.png": fakePngBytes(2048) },
    });
    await expect(readPptx(buffer, { limits: { maxEntryUncompressedBytes: 1024 } })).rejects.toThrow(
      /maxEntryUncompressedBytes/,
    );
  });
});
