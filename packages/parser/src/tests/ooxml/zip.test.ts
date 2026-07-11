import { describe, expect, it } from "vitest";

import { buildPresentation } from "../../model/presentation";
import { parseZip, parseZipLazyMedia, RECOMMENDED_ZIP_LIMITS } from "../../ooxml/zip";
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

describe("parseZip media and auxiliary parts", () => {
  it("eagerly decodes media bytes into the media map", async () => {
    const files = await parseZip(await pptxWithMedia());
    expect(files.media.size).toBe(2);
    expect(files.media.get("ppt/media/image1.png")).toEqual(PNG);
    expect(files.mediaResolver).toBeUndefined();
  });

  it("adds a decoded alias for percent-encoded media entry names", async () => {
    const buffer = await buildCustomPptx({
      extraFiles: { "ppt/media/my%20image.png": PNG },
    });
    const files = await parseZip(buffer);
    expect(files.media.get("ppt/media/my%20image.png")).toEqual(PNG);
    expect(files.media.get("ppt/media/my image.png")).toEqual(PNG);
  });

  it("extracts ppt/tableStyles.xml into files.tableStyles", async () => {
    const tableStylesXml = `<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{X}"><a:tblStyle styleId="{X}" styleName="Custom"/></a:tblStyleLst>`;
    const files = await parseZip(
      await buildCustomPptx({ extraFiles: { "ppt/tableStyles.xml": tableStylesXml } }),
    );
    expect(files.tableStyles).toBe(tableStylesXml);
  });

  it("extracts embedded font parts under ppt/fonts/*.fntdata", async () => {
    const fontBytes = new Uint8Array([0x00, 0x01, 0x00, 0x00, 1, 2, 3]);
    const files = await parseZip(
      await buildCustomPptx({ extraFiles: { "ppt/fonts/font1.fntdata": fontBytes } }),
    );
    expect(files.fonts.get("ppt/fonts/font1.fntdata")).toEqual(fontBytes);
  });

  it("stores masters, layouts, and rels keyed by normalized path", async () => {
    const files = await parseZip(await buildCustomPptx());
    expect([...files.slideMasters.keys()]).toEqual(["ppt/slideMasters/slideMaster1.xml"]);
    expect([...files.slideMasterRels.keys()]).toEqual([
      "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    ]);
    expect([...files.themes.keys()]).toEqual(["ppt/theme/theme1.xml"]);
  });

  it("leaves uncategorized parts (e.g. notesSlides) out of the result", async () => {
    const files = await parseZip(
      await buildCustomPptx({
        extraFiles: { "ppt/notesSlides/notesSlide1.xml": "<p:notes/>" },
      }),
    );
    // parseZip has no notesSlides category; the entry is only size-counted.
    expect(files.slides.size).toBe(1);
    expect(Object.keys(files)).not.toContain("notesSlides");
  });

  it("passes malformed [Content_Types].xml through without failing the pipeline", async () => {
    const buffer = await buildCustomPptx({
      slides: [""],
      contentTypesXml: "<<< this is not xml >>>",
    });
    const files = await parseZip(buffer);
    // Categorization is path-based, so a broken content-types part is inert.
    expect(files.contentTypes).toBe("<<< this is not xml >>>");
    const presentation = buildPresentation(files);
    expect(presentation.slides).toHaveLength(1);
    const handle = renderSlide(presentation, presentation.slides[0]);
    expect(handle.element).toBeInstanceOf(HTMLElement);
  });
});

describe("parseZipLazyMedia", () => {
  it("leaves media empty and exposes a resolver with totals", async () => {
    const files = await parseZipLazyMedia(await pptxWithMedia());
    expect(files.media.size).toBe(0);
    expect(files.mediaResolver).toBeDefined();
    expect(files.mediaResolver?.totalCount).toBe(2);
    expect(files.mediaResolver?.totalBytes).toBe(64 + 32);
    expect(files.mediaResolver?.loadedCount).toBe(0);
    expect(files.mediaResolver?.loadedBytes).toBe(0);
  });

  it("decodes entries on demand and populates the shared media map", async () => {
    const files = await parseZipLazyMedia(await pptxWithMedia());
    const resolved = await files.mediaResolver?.resolve("../media/image1.png");
    expect(resolved?.mediaPath).toBe("ppt/media/image1.png");
    expect(resolved?.data).toEqual(PNG);

    // The resolved entry is now visible to eager lookups too.
    expect(files.media.get("ppt/media/image1.png")).toEqual(PNG);
    expect(files.mediaResolver?.loadedCount).toBe(1);
    expect(files.mediaResolver?.loadedBytes).toBe(64);
  });

  it("serves repeat and concurrent resolves from the cache", async () => {
    const files = await parseZipLazyMedia(await pptxWithMedia());
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
    const files = await parseZipLazyMedia(await pptxWithMedia());
    await expect(files.mediaResolver!.resolve("../media/missing.png")).resolves.toBeUndefined();
  });

  it("otherwise categorizes XML parts exactly like parseZip", async () => {
    const buffer = await pptxWithMedia();
    const eager = await parseZip(buffer);
    const lazy = await parseZipLazyMedia(buffer);
    expect([...lazy.slides.keys()]).toEqual([...eager.slides.keys()]);
    expect(lazy.presentation).toBe(eager.presentation);
    expect(lazy.contentTypes).toBe(eager.contentTypes);
  });
});

describe("zip limits", () => {
  it("enforces maxMediaBytes in eager mode", async () => {
    const buffer = await buildCustomPptx({
      extraFiles: { "ppt/media/big.png": fakePngBytes(1024) },
    });
    await expect(parseZip(buffer, { maxMediaBytes: 512 })).rejects.toThrow(/media bytes/);
  });

  it("enforces maxMediaBytes up front in lazy mode via zip directory sizes", async () => {
    const buffer = await buildCustomPptx({
      extraFiles: { "ppt/media/big.png": fakePngBytes(1024) },
    });
    await expect(parseZipLazyMedia(buffer, { maxMediaBytes: 512 })).rejects.toThrow(/media bytes/);
  });

  it("counts media across multiple entries against maxMediaBytes", async () => {
    const buffer = await buildCustomPptx({
      extraFiles: {
        "ppt/media/a.png": fakePngBytes(300),
        "ppt/media/b.png": fakePngBytes(300),
      },
    });
    // Each entry fits alone but not together.
    await expect(parseZip(buffer, { maxMediaBytes: 400 })).rejects.toThrow(/media bytes/);
    await expect(parseZip(buffer, { maxMediaBytes: 600 })).resolves.toBeDefined();
  });

  it("admits the fixture under RECOMMENDED_ZIP_LIMITS in both modes", async () => {
    const buffer = await pptxWithMedia();
    await expect(parseZip(buffer, RECOMMENDED_ZIP_LIMITS)).resolves.toBeDefined();
    const lazy = await parseZipLazyMedia(buffer, RECOMMENDED_ZIP_LIMITS);
    const resolved = await lazy.mediaResolver?.resolve("../media/image1.png");
    expect(resolved?.data).toEqual(PNG);
  });

  it("rejects invalid maxConcurrency values", async () => {
    const buffer = await buildCustomPptx();
    await expect(parseZip(buffer, { maxConcurrency: 0 })).rejects.toThrow(/maxConcurrency/);
    await expect(parseZip(buffer, { maxConcurrency: 1.5 })).rejects.toThrow(/maxConcurrency/);
    await expect(parseZip(buffer, { maxConcurrency: 1 })).resolves.toBeDefined();
  });

  it("enforces maxEntryUncompressedBytes against media entries", async () => {
    const buffer = await buildCustomPptx({
      extraFiles: { "ppt/media/big.png": fakePngBytes(2048) },
    });
    await expect(parseZip(buffer, { maxEntryUncompressedBytes: 1024 })).rejects.toThrow(
      /maxEntryUncompressedBytes/,
    );
  });
});
