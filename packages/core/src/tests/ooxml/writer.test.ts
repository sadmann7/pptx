/**
 * Round-trip tests for package retention (`keepSourcePackage`) and `writePptx()`.
 *
 * The core guarantee under test: opening a deck and saving it without edits
 * reproduces every part byte-for-byte, and editing one part re-serializes
 * only that part while everything else still passes through untouched.
 */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { buildPresentation, materializeSlide } from "../../model/presentation";
import { serializePresentation } from "../../model/serialize";
import { writePptx } from "../../ooxml/writer";
import { readPptx } from "../../ooxml/zip";
import { buildPptxWithShapes } from "../fixtures/minimal-pptx";
import { buildRichPptx, tinyPngBytes } from "../fixtures/rich-pptx";

function textShape(id: number, text: string): string {
  return `<p:sp>
<p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="4572000" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody>
</p:sp>`;
}

async function zipEntries(bytes: ArrayBuffer | Uint8Array): Promise<Map<string, Uint8Array>> {
  const zip = await JSZip.loadAsync(bytes);
  const entries = new Map<string, Uint8Array>();
  for (const [path, file] of Object.entries(zip.files)) {
    if (file.dir) continue;
    entries.set(path, await file.async("uint8array"));
  }
  return entries;
}

function decode(bytes: Uint8Array): string {
  return new TextDecoder().decode(bytes);
}

describe("writePptx", () => {
  it("throws when the presentation was parsed without keepSourcePackage", async () => {
    const buffer = await buildPptxWithShapes(textShape(2, "Hello"));
    const presentation = buildPresentation(await readPptx(buffer));
    await expect(writePptx(presentation)).rejects.toThrow(/keepSourcePackage/);
  });

  it("round-trips an unchanged deck byte-for-byte, including binary media", async () => {
    const buffer = await buildRichPptx({
      shapesXml: textShape(2, "Untouched"),
      binaryParts: { "ppt/media/image1.png": tinyPngBytes() },
      extraContentTypesXml: '<Default Extension="png" ContentType="image/png"/>',
    });
    const presentation = buildPresentation(await readPptx(buffer, { keepSourcePackage: true }));

    const saved = await writePptx(presentation);

    const original = await zipEntries(buffer);
    const roundTripped = await zipEntries(saved);

    expect([...roundTripped.keys()].sort()).toEqual([...original.keys()].sort());
    for (const [path, originalBytes] of original) {
      expect(roundTripped.get(path), path).toEqual(originalBytes);
    }
  });

  it("reopens the saved output as an identical presentation", async () => {
    const buffer = await buildPptxWithShapes(textShape(2, "Stable"));
    const presentation = buildPresentation(await readPptx(buffer, { keepSourcePackage: true }));

    const saved = await writePptx(presentation);
    const reopened = buildPresentation(await readPptx(saved.slice().buffer));

    expect(serializePresentation(reopened)).toEqual(serializePresentation(presentation));
  });

  it("re-serializes a dirty slide from its mutated XML and passes other parts through", async () => {
    const buffer = await buildPptxWithShapes(textShape(2, "Before edit"));
    const presentation = buildPresentation(await readPptx(buffer, { keepSourcePackage: true }));

    const slide = presentation.slides[0];
    const shape = slide.nodes[0];
    const textEl = shape.source.child("txBody").child("p").child("r").child("t").element;
    expect(textEl).not.toBeNull();
    textEl!.textContent = "After edit";
    presentation.sourcePackage!.markDirty(slide.id);

    const saved = await writePptx(presentation);

    const original = await zipEntries(buffer);
    const roundTripped = await zipEntries(saved);
    const slideXml = decode(roundTripped.get("ppt/slides/slide1.xml")!);

    expect(slideXml).toContain("After edit");
    expect(slideXml).not.toContain("Before edit");
    // Original XML declaration is preserved on the re-serialized part.
    expect(slideXml.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')).toBe(
      true,
    );

    for (const [path, originalBytes] of original) {
      if (path === "ppt/slides/slide1.xml") continue;
      expect(roundTripped.get(path), path).toEqual(originalBytes);
    }

    // The edit is visible when the saved file is opened again.
    const reopened = buildPresentation(await readPptx(saved.slice().buffer));
    const serialized = serializePresentation(reopened);
    expect(JSON.stringify(serialized)).toContain("After edit");
  });

  it("registers lazily materialized slide XML so lazy decks stay editable", async () => {
    const buffer = await buildPptxWithShapes(textShape(2, "Lazy text"));
    const presentation = buildPresentation(await readPptx(buffer, { keepSourcePackage: true }), {
      lazySlides: true,
    });

    const slide = presentation.slides[0];
    expect(presentation.sourcePackage!.getXmlRoot(slide.id)).toBeUndefined();

    materializeSlide(presentation, slide);
    expect(presentation.sourcePackage!.getXmlRoot(slide.id)).toBeDefined();

    const textEl = slide.nodes[0].source.child("txBody").child("p").child("r").child("t").element;
    textEl!.textContent = "Lazy edited";
    presentation.sourcePackage!.markDirty(slide.id);

    const saved = await writePptx(presentation);
    const slideXml = decode((await zipEntries(saved)).get("ppt/slides/slide1.xml")!);
    expect(slideXml).toContain("Lazy edited");
  });

  it("supports raw part replacement and deletion via the package", async () => {
    const buffer = await buildPptxWithShapes(textShape(2, "Raw ops"));
    const presentation = buildPresentation(await readPptx(buffer, { keepSourcePackage: true }));
    const sourcePackage = presentation.sourcePackage!;

    sourcePackage.setEntry("docProps/custom.xml", "<custom/>");
    expect(sourcePackage.deleteEntry("ppt/tableStyles.xml")).toBe(false); // fixture has none
    expect(sourcePackage.deleteEntry("ppt/theme/theme1.xml")).toBe(true);

    const entries = await zipEntries(await writePptx(presentation));
    expect(decode(entries.get("docProps/custom.xml")!)).toBe("<custom/>");
    expect(entries.has("ppt/theme/theme1.xml")).toBe(false);
  });

  it("rejects markDirty for parts without a registered XML document", async () => {
    const buffer = await buildPptxWithShapes(textShape(2, "No doc"));
    const presentation = buildPresentation(await readPptx(buffer, { keepSourcePackage: true }));

    // Rels parts are parsed into RelEntry maps, not retained XML documents.
    expect(() => presentation.sourcePackage!.markDirty("ppt/_rels/presentation.xml.rels")).toThrow(
      /no XML document/,
    );
  });
});
