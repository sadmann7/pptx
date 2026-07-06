// @vitest-environment jsdom
/**
 * Tests for applyEdit(): every operation mutates the part XML, keeps the
 * typed model in sync, and survives a save → reopen round trip through the
 * real pipeline. Undo closures must restore both XML and model.
 *
 * Runs under jsdom (not the default happy-dom) because happy-dom's XML
 * parser drops namespaced attributes (e.g. `r:id` on `p:sldId`), which the
 * slide-level operations depend on. Browsers and jsdom parse them correctly.
 */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { applyEdit } from "../edit/operations";
import type { ShapeNodeData } from "../model/nodes/shape-node";
import { buildPresentation, PresentationData } from "../model/presentation";
import { parseZip } from "../ooxml/zip-parser";
import { writePptx } from "../write/pptx-writer";
import { buildCustomPptx } from "./fixture-extras";

function textShape(id: number, text: string, extraSpPr = ""): string {
  return `<p:sp>
<p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="4572000" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${extraSpPr}</p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody>
</p:sp>`;
}

/** Shape without an xfrm — position/size resolve to 0 (inherited-transform case). */
function shapeWithoutXfrm(id: number): string {
  return `<p:sp>
<p:nvSpPr><p:cNvPr id="${id}" name="NoXfrm ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>`;
}

async function openEditable(slides: string[]): Promise<PresentationData> {
  const buffer = await buildCustomPptx({ slides });
  return buildPresentation(await parseZip(buffer, {}, { keepPackage: true }));
}

async function saveAndReopen(pres: PresentationData): Promise<PresentationData> {
  const saved = await writePptx(pres);
  return buildPresentation(await parseZip(saved.slice().buffer));
}

async function savedPartText(pres: PresentationData, path: string): Promise<string> {
  const zip = await JSZip.loadAsync(await writePptx(pres));
  const file = zip.file(path);
  if (!file) throw new Error(`part not in saved output: ${path}`);
  return file.async("string");
}

function shapeOn(pres: PresentationData, slideIndex: number, nodeId: string): ShapeNodeData {
  const node = pres.slides[slideIndex].nodes.find((n) => n.id === nodeId);
  if (!node || node.nodeType !== "shape") throw new Error(`shape ${nodeId} not found`);
  return node as ShapeNodeData;
}

describe("applyEdit guards", () => {
  it("rejects presentations parsed without keepPackage", async () => {
    const buffer = await buildCustomPptx({ slides: [textShape(2, "x")] });
    const pres = buildPresentation(await parseZip(buffer));
    await expect(
      applyEdit(pres, {
        type: "setTextRun",
        slideId: pres.slides[0].id,
        nodeId: "2",
        paragraphIndex: 0,
        runIndex: 0,
        text: "y",
      }),
    ).rejects.toThrow(/keepPackage/);
  });

  it("rejects unknown slide and node ids", async () => {
    const pres = await openEditable([textShape(2, "x")]);
    await expect(
      applyEdit(pres, { type: "deleteNode", slideId: "ppt/slides/nope.xml", nodeId: "2" }),
    ).rejects.toThrow(/unknown slide/);
    await expect(
      applyEdit(pres, { type: "deleteNode", slideId: pres.slides[0].id, nodeId: "99" }),
    ).rejects.toThrow(/no top-level node/);
  });
});

describe("setTextRun", () => {
  it("updates model, XML, and saved output; undo restores", async () => {
    const pres = await openEditable([textShape(2, "Original")]);
    const slideId = pres.slides[0].id;

    const result = await applyEdit(pres, {
      type: "setTextRun",
      slideId,
      nodeId: "2",
      paragraphIndex: 0,
      runIndex: 0,
      text: "Edited",
    });

    expect(result.affectedSlideIds).toEqual([slideId]);
    expect(shapeOn(pres, 0, "2").textBody?.paragraphs[0].runs[0].text).toBe("Edited");

    const reopened = await saveAndReopen(pres);
    expect(shapeOn(reopened, 0, "2").textBody?.paragraphs[0].runs[0].text).toBe("Edited");

    result.undo();
    expect(shapeOn(pres, 0, "2").textBody?.paragraphs[0].runs[0].text).toBe("Original");
    const reopenedAfterUndo = await saveAndReopen(pres);
    expect(shapeOn(reopenedAfterUndo, 0, "2").textBody?.paragraphs[0].runs[0].text).toBe(
      "Original",
    );
  });
});

describe("setNodeTransform", () => {
  it("rewrites an existing xfrm with EMU-converted values", async () => {
    const pres = await openEditable([textShape(2, "Move me")]);
    const slideId = pres.slides[0].id;

    await applyEdit(pres, {
      type: "setNodeTransform",
      slideId,
      nodeId: "2",
      position: { x: 96, y: 192 }, // 1in, 2in
      size: { w: 288, h: 96 }, // 3in, 1in
      rotation: 45,
    });

    const node = shapeOn(pres, 0, "2");
    expect(node.position).toEqual({ x: 96, y: 192 });
    expect(node.size).toEqual({ w: 288, h: 96 });
    expect(node.rotation).toBe(45);

    const slideXml = await savedPartText(pres, "ppt/slides/slide1.xml");
    expect(slideXml).toContain('<a:off x="914400" y="1828800"/>');
    expect(slideXml).toContain('<a:ext cx="2743200" cy="914400"/>');
    expect(slideXml).toContain('rot="2700000"');

    const reopened = await saveAndReopen(pres);
    expect(shapeOn(reopened, 0, "2").position).toEqual({ x: 96, y: 192 });
    expect(shapeOn(reopened, 0, "2").rotation).toBe(45);
  });

  it("creates an xfrm for shapes without one, and undo removes it again", async () => {
    const pres = await openEditable([shapeWithoutXfrm(2)]);
    const slideId = pres.slides[0].id;

    const result = await applyEdit(pres, {
      type: "setNodeTransform",
      slideId,
      nodeId: "2",
      position: { x: 48, y: 48 },
      size: { w: 96, h: 96 },
    });

    const slideXml = await savedPartText(pres, "ppt/slides/slide1.xml");
    // xfrm must be the first child of spPr, before prstGeom.
    expect(slideXml).toMatch(/<p:spPr><a:xfrm><a:off [^>]+\/><a:ext [^>]+\/><\/a:xfrm><a:prstGeom/);

    const reopened = await saveAndReopen(pres);
    expect(shapeOn(reopened, 0, "2").position).toEqual({ x: 48, y: 48 });
    expect(shapeOn(reopened, 0, "2").size).toEqual({ w: 96, h: 96 });

    result.undo();
    const undoneXml = await savedPartText(pres, "ppt/slides/slide1.xml");
    // The created xfrm is removed again — spPr starts with prstGeom as before.
    expect(undoneXml).toContain("<p:spPr><a:prstGeom");
    expect(shapeOn(pres, 0, "2").position).toEqual({ x: 0, y: 0 });
  });

  it("sets and clears flips", async () => {
    const pres = await openEditable([textShape(2, "Flip")]);
    const slideId = pres.slides[0].id;

    await applyEdit(pres, { type: "setNodeTransform", slideId, nodeId: "2", flipH: true });
    expect(await savedPartText(pres, "ppt/slides/slide1.xml")).toContain('flipH="1"');

    await applyEdit(pres, { type: "setNodeTransform", slideId, nodeId: "2", flipH: false });
    expect(await savedPartText(pres, "ppt/slides/slide1.xml")).not.toContain("flipH");
    expect(shapeOn(pres, 0, "2").flipH).toBe(false);
  });
});

describe("setSolidFill", () => {
  it("replaces an existing fill and keeps schema position before a:ln", async () => {
    const pres = await openEditable([
      textShape(
        2,
        "Filled",
        '<a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>',
      ),
    ]);
    const slideId = pres.slides[0].id;

    const result = await applyEdit(pres, {
      type: "setSolidFill",
      slideId,
      nodeId: "2",
      color: "#ff8800",
    });

    const slideXml = await savedPartText(pres, "ppt/slides/slide1.xml");
    expect(slideXml).toContain('<a:solidFill><a:srgbClr val="FF8800"/></a:solidFill><a:ln');
    expect(slideXml).not.toContain('val="112233"');
    expect(shapeOn(pres, 0, "2").fill?.child("srgbClr").attr("val")).toBe("FF8800");

    result.undo();
    const undoneXml = await savedPartText(pres, "ppt/slides/slide1.xml");
    expect(undoneXml).toContain('val="112233"');
    expect(shapeOn(pres, 0, "2").fill?.child("srgbClr").attr("val")).toBe("112233");
  });

  it("adds a fill to shapes without one and validates colors", async () => {
    const pres = await openEditable([textShape(2, "No fill")]);
    const slideId = pres.slides[0].id;

    await expect(
      applyEdit(pres, { type: "setSolidFill", slideId, nodeId: "2", color: "red" }),
    ).rejects.toThrow(/invalid color/);

    await applyEdit(pres, { type: "setSolidFill", slideId, nodeId: "2", color: "00FF00" });
    const reopened = await saveAndReopen(pres);
    expect(shapeOn(reopened, 0, "2").fill?.child("srgbClr").attr("val")).toBe("00FF00");
  });
});

describe("deleteNode", () => {
  it("removes the node from model and XML; undo restores order", async () => {
    const pres = await openEditable([textShape(2, "First") + textShape(3, "Second")]);
    const slideId = pres.slides[0].id;

    const result = await applyEdit(pres, { type: "deleteNode", slideId, nodeId: "2" });

    expect(pres.slides[0].nodes.map((n) => n.id)).toEqual(["3"]);
    const reopened = await saveAndReopen(pres);
    expect(reopened.slides[0].nodes.map((n) => n.id)).toEqual(["3"]);

    result.undo();
    expect(pres.slides[0].nodes.map((n) => n.id)).toEqual(["2", "3"]);
    const reopenedAfterUndo = await saveAndReopen(pres);
    expect(reopenedAfterUndo.slides[0].nodes.map((n) => n.id)).toEqual(["2", "3"]);
    expect(shapeOn(reopenedAfterUndo, 0, "2").textBody?.paragraphs[0].runs[0].text).toBe("First");
  });
});

describe("moveSlide", () => {
  it("reorders slides in model and saved sldIdLst; undo restores", async () => {
    const pres = await openEditable([textShape(2, "A"), textShape(2, "B"), textShape(2, "C")]);
    const ids = pres.slides.map((s) => s.id);

    const result = await applyEdit(pres, { type: "moveSlide", slideId: ids[0], toIndex: 2 });

    expect(pres.slides.map((s) => s.id)).toEqual([ids[1], ids[2], ids[0]]);
    expect(pres.slides.map((s) => s.index)).toEqual([0, 1, 2]);
    expect(pres.slideToLayout.get(0)).toBe("ppt/slideLayouts/slideLayout1.xml");

    const reopened = await saveAndReopen(pres);
    const texts = reopened.slides.map(
      (s) => (s.nodes[0] as ShapeNodeData).textBody?.paragraphs[0].runs[0].text,
    );
    expect(texts).toEqual(["B", "C", "A"]);

    result.undo();
    expect(pres.slides.map((s) => s.id)).toEqual(ids);
    const reopenedAfterUndo = await saveAndReopen(pres);
    const undoneTexts = reopenedAfterUndo.slides.map(
      (s) => (s.nodes[0] as ShapeNodeData).textBody?.paragraphs[0].runs[0].text,
    );
    expect(undoneTexts).toEqual(["A", "B", "C"]);
  });

  it("is a no-op when the slide is already at the target index", async () => {
    const pres = await openEditable([textShape(2, "A"), textShape(2, "B")]);
    const result = await applyEdit(pres, {
      type: "moveSlide",
      slideId: pres.slides[0].id,
      toIndex: 0,
    });
    expect(result.affectedSlideIds).toEqual([]);
    expect(pres.pkg!.isDirty("ppt/presentation.xml")).toBe(false);
  });
});

describe("duplicateSlide", () => {
  it("inserts an independent copy after the source", async () => {
    const pres = await openEditable([textShape(2, "One"), textShape(2, "Two")]);
    const sourceId = pres.slides[0].id;

    const result = await applyEdit(pres, { type: "duplicateSlide", slideId: sourceId });

    expect(result.createdSlideId).toBe("ppt/slides/slide3.xml");
    expect(pres.slides).toHaveLength(3);
    expect(pres.slides[1].id).toBe("ppt/slides/slide3.xml");

    // The copy is editable independently of the source.
    await applyEdit(pres, {
      type: "setTextRun",
      slideId: result.createdSlideId!,
      nodeId: "2",
      paragraphIndex: 0,
      runIndex: 0,
      text: "One (copy)",
    });

    const reopened = await saveAndReopen(pres);
    const texts = reopened.slides.map(
      (s) => (s.nodes[0] as ShapeNodeData).textBody?.paragraphs[0].runs[0].text,
    );
    expect(texts).toEqual(["One", "One (copy)", "Two"]);

    const ctXml = await savedPartText(pres, "[Content_Types].xml");
    expect(ctXml).toContain('PartName="/ppt/slides/slide3.xml"');
    const relsXml = await savedPartText(pres, "ppt/_rels/presentation.xml.rels");
    expect(relsXml).toContain('Target="slides/slide3.xml"');
  });

  it("undo removes the copy and its package entries", async () => {
    const pres = await openEditable([textShape(2, "Solo")]);
    const result = await applyEdit(pres, { type: "duplicateSlide", slideId: pres.slides[0].id });

    result.undo();

    expect(pres.slides).toHaveLength(1);
    const reopened = await saveAndReopen(pres);
    expect(reopened.slides).toHaveLength(1);
    const zip = await JSZip.loadAsync(await writePptx(pres));
    expect(zip.file("ppt/slides/slide2.xml")).toBeNull();
    expect(await savedPartText(pres, "[Content_Types].xml")).not.toContain("slide2.xml");
  });
});

describe("deleteSlide", () => {
  it("removes the slide, its parts, rel, and content-type override", async () => {
    const pres = await openEditable([textShape(2, "A"), textShape(2, "B"), textShape(2, "C")]);
    const deletedId = pres.slides[1].id;

    await applyEdit(pres, { type: "deleteSlide", slideId: deletedId });

    expect(pres.slides.map((s) => s.index)).toEqual([0, 1]);
    const reopened = await saveAndReopen(pres);
    const texts = reopened.slides.map(
      (s) => (s.nodes[0] as ShapeNodeData).textBody?.paragraphs[0].runs[0].text,
    );
    expect(texts).toEqual(["A", "C"]);

    const zip = await JSZip.loadAsync(await writePptx(pres));
    expect(zip.file("ppt/slides/slide2.xml")).toBeNull();
    expect(zip.file("ppt/slides/_rels/slide2.xml.rels")).toBeNull();
    expect(await savedPartText(pres, "[Content_Types].xml")).not.toContain(
      "/ppt/slides/slide2.xml",
    );
    expect(await savedPartText(pres, "ppt/_rels/presentation.xml.rels")).not.toContain(
      'Target="slides/slide2.xml"',
    );
  });

  it("undo fully restores the deck, and the restored slide stays editable", async () => {
    const pres = await openEditable([textShape(2, "Keep"), textShape(2, "Gone")]);
    const deletedId = pres.slides[1].id;

    const result = await applyEdit(pres, { type: "deleteSlide", slideId: deletedId });
    result.undo();

    expect(pres.slides).toHaveLength(2);
    const reopened = await saveAndReopen(pres);
    const texts = reopened.slides.map(
      (s) => (s.nodes[0] as ShapeNodeData).textBody?.paragraphs[0].runs[0].text,
    );
    expect(texts).toEqual(["Keep", "Gone"]);

    // The restored slide's live XML is re-registered — edits still work.
    await applyEdit(pres, {
      type: "setTextRun",
      slideId: deletedId,
      nodeId: "2",
      paragraphIndex: 0,
      runIndex: 0,
      text: "Back again",
    });
    const reopened2 = await saveAndReopen(pres);
    expect(
      (reopened2.slides[1].nodes[0] as ShapeNodeData).textBody?.paragraphs[0].runs[0].text,
    ).toBe("Back again");
  });

  it("refuses to delete the last slide", async () => {
    const pres = await openEditable([textShape(2, "Only")]);
    await expect(
      applyEdit(pres, { type: "deleteSlide", slideId: pres.slides[0].id }),
    ).rejects.toThrow(/last slide/);
  });
});
