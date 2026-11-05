// @vitest-environment jsdom
/**
 * Tests for the formatting and slide-creation operations: setTextStyle,
 * setParagraphStyle, addSlide, and setSlideBackground. Like operation.test.ts,
 * each one must keep XML and model in sync, undo cleanly, and survive a
 * save → reopen round trip. jsdom for the same namespaced-attribute reason.
 */
import { describe, expect, it } from "vitest";

import { applyEdit } from "../../edit/operation";
import type { ShapeNodeData } from "../../model/nodes/shape";
import { buildPresentation, PresentationData } from "../../model/presentation";
import { writePptx } from "../../ooxml/writer";
import { readPptx } from "../../ooxml/zip";
import { buildCustomPptx } from "../fixtures/fixture-extras";

function textShape(id: number, paragraphsXml: string): string {
  return `<p:sp>
<p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="4572000" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphsXml}</p:txBody>
</p:sp>`;
}

const run = (text: string, rPr = "") => `<a:r>${rPr}<a:t>${text}</a:t></a:r>`;

const TITLE_PLACEHOLDER = `<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title" hasCustomPrompt="1"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Click to edit title</a:t></a:r></a:p></p:txBody>
</p:sp>
<p:sp>
<p:nvSpPr><p:cNvPr id="3" name="Footer 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" idx="11"/></p:nvPr></p:nvSpPr>
<p:spPr/>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>`;

async function openEditable(slides: string[], layoutShapesXml?: string): Promise<PresentationData> {
  const buffer = await buildCustomPptx({ slides, layoutShapesXml });
  return buildPresentation(await readPptx(buffer, { keepSourcePackage: true }));
}

async function saveAndReopen(pres: PresentationData): Promise<PresentationData> {
  const saved = await writePptx(pres);
  return buildPresentation(await readPptx(saved.slice().buffer));
}

function shapeOn(pres: PresentationData, slideIndex: number, nodeId: string): ShapeNodeData {
  const node = pres.slides[slideIndex].nodes.find((n) => n.id === nodeId);
  if (!node || node.nodeType !== "shape") throw new Error(`shape ${nodeId} not found.`);
  return node as ShapeNodeData;
}

/** `[text, b, i, sz, color, typeface]` per run of one paragraph. */
function runSummary(shape: ShapeNodeData, paragraphIndex = 0) {
  return shape.textBody!.paragraphs[paragraphIndex].runs.map((r) => [
    r.text,
    r.properties?.attr("b"),
    r.properties?.attr("i"),
    r.properties?.attr("sz"),
    r.properties?.child("solidFill").child("srgbClr").attr("val"),
    r.properties?.child("latin").attr("typeface"),
  ]);
}

describe("setTextStyle", () => {
  it("rejects font sizes outside 1–4000pt", async () => {
    const pres = await openEditable([textShape(2, `<a:p>${run("Hello")}</a:p>`)]);
    const edit = (fontSize: number) =>
      applyEdit(pres, {
        type: "setTextStyle",
        slideId: pres.slides[0].id,
        nodeId: "2",
        style: { fontSize },
      });

    await expect(edit(0.5)).rejects.toThrow("invalid font size");
    await expect(edit(4001)).rejects.toThrow("invalid font size");
    await expect(edit(4000)).resolves.toBeDefined();
  });

  it("formats the whole text body when no range is given", async () => {
    const pres = await openEditable([
      textShape(2, `<a:p>${run("Hello")}${run(" world")}</a:p><a:p><a:endParaRPr/></a:p>`),
    ]);
    const slideId = pres.slides[0].id;

    const result = await applyEdit(pres, {
      type: "setTextStyle",
      slideId,
      nodeId: "2",
      style: { bold: true, fontSize: 24, color: "#ff0000", fontFamily: "Arial" },
    });

    const shape = shapeOn(pres, 0, "2");
    expect(runSummary(shape)).toEqual([
      ["Hello", "1", undefined, "2400", "FF0000", "Arial"],
      [" world", "1", undefined, "2400", "FF0000", "Arial"],
    ]);
    // The empty paragraph's mark picks the formatting up for typed text.
    expect(shape.textBody!.paragraphs[1].endParaRPr?.attr("b")).toBe("1");

    const reopened = shapeOn(await saveAndReopen(pres), 0, "2");
    expect(runSummary(reopened)[0]).toEqual(["Hello", "1", undefined, "2400", "FF0000", "Arial"]);

    result.undo();
    expect(runSummary(shapeOn(pres, 0, "2"))).toEqual([
      ["Hello", undefined, undefined, undefined, undefined, undefined],
      [" world", undefined, undefined, undefined, undefined, undefined],
    ]);
    const undone = shapeOn(await saveAndReopen(pres), 0, "2");
    expect(undone.textBody!.paragraphs[0].runs[0].properties?.attr("b")).toBeUndefined();
  });

  it("splits runs at the range edges and keeps their existing formatting", async () => {
    const pres = await openEditable([
      textShape(2, `<a:p>${run("Hello world", '<a:rPr lang="en-US" i="1"/>')}</a:p>`),
    ]);

    await applyEdit(pres, {
      type: "setTextStyle",
      slideId: pres.slides[0].id,
      nodeId: "2",
      style: { bold: true },
      range: { start: { paragraphIndex: 0, offset: 2 }, end: { paragraphIndex: 0, offset: 7 } },
    });

    expect(runSummary(shapeOn(pres, 0, "2"))).toEqual([
      ["He", undefined, "1", undefined, undefined, undefined],
      ["llo w", "1", "1", undefined, undefined, undefined],
      ["orld", undefined, "1", undefined, undefined, undefined],
    ]);
    const reopened = shapeOn(await saveAndReopen(pres), 0, "2");
    expect(reopened.textBody!.paragraphs[0].runs.map((r) => r.text)).toEqual([
      "He",
      "llo w",
      "orld",
    ]);
  });

  it("formats a range across paragraphs, and accepts a backwards range", async () => {
    const pres = await openEditable([
      textShape(2, `<a:p>${run("One")}</a:p><a:p>${run("Two")}</a:p><a:p>${run("Three")}</a:p>`),
    ]);

    await applyEdit(pres, {
      type: "setTextStyle",
      slideId: pres.slides[0].id,
      nodeId: "2",
      style: { italic: true },
      range: { start: { paragraphIndex: 2, offset: 2 }, end: { paragraphIndex: 0, offset: 1 } },
    });

    const shape = shapeOn(pres, 0, "2");
    expect(runSummary(shape, 0).map((r) => [r[0], r[2]])).toEqual([
      ["O", undefined],
      ["ne", "1"],
    ]);
    expect(runSummary(shape, 1).map((r) => [r[0], r[2]])).toEqual([["Two", "1"]]);
    expect(runSummary(shape, 2).map((r) => [r[0], r[2]])).toEqual([
      ["Th", "1"],
      ["ree", undefined],
    ]);
  });

  it("removes run-level values with null so they inherit again", async () => {
    const pres = await openEditable([
      textShape(
        2,
        `<a:p>${run("Hi", '<a:rPr sz="3200"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="Georgia"/></a:rPr>')}</a:p>`,
      ),
    ]);

    await applyEdit(pres, {
      type: "setTextStyle",
      slideId: pres.slides[0].id,
      nodeId: "2",
      style: {
        fontSize: null,
        color: null,
        fontFamily: null,
        underline: true,
        strikethrough: true,
      },
    });

    const rPr = shapeOn(pres, 0, "2").textBody!.paragraphs[0].runs[0].properties!;
    expect(rPr.attr("sz")).toBeUndefined();
    expect(rPr.child("solidFill").exists()).toBe(false);
    expect(rPr.child("latin").exists()).toBe(false);
    expect(rPr.attr("u")).toBe("sng");
    expect(rPr.attr("strike")).toBe("sngStrike");
  });

  it("writes the fill and typeface in schema order", async () => {
    const pres = await openEditable([
      textShape(
        2,
        `<a:p>${run("Hi", '<a:rPr><a:ln w="9525"/><a:hlinkClick r:id=""/></a:rPr>')}</a:p>`,
      ),
    ]);

    await applyEdit(pres, {
      type: "setTextStyle",
      slideId: pres.slides[0].id,
      nodeId: "2",
      style: { color: "00FF00", fontFamily: "Arial" },
    });

    const rPr = shapeOn(pres, 0, "2").textBody!.paragraphs[0].runs[0].properties!;
    expect(rPr.allChildren().map((c) => c.localName)).toEqual([
      "ln",
      "solidFill",
      "latin",
      "hlinkClick",
    ]);
  });

  it("rejects invalid values without touching the text", async () => {
    const pres = await openEditable([textShape(2, `<a:p>${run("Hi")}</a:p>`)]);
    const slideId = pres.slides[0].id;

    await expect(
      applyEdit(pres, { type: "setTextStyle", slideId, nodeId: "2", style: { color: "red" } }),
    ).rejects.toThrow(/invalid color/);
    await expect(
      applyEdit(pres, {
        type: "setTextStyle",
        slideId,
        nodeId: "2",
        style: { bold: true },
        range: { start: { paragraphIndex: 0, offset: 0 }, end: { paragraphIndex: 4, offset: 0 } },
      }),
    ).rejects.toThrow(/out of bounds/);
    expect(shapeOn(pres, 0, "2").textBody!.paragraphs[0].runs[0].properties).toBeUndefined();
  });
});

describe("setParagraphStyle", () => {
  it("aligns the paragraphs a range touches, and undoes", async () => {
    const pres = await openEditable([
      textShape(
        2,
        `<a:p><a:pPr lvl="1"/>${run("One")}</a:p><a:p>${run("Two")}</a:p><a:p>${run("Three")}</a:p>`,
      ),
    ]);

    const result = await applyEdit(pres, {
      type: "setParagraphStyle",
      slideId: pres.slides[0].id,
      nodeId: "2",
      style: { align: "center" },
      range: { start: { paragraphIndex: 0, offset: 1 }, end: { paragraphIndex: 1, offset: 0 } },
    });

    const aligns = (shape: ShapeNodeData) =>
      shape.textBody!.paragraphs.map((p) => p.properties?.attr("algn"));
    expect(aligns(shapeOn(pres, 0, "2"))).toEqual(["ctr", "ctr", undefined]);
    // Existing paragraph properties are kept.
    expect(shapeOn(pres, 0, "2").textBody!.paragraphs[0].level).toBe(1);
    expect(aligns(shapeOn(await saveAndReopen(pres), 0, "2"))).toEqual(["ctr", "ctr", undefined]);

    result.undo();
    expect(aligns(shapeOn(pres, 0, "2"))).toEqual([undefined, undefined, undefined]);
  });

  it("aligns every paragraph without a range", async () => {
    const pres = await openEditable([
      textShape(2, `<a:p>${run("One")}</a:p><a:p>${run("Two")}</a:p>`),
    ]);

    await applyEdit(pres, {
      type: "setParagraphStyle",
      slideId: pres.slides[0].id,
      nodeId: "2",
      style: { align: "justify" },
    });

    expect(
      shapeOn(pres, 0, "2").textBody!.paragraphs.map((p) => p.properties?.attr("algn")),
    ).toEqual(["just", "just"]);
  });
});

describe("addSlide", () => {
  it("adds a slide with the layout's content placeholders after the anchor", async () => {
    const pres = await openEditable(
      [textShape(2, `<a:p>${run("A")}</a:p>`), textShape(2, `<a:p>${run("B")}</a:p>`)],
      TITLE_PLACEHOLDER,
    );
    const [first, second] = pres.slides.map((s) => s.id);

    const result = await applyEdit(pres, { type: "addSlide", afterSlideId: first });

    expect(result.createdSlideId).toBe("ppt/slides/slide3.xml");
    expect(pres.slides.map((s) => s.id)).toEqual([first, "ppt/slides/slide3.xml", second]);
    const added = pres.slides[1];
    expect(added.layoutIndex).toBe("ppt/slideLayouts/slideLayout1.xml");
    // The title placeholder is copied empty; the footer is left off.
    expect(added.nodes).toHaveLength(1);
    const title = added.nodes[0] as ShapeNodeData;
    expect(title.placeholder?.type).toBe("title");
    expect(title.textBody?.paragraphs[0].runs).toEqual([]);

    const reopened = await saveAndReopen(pres);
    expect(reopened.slides).toHaveLength(3);
    expect(reopened.slides[1].nodes).toHaveLength(1);
    expect(reopened.slides[2].nodes[0].id).toBe("2");

    result.undo();
    expect(pres.slides.map((s) => s.id)).toEqual([first, second]);
    expect((await saveAndReopen(pres)).slides).toHaveLength(2);
  });

  it("appends to the end by default and stays editable", async () => {
    const pres = await openEditable([textShape(2, `<a:p>${run("A")}</a:p>`)], TITLE_PLACEHOLDER);

    const result = await applyEdit(pres, { type: "addSlide" });
    expect(pres.slides.map((s) => s.id)).toEqual([pres.slides[0].id, result.createdSlideId]);

    await applyEdit(pres, {
      type: "setTextStyle",
      slideId: result.createdSlideId!,
      nodeId: "2",
      style: { bold: true },
    });
    const reopened = await saveAndReopen(pres);
    const title = shapeOn(reopened, 1, "2");
    expect(title.textBody!.paragraphs[0].endParaRPr?.attr("b")).toBe("1");
  });

  it("rejects an unknown layout", async () => {
    const pres = await openEditable([""]);
    await expect(
      applyEdit(pres, { type: "addSlide", layoutId: "ppt/slideLayouts/nope.xml" }),
    ).rejects.toThrow(/unknown slide layout/);
    expect(pres.slides).toHaveLength(1);
  });
});

describe("setSlideBackground", () => {
  it("sets, replaces, and clears the slide's own background", async () => {
    const pres = await openEditable([textShape(2, `<a:p>${run("A")}</a:p>`)]);
    const slideId = pres.slides[0].id;

    const set = await applyEdit(pres, { type: "setSlideBackground", slideId, color: "#112233" });
    const bgColor = (p: PresentationData) =>
      p.slides[0].background?.child("bgPr").child("solidFill").child("srgbClr").attr("val");
    expect(bgColor(pres)).toBe("112233");
    expect(bgColor(await saveAndReopen(pres))).toBe("112233");

    const replace = await applyEdit(pres, { type: "setSlideBackground", slideId, color: "445566" });
    expect(bgColor(pres)).toBe("445566");

    const clear = await applyEdit(pres, { type: "setSlideBackground", slideId, color: null });
    expect(pres.slides[0].background).toBeUndefined();
    expect((await saveAndReopen(pres)).slides[0].background).toBeUndefined();

    clear.undo();
    expect(bgColor(pres)).toBe("445566");
    replace.undo();
    expect(bgColor(pres)).toBe("112233");
    set.undo();
    expect(pres.slides[0].background).toBeUndefined();
    const reopened = await saveAndReopen(pres);
    expect(reopened.slides[0].background).toBeUndefined();
    expect(reopened.slides[0].nodes).toHaveLength(1);
  });
});
