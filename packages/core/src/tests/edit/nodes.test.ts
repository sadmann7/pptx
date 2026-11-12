// @vitest-environment jsdom
// jsdom, since happy-dom's XML parser drops namespaced attributes (see operation.test.ts).
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { applyEdit, copyNodes, readImageSize } from "../../edit/operation";
import type { PictureNodeData } from "../../model/nodes/picture";
import type { ShapeNodeData } from "../../model/nodes/shape";
import { buildPresentation, PresentationData } from "../../model/presentation";
import { getPresentationFonts } from "../../model/presentation-fonts";
import { writePptx } from "../../ooxml/writer";
import { readPptx } from "../../ooxml/zip";
import { buildCustomPptx } from "../fixtures/fixture-extras";

function textShape(id: number, text: string, x = 914400): string {
  return `<p:sp>
<p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="${x}" y="914400"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody>
</p:sp>`;
}

function group(id: number, ...children: string[]): string {
  return `<p:grpSp>
<p:nvGrpSpPr><p:cNvPr id="${id}" name="Group ${id}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/><a:chOff x="914400" y="914400"/><a:chExt cx="1828800" cy="914400"/></a:xfrm></p:grpSpPr>
${children.join("\n")}
</p:grpSp>`;
}

async function openEditable(shapes: string[]): Promise<PresentationData> {
  const buffer = await buildCustomPptx({ slides: [shapes.join("\n")] });
  return buildPresentation(await readPptx(buffer, { keepSourcePackage: true }));
}

async function saveAndReopen(pres: PresentationData): Promise<PresentationData> {
  const saved = await writePptx(pres);
  return buildPresentation(await readPptx(saved.slice().buffer));
}

async function savedZip(pres: PresentationData): Promise<JSZip> {
  return JSZip.loadAsync(await writePptx(pres));
}

const ids = (pres: PresentationData) => pres.slides[0].nodes.map((node) => node.id);

// A PNG header is enough for the size reader.
function pngBytes(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(33);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52]);
  const view = new DataView(bytes.buffer);
  view.setUint32(16, width);
  view.setUint32(20, height);
  return bytes;
}

describe("duplicateNode", () => {
  it("copies a shape above the original with a fresh id and an offset", async () => {
    const pres = await openEditable([textShape(2, "A"), textShape(3, "B", 3657600)]);
    const slideId = pres.slides[0].id;
    const source = pres.slides[0].nodes[0];

    const result = await applyEdit(pres, { type: "duplicateNode", slideId, nodeId: "2" });

    expect(result.createdNodeIds).toEqual(["4"]);
    expect(ids(pres)).toEqual(["2", "4", "3"]);
    const copy = pres.slides[0].nodes[1] as ShapeNodeData;
    expect(copy.position).toEqual({ x: source.position.x + 10, y: source.position.y + 10 });
    expect(copy.size).toEqual(source.size);
    expect(copy.textBody!.paragraphs[0].runs[0].text).toBe("A");

    const reopened = await saveAndReopen(pres);
    expect(ids(reopened)).toEqual(["2", "4", "3"]);
    expect(reopened.slides[0].nodes[1].position.x).toBeCloseTo(source.position.x + 10);

    result.undo();
    expect(ids(pres)).toEqual(["2", "3"]);
    expect(ids(await saveAndReopen(pres))).toEqual(["2", "3"]);
  });

  it("gives a group's children fresh ids too", async () => {
    const pres = await openEditable([group(2, textShape(3, "A"), textShape(4, "B", 1828800))]);
    const slideId = pres.slides[0].id;

    await applyEdit(pres, {
      type: "duplicateNode",
      slideId,
      nodeId: "2",
      offset: { x: 0, y: 50 },
    });

    const zip = await savedZip(pres);
    const xml = await zip.file(slideId)!.async("string");
    const allIds = [...xml.matchAll(/cNvPr id="(\d+)"/g)].map((match) => match[1]);
    expect(allIds).toEqual(["1", "2", "3", "4", "5", "6", "7"]);
  });

  it("is one undo step for a whole selection in a batch", async () => {
    const pres = await openEditable([textShape(2, "A"), textShape(3, "B", 3657600)]);
    const slideId = pres.slides[0].id;

    const result = await applyEdit(pres, {
      type: "batch",
      operations: [
        { type: "duplicateNode", slideId, nodeId: "2" },
        { type: "duplicateNode", slideId, nodeId: "3" },
      ],
    });

    expect(result.createdNodeIds).toEqual(["4", "5"]);
    expect(ids(pres)).toEqual(["2", "4", "3", "5"]);
    result.undo();
    expect(ids(pres)).toEqual(["2", "3"]);
  });
});

describe("reorderNode", () => {
  const shapes = [textShape(2, "A"), textShape(3, "B"), textShape(4, "C")];

  it.each([
    ["front", "2", ["3", "4", "2"]],
    ["back", "4", ["4", "2", "3"]],
    ["forward", "2", ["3", "2", "4"]],
    ["backward", "4", ["2", "4", "3"]],
  ] as const)("brings a node %s", async (order, nodeId, expected) => {
    const pres = await openEditable(shapes);
    const slideId = pres.slides[0].id;

    const result = await applyEdit(pres, { type: "reorderNode", slideId, nodeId, order });

    expect(ids(pres)).toEqual(expected);
    expect(ids(await saveAndReopen(pres))).toEqual(expected);
    result.undo();
    expect(ids(pres)).toEqual(["2", "3", "4"]);
    expect(ids(await saveAndReopen(pres))).toEqual(["2", "3", "4"]);
  });

  it("does nothing at the end of the stack", async () => {
    const pres = await openEditable(shapes);
    const result = await applyEdit(pres, {
      type: "reorderNode",
      slideId: pres.slides[0].id,
      nodeId: "4",
      order: "forward",
    });
    expect(result.affectedSlideIds).toEqual([]);
    expect(ids(pres)).toEqual(["2", "3", "4"]);
  });
});

describe("addImage", () => {
  it("adds a picture on top, sized from the file and centered", async () => {
    const pres = await openEditable([textShape(2, "A")]);
    const slideId = pres.slides[0].id;

    const result = await applyEdit(pres, {
      type: "addImage",
      slideId,
      data: pngBytes(200, 100),
      contentType: "image/png",
      name: "Logo",
    });

    expect(result.createdNodeIds).toEqual(["3"]);
    const picture = pres.slides[0].nodes[1] as PictureNodeData;
    expect(picture.nodeType).toBe("picture");
    expect(picture.size).toEqual({ w: 200, h: 100 });
    expect(picture.position).toEqual({
      x: (pres.width - 200) / 2,
      y: (pres.height - 100) / 2,
    });
    const rel = pres.slides[0].rels.get(picture.blipEmbed!);
    expect(rel?.target).toBe("../media/image1.png");
    expect(pres.media.get("ppt/media/image1.png")).toEqual(pngBytes(200, 100));

    const zip = await savedZip(pres);
    expect(zip.file("ppt/media/image1.png")).not.toBeNull();
    // The deck already declares PNG, so it isn't declared twice.
    const contentTypes = await zip.file("[Content_Types].xml")!.async("string");
    expect(contentTypes.match(/Extension="png"/g)).toHaveLength(1);
    const reopened = await saveAndReopen(pres);
    const reopenedPicture = reopened.slides[0].nodes[1] as PictureNodeData;
    expect(reopenedPicture.nodeType).toBe("picture");
    expect(reopenedPicture.size.w).toBeCloseTo(200);
    expect(reopened.media.has("ppt/media/image1.png")).toBe(true);

    result.undo();
    expect(ids(pres)).toEqual(["2"]);
    expect(pres.media.has("ppt/media/image1.png")).toBe(false);
    const undone = await savedZip(pres);
    expect(undone.file("ppt/media/image1.png")).toBeNull();
  });

  it("declares a new image type, and takes it back on undo", async () => {
    const pres = await openEditable([""]);
    const gif = new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x10, 0x00, 0x10, 0x00]);

    const result = await applyEdit(pres, {
      type: "addImage",
      slideId: pres.slides[0].id,
      data: gif,
      contentType: "image/gif",
    });
    const contentTypes = async () =>
      (await savedZip(pres)).file("[Content_Types].xml")!.async("string");
    expect(await contentTypes()).toContain('Extension="gif" ContentType="image/gif"');

    result.undo();
    expect(await contentTypes()).not.toContain('Extension="gif"');
  });

  it("scales a large image down to fit the slide", async () => {
    const pres = await openEditable([""]);
    await applyEdit(pres, {
      type: "addImage",
      slideId: pres.slides[0].id,
      data: pngBytes(4000, 1000),
      contentType: "image/png",
    });

    const picture = pres.slides[0].nodes[0];
    expect(picture.size.w).toBeCloseTo(pres.width * 0.8);
    expect(picture.size.h).toBeCloseTo((pres.width * 0.8) / 4);
  });

  it("uses the size and position it is given, and picks a free media name", async () => {
    const pres = await openEditable([""]);
    const slideId = pres.slides[0].id;
    for (let i = 0; i < 2; i++) {
      await applyEdit(pres, {
        type: "addImage",
        slideId,
        data: pngBytes(10, 10),
        contentType: "image/png",
        size: { w: 50, h: 40 },
        position: { x: 5, y: 6 },
      });
    }

    const [first, second] = pres.slides[0].nodes as PictureNodeData[];
    expect(first.size).toEqual({ w: 50, h: 40 });
    expect(first.position).toEqual({ x: 5, y: 6 });
    expect(pres.slides[0].rels.get(second.blipEmbed!)?.target).toBe("../media/image2.png");
  });

  it("rejects unsupported image types", async () => {
    const pres = await openEditable([""]);
    await expect(
      applyEdit(pres, {
        type: "addImage",
        slideId: pres.slides[0].id,
        data: new Uint8Array(4),
        contentType: "image/tiff",
      }),
    ).rejects.toThrow(/unsupported image type/);
    expect(pres.slides[0].nodes).toHaveLength(0);
  });
});

describe("readImageSize", () => {
  it("reads PNG, GIF, and JPEG headers", () => {
    expect(readImageSize(pngBytes(640, 480))).toEqual({ w: 640, h: 480 });

    const gif = new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x20, 0x03, 0x58, 0x02]);
    expect(readImageSize(gif)).toEqual({ w: 800, h: 600 });

    // SOI, an APP0 segment, then SOF0 with height 300 and width 400.
    const jpeg = new Uint8Array([
      0xff, 0xd8, 0xff, 0xe0, 0x00, 0x04, 0x00, 0x00, 0xff, 0xc0, 0x00, 0x11, 0x08, 0x01, 0x2c,
      0x01, 0x90, 0x03,
    ]);
    expect(readImageSize(jpeg)).toEqual({ w: 400, h: 300 });

    expect(readImageSize(new Uint8Array([1, 2, 3]))).toBeUndefined();
  });
});

describe("copyNodes and pasteNodes", () => {
  const placeholder = `<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
<p:spPr/>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Title</a:t></a:r></a:p></p:txBody>
</p:sp>`;
  const layoutTitle = `<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>`;

  it("pastes shapes on top of a slide with fresh ids, one undo step", async () => {
    const pres = await openEditable([textShape(2, "A"), textShape(3, "B", 3657600)]);
    const slideId = pres.slides[0].id;
    const clipboard = await copyNodes(pres, slideId, ["3", "2"]);
    expect(clipboard.nodes).toHaveLength(2);

    const result = await applyEdit(pres, { type: "pasteNodes", slideId, clipboard });

    expect(result.createdNodeIds).toEqual(["4", "5"]);
    expect(ids(pres)).toEqual(["2", "3", "4", "5"]);
    const texts = pres.slides[0].nodes.map(
      (node) => (node as ShapeNodeData).textBody!.paragraphs[0].runs[0].text,
    );
    expect(texts).toEqual(["A", "B", "A", "B"]);
    expect(pres.slides[0].nodes[2].position).toEqual(pres.slides[0].nodes[0].position);
    expect(ids(await saveAndReopen(pres))).toEqual(["2", "3", "4", "5"]);

    result.undo();
    expect(ids(pres)).toEqual(["2", "3"]);
  });

  it("offsets pasted shapes", async () => {
    const pres = await openEditable([textShape(2, "A")]);
    const slideId = pres.slides[0].id;
    const clipboard = await copyNodes(pres, slideId, ["2"]);
    const original = pres.slides[0].nodes[0].position;

    await applyEdit(pres, { type: "pasteNodes", slideId, clipboard, offset: { x: 10, y: 20 } });
    expect(pres.slides[0].nodes[1].position).toEqual({ x: original.x + 10, y: original.y + 20 });
  });

  it("writes out an inherited position so it survives another deck", async () => {
    const buffer = await buildCustomPptx({ slides: [placeholder], layoutShapesXml: layoutTitle });
    const pres = buildPresentation(await readPptx(buffer, { keepSourcePackage: true }));
    const clipboard = await copyNodes(pres, pres.slides[0].id, ["2"]);

    expect(clipboard.nodes[0]).toContain('<a:off x="457200" y="274638"/>');
    expect(clipboard.nodes[0]).toContain('<a:ext cx="8229600" cy="1143000"/>');
    const sourceXml = await (await savedZip(pres)).file(pres.slides[0].id)!.async("string");
    expect(sourceXml).not.toContain('<a:off x="457200"');
  });

  it("brings a picture's image into another deck, and reuses it within one", async () => {
    const source = await openEditable([""]);
    const sourceSlide = source.slides[0].id;
    await applyEdit(source, {
      type: "addImage",
      slideId: sourceSlide,
      data: pngBytes(40, 20),
      contentType: "image/png",
    });
    const clipboard = await copyNodes(source, sourceSlide, ["2"]);
    expect(Object.values(clipboard.relationships)[0].media?.extension).toBe("png");

    await applyEdit(source, { type: "pasteNodes", slideId: sourceSlide, clipboard });
    const [first, second] = source.slides[0].nodes as PictureNodeData[];
    expect(source.slides[0].rels.get(second.blipEmbed!)?.target).toBe(
      source.slides[0].rels.get(first.blipEmbed!)?.target,
    );
    expect([...source.media.keys()]).toEqual(["ppt/media/image1.png"]);

    const target = await openEditable([textShape(2, "A")]);
    const targetSlide = target.slides[0].id;
    const result = await applyEdit(target, {
      type: "pasteNodes",
      slideId: targetSlide,
      clipboard,
    });
    const pasted = target.slides[0].nodes[1] as PictureNodeData;
    expect(pasted.nodeType).toBe("picture");
    expect(pasted.size).toEqual({ w: 40, h: 20 });
    expect(target.media.get("ppt/media/image1.png")).toEqual(pngBytes(40, 20));
    const reopened = await saveAndReopen(target);
    expect(reopened.slides[0].nodes[1].nodeType).toBe("picture");
    expect(reopened.media.has("ppt/media/image1.png")).toBe(true);

    result.undo();
    expect(ids(target)).toEqual(["2"]);
    expect(target.media.has("ppt/media/image1.png")).toBe(false);
    expect((await savedZip(target)).file("ppt/media/image1.png")).toBeNull();
  });

  it("leaves out nodes that refer to parts other than images", async () => {
    const linked = textShape(3, "Jump").replace(
      "<a:r><a:t>",
      '<a:r><a:rPr><a:hlinkClick r:id="rId9" action="ppaction://hlinksldjump"/></a:rPr><a:t>',
    );
    const buffer = await buildCustomPptx({
      slides: [`${textShape(2, "A")}\n${linked}`],
      extraFiles: {
        "ppt/slides/_rels/slide1.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slide1.xml"/>
</Relationships>`,
      },
    });
    const pres = buildPresentation(await readPptx(buffer, { keepSourcePackage: true }));

    const clipboard = await copyNodes(pres, pres.slides[0].id, ["2", "3"]);
    expect(clipboard.nodes).toHaveLength(1);
    expect(clipboard.omittedNodeIds).toEqual(["3"]);
  });
});

describe("addTextBox", () => {
  it("adds an empty text box, centered, on top", async () => {
    const pres = await openEditable([textShape(2, "A")]);
    const slideId = pres.slides[0].id;

    const result = await applyEdit(pres, { type: "addTextBox", slideId });

    expect(result.createdNodeIds).toEqual(["3"]);
    const box = pres.slides[0].nodes[1] as ShapeNodeData;
    expect(box.isTextBox).toBe(true);
    expect(box.size.w).toBeCloseTo(pres.width / 3);
    expect(box.position.x).toBeCloseTo((pres.width - box.size.w) / 2);
    expect(box.textBody!.paragraphs).toHaveLength(1);
    expect(box.textBody!.paragraphs[0].runs).toEqual([]);

    const reopened = (await saveAndReopen(pres)).slides[0].nodes[1] as ShapeNodeData;
    expect(reopened.isTextBox).toBe(true);

    result.undo();
    expect(ids(pres)).toEqual(["2"]);
  });

  it("takes text, size, and position", async () => {
    const pres = await openEditable([""]);
    await applyEdit(pres, {
      type: "addTextBox",
      slideId: pres.slides[0].id,
      text: "Hello",
      size: { w: 200, h: 50 },
      position: { x: 10, y: 20 },
    });

    const box = pres.slides[0].nodes[0] as ShapeNodeData;
    expect(box.textBody!.paragraphs[0].runs.map((run) => run.text)).toEqual(["Hello"]);
    expect(box.size).toEqual({ w: 200, h: 50 });
    expect(box.position).toEqual({ x: 10, y: 20 });
  });
});

describe("addShape", () => {
  it("adds a theme-styled preset shape, centered, on top", async () => {
    const pres = await openEditable([textShape(2, "A")]);
    const slideId = pres.slides[0].id;

    const result = await applyEdit(pres, { type: "addShape", slideId, geometry: "ellipse" });

    expect(result.createdNodeIds).toEqual(["3"]);
    const shape = pres.slides[0].nodes[1] as ShapeNodeData;
    expect(shape.presetGeometry).toBe("ellipse");
    expect(shape.isTextBox).toBeFalsy();
    expect(shape.size).toEqual({ w: pres.width / 4, h: pres.width / 4 });
    expect(shape.position.x).toBeCloseTo((pres.width - shape.size.w) / 2);

    const xml = await (await savedZip(pres)).file(slideId)!.async("string");
    expect(xml).toContain('name="Ellipse 2"');
    expect(xml).toMatch(/<a:fillRef idx="1"><a:schemeClr val="accent1"\/><\/a:fillRef>/);
    const reopened = (await saveAndReopen(pres)).slides[0].nodes[1] as ShapeNodeData;
    expect(reopened.presetGeometry).toBe("ellipse");

    result.undo();
    expect(ids(pres)).toEqual(["2"]);
  });

  it("takes size and position, and rejects names that aren't presets", async () => {
    const pres = await openEditable([""]);
    const slideId = pres.slides[0].id;
    await applyEdit(pres, {
      type: "addShape",
      slideId,
      geometry: "rightArrow",
      size: { w: 200, h: 50 },
      position: { x: 10, y: 20 },
    });

    const shape = pres.slides[0].nodes[0] as ShapeNodeData;
    expect(shape.size).toEqual({ w: 200, h: 50 });
    expect(shape.position).toEqual({ x: 10, y: 20 });
    await expect(
      applyEdit(pres, { type: "addShape", slideId, geometry: 'rect"/><x' }),
    ).rejects.toThrow("not a preset geometry");
  });
});

describe("setOutline", () => {
  const lineOf = async (pres: PresentationData) => {
    const xml = await (await savedZip(pres)).file(pres.slides[0].id)!.async("string");
    return /<a:ln[ >].*?<\/a:ln>|<a:ln[^>]*\/>/s.exec(xml)?.[0];
  };

  it("sets an outline's color and width, after the fill", async () => {
    const pres = await openEditable([textShape(2, "A")]);
    const slideId = pres.slides[0].id;

    const result = await applyEdit(pres, {
      type: "setOutline",
      slideId,
      nodeId: "2",
      color: "#336699",
      width: 2,
    });

    expect(await lineOf(pres)).toBe(
      '<a:ln w="25400"><a:solidFill><a:srgbClr val="336699"/></a:solidFill></a:ln>',
    );
    const shape = pres.slides[0].nodes[0] as ShapeNodeData;
    expect(shape.line?.attr("w")).toBe("25400");
    const spPrChildren = shape.source
      .child("spPr")
      .allChildren()
      .map((child) => child.localName);
    expect(spPrChildren).toEqual(["xfrm", "prstGeom", "ln"]);

    result.undo();
    expect(await lineOf(pres)).toBeUndefined();
    expect((pres.slides[0].nodes[0] as ShapeNodeData).line).toBeUndefined();
  });

  it("changes one field at a time, and removes the outline with null", async () => {
    const pres = await openEditable([textShape(2, "A")]);
    const slideId = pres.slides[0].id;
    await applyEdit(pres, { type: "setOutline", slideId, nodeId: "2", color: "FF0000", width: 1 });

    await applyEdit(pres, { type: "setOutline", slideId, nodeId: "2", width: 3 });
    expect(await lineOf(pres)).toBe(
      '<a:ln w="38100"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>',
    );

    await applyEdit(pres, { type: "setOutline", slideId, nodeId: "2", color: null });
    expect(await lineOf(pres)).toBe('<a:ln w="38100"><a:noFill/></a:ln>');
  });

  it("rejects nodes without an outline and bad values", async () => {
    const pres = await openEditable([group(2, textShape(3, "A"))]);
    const slideId = pres.slides[0].id;
    await expect(
      applyEdit(pres, { type: "setOutline", slideId, nodeId: "2", color: "000000" }),
    ).rejects.toThrow(/only supports shapes and pictures/);
  });
});

describe("getPresentationFonts", () => {
  it("lists the theme fonts and the other fonts the text uses", async () => {
    const run = (typeface: string) =>
      `<a:r><a:rPr><a:latin typeface="${typeface}"/></a:rPr><a:t>x</a:t></a:r>`;
    const shape = textShape(2, "A").replace(
      "<a:p><a:r><a:t>A</a:t></a:r></a:p>",
      `<a:p>${run("Georgia")}${run("+mj-lt")}${run("Calibri")}${run("Arial")}</a:p>`,
    );
    const buffer = await buildCustomPptx({
      slides: [shape],
      majorLatin: "Calibri Light",
      minorLatin: "Calibri",
    });
    const pres = buildPresentation(await readPptx(buffer, { keepSourcePackage: true }));

    expect(getPresentationFonts(pres)).toEqual({
      theme: { heading: "Calibri Light", body: "Calibri" },
      embedded: [],
      used: ["Arial", "Georgia"],
    });
  });
});

describe("review fixes", () => {
  it("formats the paragraph mark of an empty <a:p/> at a caret", async () => {
    const pres = await openEditable([
      textShape(2, "A").replace("<a:p><a:r><a:t>A</a:t></a:r></a:p>", "<a:p/>"),
    ]);
    await applyEdit(pres, {
      type: "setTextStyle",
      slideId: pres.slides[0].id,
      nodeId: "2",
      style: { bold: true },
      range: { start: { paragraphIndex: 0, offset: 0 }, end: { paragraphIndex: 0, offset: 0 } },
    });
    const paragraph = (pres.slides[0].nodes[0] as ShapeNodeData).textBody!.paragraphs[0];
    expect(paragraph.endParaRPr?.attr("b")).toBe("1");
  });

  it("points connectors in a duplicated group at the copies", async () => {
    const connector = (id: number, start: number, end: number) =>
      `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="Connector ${id}"/><p:cNvCxnSpPr><a:stCxn id="${start}" idx="0"/><a:endCxn id="${end}" idx="0"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>`;
    const pres = await openEditable([
      textShape(9, "Outside", 5486400),
      group(2, textShape(3, "A"), textShape(4, "B", 1828800), connector(5, 3, 9)),
    ]);
    const slideId = pres.slides[0].id;

    await applyEdit(pres, { type: "duplicateNode", slideId, nodeId: "2" });

    const xml = await (await savedZip(pres)).file(slideId)!.async("string");
    const copy = xml.slice(xml.lastIndexOf("<p:grpSp>"));
    // The copied group is 10 and its "A" is 11: the connector follows the
    // copy, and still ends at the shape outside the group.
    expect(copy).toContain('<a:stCxn id="11" idx="0"/>');
    expect(copy).toContain('<a:endCxn id="9" idx="0"/>');
  });

  it("reorders and duplicates OLE objects as whole frames", async () => {
    const ole = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Object 2"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj r:id="rId5" progId="Excel.Sheet.12"><p:embed/><p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId6"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>`;
    const pres = await openEditable([textShape(2, "A"), ole]);
    const slideId = pres.slides[0].id;
    expect(ids(pres)).toEqual(["2", "3"]);

    await applyEdit(pres, { type: "reorderNode", slideId, nodeId: "3", order: "back" });
    expect(ids(pres)).toEqual(["3", "2"]);
    await applyEdit(pres, { type: "duplicateNode", slideId, nodeId: "3" });

    const reopened = await saveAndReopen(pres);
    expect(reopened.slides[0].nodes.map((node) => node.nodeType)).toEqual([
      "picture",
      "picture",
      "shape",
    ]);
  });

  it("doesn't reuse an undone image's media path", async () => {
    const pres = await openEditable([""]);
    const slideId = pres.slides[0].id;
    const first = await applyEdit(pres, {
      type: "addImage",
      slideId,
      data: pngBytes(10, 10),
      contentType: "image/png",
    });
    first.undo();
    await applyEdit(pres, {
      type: "addImage",
      slideId,
      data: pngBytes(20, 20),
      contentType: "image/png",
    });

    const picture = pres.slides[0].nodes[0] as PictureNodeData;
    expect(pres.slides[0].rels.get(picture.blipEmbed!)?.target).toBe("../media/image2.png");
  });
});
