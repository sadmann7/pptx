// @vitest-environment jsdom
/**
 * Tests for the node-level operations: duplicateNode, reorderNode, and
 * addImage. Each keeps XML and model in sync, undoes cleanly, and survives a
 * save → reopen round trip. jsdom for namespaced attributes (see
 * operation.test.ts).
 */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { applyEdit, readImageSize } from "../../edit/operation";
import type { PictureNodeData } from "../../model/nodes/picture";
import type { ShapeNodeData } from "../../model/nodes/shape";
import { buildPresentation, PresentationData } from "../../model/presentation";
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

/** A PNG header is enough for the size reader; nothing decodes the pixels here. */
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
