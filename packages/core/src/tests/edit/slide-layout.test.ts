// @vitest-environment jsdom
// jsdom, since happy-dom's XML parser drops namespaced attributes (see operation.test.ts).
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { applyEdit } from "../../edit/operation";
import type { ShapeNodeData } from "../../model/nodes/shape";
import { buildPresentation, PresentationData } from "../../model/presentation";
import { getLayoutGroups, getLayouts } from "../../model/presentation-layouts";
import { emuToPx } from "../../ooxml/unit";
import { writePptx } from "../../ooxml/writer";
import { readPptx } from "../../ooxml/zip";
import { renderLayout } from "../../renderer/slide";
import { buildCustomPptx } from "../fixtures/fixture-extras";

const TITLE_SLIDE = "ppt/slideLayouts/slideLayout1.xml";
const TWO_CONTENT = "ppt/slideLayouts/slideLayout2.xml";
const TITLE_ONLY = "ppt/slideLayouts/slideLayout3.xml";

function layoutPlaceholder(
  id: number,
  ph: string,
  [x, y, cx, cy]: number[],
  prompt = "",
): string {
  return `<p:sp>
<p:nvSpPr><p:cNvPr id="${id}" name="Placeholder ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${ph}</p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p>${prompt ? `<a:r><a:t>${prompt}</a:t></a:r>` : ""}</a:p></p:txBody>
</p:sp>`;
}

const CTR_TITLE_BOX = [914400, 2130425, 10363200, 1470025];
const SUBTITLE_BOX = [1828800, 3886200, 8534400, 1752600];
const TITLE_BOX = [609600, 274638, 10972800, 1143000];
const LEFT_BOX = [609600, 1600200, 5384800, 4525963];
const RIGHT_BOX = [6197600, 1600200, 5384800, 4525963];

const TITLE_SLIDE_SHAPES =
  layoutPlaceholder(2, '<p:ph type="ctrTitle"/>', CTR_TITLE_BOX) +
  layoutPlaceholder(3, '<p:ph type="subTitle" idx="1"/>', SUBTITLE_BOX);

const TWO_CONTENT_SHAPES =
  layoutPlaceholder(
    2,
    '<p:ph type="title" hasCustomPrompt="1"/>',
    TITLE_BOX,
    "Add a short title",
  ) +
  layoutPlaceholder(3, '<p:ph sz="half" idx="1"/>', LEFT_BOX) +
  layoutPlaceholder(4, '<p:ph sz="half" idx="2"/>', RIGHT_BOX) +
  layoutPlaceholder(
    5,
    '<p:ph type="sldNum" sz="quarter" idx="12"/>',
    [8610600, 6356350, 2743200, 365125],
  );

const TITLE_ONLY_SHAPES = layoutPlaceholder(
  2,
  '<p:ph type="title"/>',
  TITLE_BOX,
);

function slidePlaceholder(id: number, ph: string, text: string): string {
  const paragraph = text ? `<a:r><a:t>${text}</a:t></a:r>` : "";
  return `<p:sp>
<p:nvSpPr><p:cNvPr id="${id}" name="Placeholder ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${ph}</p:nvPr></p:nvSpPr>
<p:spPr/>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p>${paragraph}</a:p></p:txBody>
</p:sp>`;
}

const TEXT_BOX = `<p:sp>
<p:nvSpPr><p:cNvPr id="9" name="Note"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="457200" y="6096000"/><a:ext cx="2286000" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Note</a:t></a:r></a:p></p:txBody>
</p:sp>`;

// 400x200px, so its aspect ratio differs from the content placeholder it lands in.
const PICTURE_PLACEHOLDER = `<p:pic>
<p:nvPicPr><p:cNvPr id="4" name="Picture 3"/><p:cNvPicPr><a:picLocks noGrp="1"/></p:cNvPicPr><p:nvPr><p:ph type="pic" idx="3"/></p:nvPr></p:nvPicPr>
<p:blipFill><a:blip r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="1905000"/></a:xfrm></p:spPr>
</p:pic>`;

async function openEditable(slides: string[]): Promise<PresentationData> {
  const buffer = await buildCustomPptx({
    slides,
    layoutShapesXml: TITLE_SLIDE_SHAPES,
    extraLayouts: [
      { name: "Two Content", shapesXml: TWO_CONTENT_SHAPES },
      { name: "Title Only", shapesXml: TITLE_ONLY_SHAPES },
    ],
  });
  return buildPresentation(await readPptx(buffer, { keepSourcePackage: true }));
}

async function saveAndReopen(
  pres: PresentationData,
): Promise<PresentationData> {
  const saved = await writePptx(pres);
  return buildPresentation(await readPptx(saved.slice().buffer));
}

async function savedPartText(
  pres: PresentationData,
  path: string,
): Promise<string> {
  const zip = await JSZip.loadAsync(await writePptx(pres));
  const file = zip.file(path);
  if (!file) throw new Error(`part not in saved output: ${path}.`);
  return file.async("string");
}

function summarize(pres: PresentationData, slideIndex = 0) {
  return pres.slides[slideIndex].nodes.map((node) => ({
    id: node.id,
    type: node.placeholder?.type,
    idx: node.placeholder?.idx,
    text: (node as ShapeNodeData).textBody?.paragraphs
      .flatMap((paragraph) => paragraph.runs.map((run) => run.text))
      .join(""),
    box: [node.position.x, node.position.y, node.size.w, node.size.h],
  }));
}

function px(box: number[]): number[] {
  return box.map(emuToPx);
}

describe("getLayouts", () => {
  it("lists layouts in the master's order with their names and theme", async () => {
    const pres = await openEditable([""]);

    expect(getLayouts(pres)).toEqual([
      {
        id: TITLE_SLIDE,
        name: "Blank",
        masterId: "ppt/slideMasters/slideMaster1.xml",
        themeName: "Office Theme",
      },
      {
        id: TWO_CONTENT,
        name: "Two Content",
        masterId: "ppt/slideMasters/slideMaster1.xml",
        themeName: "Office Theme",
      },
      {
        id: TITLE_ONLY,
        name: "Title Only",
        masterId: "ppt/slideMasters/slideMaster1.xml",
        themeName: "Office Theme",
      },
    ]);
  });
});

function textCard(paragraphs: string[]): string {
  return `<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="TextBox 1"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr/>
<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs.map((text) => `<a:p><a:r><a:t>${text}</a:t></a:r></a:p>`).join("")}</p:txBody>
</p:sp>`;
}

describe("getLayoutGroups", () => {
  it("heads groups with the template's section divider layouts and leaves them out", async () => {
    const buffer = await buildCustomPptx({
      slides: [""],
      layoutShapesXml: TITLE_SLIDE_SHAPES,
      extraLayouts: [
        {
          name: "Process layouts",
          shapesXml: textCard(["Process", "layouts"]),
        },
        { name: "Two Content", shapesXml: TWO_CONTENT_SHAPES },
        { name: "Empty", shapesXml: "" },
        {
          name: "Warning",
          shapesXml: textCard(["Do not use the next layouts"]),
        },
        { name: "Unused layouts", shapesXml: textCard(["Unused layouts"]) },
      ],
    });
    const pres = buildPresentation(await readPptx(buffer));

    expect(
      getLayoutGroups(pres).map((group) => ({
        id: group.id,
        name: group.name,
        layouts: group.layouts.map((layout) => layout.name),
      })),
    ).toEqual([
      {
        id: "ppt/slideMasters/slideMaster1.xml",
        name: "Office Theme",
        layouts: ["Blank"],
      },
      {
        id: TWO_CONTENT,
        name: "Process layouts",
        layouts: ["Two Content", "Empty", "Warning"],
      },
    ]);
  });
});

describe("setSlideLayout", () => {
  it("moves placeholder content into the new layout's placeholders", async () => {
    const pres = await openEditable([
      slidePlaceholder(2, '<p:ph type="ctrTitle"/>', "Quarterly review") +
        slidePlaceholder(3, '<p:ph type="subTitle" idx="1"/>', "Q3 results") +
        TEXT_BOX,
    ]);
    const slideId = pres.slides[0].id;
    const before = summarize(pres);
    const slideXmlBefore = await savedPartText(pres, slideId);

    const result = await applyEdit(pres, {
      type: "setSlideLayout",
      slideId,
      layoutId: TWO_CONTENT,
    });

    const expected = [
      {
        id: "2",
        type: "title",
        idx: undefined,
        text: "Quarterly review",
        box: px(TITLE_BOX),
      },
      {
        id: "3",
        type: undefined,
        idx: 1,
        text: "Q3 results",
        box: px(LEFT_BOX),
      },
      before[2],
      { id: "10", type: undefined, idx: 2, text: "", box: px(RIGHT_BOX) },
    ];
    expect(result.affectedSlideIds).toEqual([slideId]);
    expect(pres.slides[0].layoutIndex).toBe(TWO_CONTENT);
    expect(summarize(pres)).toEqual(expected);

    const reopened = await saveAndReopen(pres);
    expect(reopened.slides[0].layoutIndex).toBe(TWO_CONTENT);
    expect(summarize(reopened)).toEqual(expected);
    const savedXml = await savedPartText(pres, slideId);
    expect(savedXml).not.toContain("hasCustomPrompt");
    expect(savedXml).not.toContain("sldNum");

    result.undo();
    expect(pres.slides[0].layoutIndex).toBe(TITLE_SLIDE);
    expect(summarize(pres)).toEqual(before);
    expect(await savedPartText(pres, slideId)).toBe(slideXmlBefore);
    expect(
      await savedPartText(pres, "ppt/slides/_rels/slide1.xml.rels"),
    ).toContain('Target="../slideLayouts/slideLayout1.xml"');
  });

  it("keeps content that has no matching placeholder in place and drops empty ones", async () => {
    const pres = await openEditable([
      slidePlaceholder(2, '<p:ph type="ctrTitle"/>', "Title") +
        slidePlaceholder(3, '<p:ph type="subTitle" idx="1"/>', "Kept text") +
        slidePlaceholder(4, '<p:ph type="body" idx="7"/>', ""),
    ]);
    const slideId = pres.slides[0].id;

    await applyEdit(pres, {
      type: "setSlideLayout",
      slideId,
      layoutId: TITLE_ONLY,
    });

    const expected = [
      {
        id: "2",
        type: "title",
        idx: undefined,
        text: "Title",
        box: px(TITLE_BOX),
      },
      {
        id: "3",
        type: "subTitle",
        idx: 1,
        text: "Kept text",
        box: px(SUBTITLE_BOX),
      },
    ];
    expect(summarize(pres)).toEqual(expected);
    expect(summarize(await saveAndReopen(pres))).toEqual(expected);
  });

  it("moves kept content off an idx the new layout uses", async () => {
    const buffer = await buildCustomPptx({
      slides: [
        slidePlaceholder(2, '<p:ph type="ctrTitle"/>', "Title") +
          slidePlaceholder(3, '<p:ph type="subTitle" idx="1"/>', "Kept text"),
      ],
      layoutShapesXml: TITLE_SLIDE_SHAPES,
      extraLayouts: [
        {
          name: "Title and Picture",
          shapesXml:
            layoutPlaceholder(2, '<p:ph type="title"/>', TITLE_BOX) +
            layoutPlaceholder(3, '<p:ph type="pic" idx="1"/>', RIGHT_BOX),
        },
      ],
    });
    const pres = buildPresentation(
      await readPptx(buffer, { keepSourcePackage: true }),
    );
    const slideId = pres.slides[0].id;
    const before = summarize(pres);
    const slideXmlBefore = await savedPartText(pres, slideId);

    const result = await applyEdit(pres, {
      type: "setSlideLayout",
      slideId,
      layoutId: "ppt/slideLayouts/slideLayout2.xml",
    });

    const expected = [
      {
        id: "2",
        type: "title",
        idx: undefined,
        text: "Title",
        box: px(TITLE_BOX),
      },
      {
        id: "3",
        type: "subTitle",
        idx: 2,
        text: "Kept text",
        box: px(SUBTITLE_BOX),
      },
      { id: "4", type: "pic", idx: 1, text: "", box: px(RIGHT_BOX) },
    ];
    expect(summarize(pres)).toEqual(expected);
    expect(summarize(await saveAndReopen(pres))).toEqual(expected);

    result.undo();
    expect(summarize(pres)).toEqual(before);
    expect(await savedPartText(pres, slideId)).toBe(slideXmlBefore);
  });

  it("fits a picture into a content placeholder, keeping its aspect ratio", async () => {
    const pres = await openEditable([
      slidePlaceholder(2, '<p:ph type="ctrTitle"/>', "Title") +
        PICTURE_PLACEHOLDER,
    ]);
    const slideId = pres.slides[0].id;

    await applyEdit(pres, {
      type: "setSlideLayout",
      slideId,
      layoutId: TWO_CONTENT,
    });

    const picture = pres.slides[0].nodes.find((node) => node.id === "4")!;
    const [x, y, w] = px(LEFT_BOX);
    expect(picture.placeholder?.idx).toBe(1);
    expect(picture.size.w).toBeCloseTo(w);
    expect(picture.size.h).toBeCloseTo(w / 2);
    expect(picture.position.x).toBeCloseTo(x);
    expect(picture.position.y).toBeGreaterThan(y);
    expect(summarize(pres).map((node) => node.id)).toEqual(["2", "4", "5"]);
  });

  it("rejects an unknown layout and leaves the slide untouched", async () => {
    const pres = await openEditable([
      slidePlaceholder(2, '<p:ph type="ctrTitle"/>', "Title"),
    ]);
    const before = summarize(pres);

    await expect(
      applyEdit(pres, {
        type: "setSlideLayout",
        slideId: pres.slides[0].id,
        layoutId: "ppt/slideLayouts/nope.xml",
      }),
    ).rejects.toThrow(/unknown slide layout/);
    expect(pres.slides[0].layoutIndex).toBe(TITLE_SLIDE);
    expect(summarize(pres)).toEqual(before);
  });
});

describe("renderLayout", () => {
  it("renders the layout's empty placeholders with their prompts", async () => {
    const pres = await openEditable([""]);

    const { element } = renderLayout(pres, TWO_CONTENT);

    expect(
      element.querySelectorAll("[data-pptx-placeholder-empty]"),
    ).toHaveLength(3);
    expect(element.textContent).toContain("Add a short title");
  });
});
