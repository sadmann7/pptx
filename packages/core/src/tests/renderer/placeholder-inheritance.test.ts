import { describe, expect, it } from "vitest";

import { buildPresentation } from "../../model/presentation";
import { emuToPx } from "../../ooxml/unit";
import { readPptx } from "../../ooxml/zip";
import { renderSlide } from "../../renderer/slide";
import { buildCustomPptx } from "../fixtures/fixture-extras";
import { tinyPngBytes } from "../fixtures/rich-pptx";

const LAYOUT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`;

const BOX = `<a:xfrm><a:off x="4373880" y="0"/><a:ext cx="7818120" cy="6858000"/></a:xfrm>`;
const PICTURE_FILL = `<a:blipFill><a:blip r:embed="rId5"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`;
const SOLID_FILL = `<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>`;

function placeholder(
  id: number,
  ph: string,
  spPr: string,
  prompt = "",
): string {
  return `<p:sp>
<p:nvSpPr><p:cNvPr id="${id}" name="Placeholder ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${ph}</p:nvPr></p:nvSpPr>
${spPr}
<p:txBody><a:bodyPr/><a:lstStyle/><a:p>${prompt ? `<a:r><a:t>${prompt}</a:t></a:r>` : ""}<a:endParaRPr lang="en-US"/></a:p></p:txBody>
</p:sp>`;
}

const PICTURE_PH = '<p:ph type="pic" sz="quarter" idx="13"/>';

// No prstGeom, as PowerPoint writes picture placeholders.
async function renderPlaceholder(
  layoutFill: string,
  slideSpPr: string,
): Promise<HTMLElement> {
  const buffer = await buildCustomPptx({
    slides: [placeholder(3, PICTURE_PH, slideSpPr)],
    layoutShapesXml: placeholder(
      2,
      PICTURE_PH,
      `<p:spPr>${BOX}${layoutFill}</p:spPr>`,
    ),
    extraFiles: {
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels": LAYOUT_RELS,
      "ppt/media/image1.png": tinyPngBytes(),
    },
  });
  const presentation = buildPresentation(await readPptx(buffer));
  const { element, ready } = renderSlide(presentation, presentation.slides[0]);
  await ready;
  return element;
}

describe("placeholder fill inheritance", () => {
  it("paints the layout placeholder's picture fill on an empty slide placeholder", async () => {
    const element = await renderPlaceholder(PICTURE_FILL, "<p:spPr/>");

    const image = element.querySelector("svg image");
    expect(image?.getAttribute("href")).toMatch(/^blob:/);
  });

  it("keeps the slide placeholder's own fill", async () => {
    const element = await renderPlaceholder(
      PICTURE_FILL,
      `<p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>`,
    );

    expect(element.querySelector("svg image")).toBeNull();
    expect(element.querySelector("svg path")?.getAttribute("fill")).toMatch(
      /ff0000|255, 0, 0/i,
    );
  });

  it("keeps the slide placeholder's own effects with the layout's fill", async () => {
    const element = await renderPlaceholder(
      SOLID_FILL,
      `<p:spPr><a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst></p:spPr>`,
    );

    expect(element.querySelector("svg path")?.getAttribute("fill")).toMatch(
      /00ff00|0, 255, 0/i,
    );
    expect(element.querySelector("svg filter")).not.toBeNull();
  });
});

describe("placeholder template lookup", () => {
  it("takes an untyped placeholder's prompt from the layout placeholder with the same idx", async () => {
    // The slide placeholder inherits `body` from the master, which matches the
    // section placeholder's type.
    const buffer = await buildCustomPptx({
      slides: [placeholder(4, '<p:ph idx="1"/>', "<p:spPr/>")],
      masterShapesXml: placeholder(
        2,
        '<p:ph type="body" idx="1"/>',
        `<p:spPr>${BOX}</p:spPr>`,
      ),
      layoutShapesXml:
        placeholder(
          2,
          '<p:ph type="body" sz="quarter" idx="13" hasCustomPrompt="1"/>',
          `<p:spPr>${BOX}</p:spPr>`,
          "Add a section name",
        ) +
        placeholder(
          3,
          '<p:ph idx="1" hasCustomPrompt="1"/>',
          `<p:spPr>${BOX}</p:spPr>`,
          "Add text",
        ),
    });
    const presentation = buildPresentation(await readPptx(buffer));

    const { element } = renderSlide(presentation, presentation.slides[0], {
      placeholderPrompts: true,
    });

    expect(
      [...element.querySelectorAll("[data-pptx-placeholder-prompt]")].map(
        (prompt) => prompt.textContent,
      ),
    ).toEqual(["Add text"]);
  });

  it("takes position and fill from the same layout placeholder", async () => {
    const buffer = await buildCustomPptx({
      slides: [placeholder(4, '<p:ph type="body" idx="2"/>', "<p:spPr/>")],
      layoutShapesXml:
        placeholder(
          2,
          '<p:ph type="body" idx="1"/>',
          `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="1828800"/></a:xfrm><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>`,
        ) +
        placeholder(
          3,
          '<p:ph type="pic" idx="2"/>',
          `<p:spPr>${BOX}${SOLID_FILL}</p:spPr>`,
        ),
    });
    const presentation = buildPresentation(await readPptx(buffer));
    const [node] = presentation.slides[0].nodes;

    const { element } = renderSlide(presentation, presentation.slides[0]);

    expect(node.position.x).toBe(emuToPx(4373880));
    expect(element.querySelector("svg path")?.getAttribute("fill")).toMatch(
      /00ff00|0, 255, 0/i,
    );
  });
});
