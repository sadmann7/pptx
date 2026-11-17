import { describe, expect, it } from "vitest";

import { buildPresentation } from "../../model/presentation";
import { readPptx } from "../../ooxml/zip";
import { renderSlide } from "../../renderer/slide";
import { buildCustomPptx } from "../fixtures/fixture-extras";
import { tinyPngBytes } from "../fixtures/rich-pptx";

// No prstGeom, as PowerPoint writes picture placeholders.
const LAYOUT_PICTURE_PLACEHOLDER = `<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Picture Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="pic" sz="quarter" idx="13"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="4373880" y="0"/><a:ext cx="7818120" cy="6858000"/></a:xfrm>
<a:blipFill><a:blip r:embed="rId5"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>
</p:sp>`;

const LAYOUT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`;

function slidePicturePlaceholder(spPr: string): string {
  return `<p:sp>
<p:nvSpPr><p:cNvPr id="3" name="Picture Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="pic" sz="quarter" idx="13"/></p:nvPr></p:nvSpPr>
${spPr}
</p:sp>`;
}

async function renderPlaceholder(spPr: string): Promise<HTMLElement> {
  const buffer = await buildCustomPptx({
    slides: [slidePicturePlaceholder(spPr)],
    layoutShapesXml: LAYOUT_PICTURE_PLACEHOLDER,
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
    const element = await renderPlaceholder("<p:spPr/>");

    const image = element.querySelector("svg image");
    expect(image?.getAttribute("href")).toMatch(/^blob:/);
  });

  it("keeps the slide placeholder's own fill", async () => {
    const element = await renderPlaceholder(
      `<p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>`,
    );

    expect(element.querySelector("svg image")).toBeNull();
    expect(element.querySelector("svg path")?.getAttribute("fill")).toMatch(
      /ff0000|255, 0, 0/i,
    );
  });
});
