import { describe, expect, it } from "vitest";

import { buildPresentation } from "../../model/presentation";
import { parseZip } from "../../ooxml/zip-parser";
import { renderSlide } from "../../renderer/slide-renderer";
import { buildCustomPptx } from "../fixtures/fixture-extras";

// Layout title placeholder with prompt text, plus a promptless sldNum placeholder.
const LAYOUT_SHAPES = `<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title" hasCustomPrompt="1"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="4572000" cy="914400"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Presentation Title</a:t></a:r></a:p></p:txBody>
</p:sp>
<p:sp>
<p:nvSpPr><p:cNvPr id="3" name="Slide Number"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum" sz="quarter" idx="2"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="914400" y="5486400"/><a:ext cx="914400" cy="365760"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{00000000-0000-0000-0000-000000000000}" type="slidenum"/></a:p></p:txBody>
</p:sp>`;

// Slide placeholders with empty text bodies (ctrTitle inherits from layout title,
// sldNum from the layout sldNum).
const EMPTY_CTR_TITLE = `<p:sp>
<p:nvSpPr><p:cNvPr id="10" name="Title"/><p:cNvSpPr txBox="1"/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr>
<p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr/></a:p></p:txBody>
</p:sp>`;

const EMPTY_SLD_NUM = `<p:sp>
<p:nvSpPr><p:cNvPr id="11" name="Slide Number"/><p:cNvSpPr txBox="1"/><p:nvPr><p:ph type="sldNum" idx="2"/></p:nvPr></p:nvSpPr>
<p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr/></a:p></p:txBody>
</p:sp>`;

const FILLED_TITLE = `<p:sp>
<p:nvSpPr><p:cNvPr id="12" name="Title"/><p:cNvSpPr txBox="1"/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr>
<p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Real Title</a:t></a:r></a:p></p:txBody>
</p:sp>`;

async function renderFirstSlide(slideXml: string, placeholderPrompts: boolean) {
  const buffer = await buildCustomPptx({
    slides: [slideXml],
    layoutShapesXml: LAYOUT_SHAPES,
  });
  const files = await parseZip(buffer);
  const presentation = buildPresentation(files);
  return renderSlide(presentation, presentation.slides[0], { placeholderPrompts }).element;
}

describe("placeholder prompts (edit-mode affordance)", () => {
  it("renders layout prompt text and a dashed outline for empty placeholders", async () => {
    const element = await renderFirstSlide(EMPTY_CTR_TITLE, true);

    expect(element.textContent).toContain("Presentation Title");

    const empty = element.querySelector<HTMLElement>("[data-pptx-placeholder-empty]");
    expect(empty).not.toBeNull();
    expect(empty!.style.outline).toContain("dashed");

    const prompt = element.querySelector<HTMLElement>("[data-pptx-placeholder-prompt]");
    expect(prompt).not.toBeNull();
    expect(prompt!.style.pointerEvents).toBe("none");
  });

  it("does not render prompts by default", async () => {
    const element = await renderFirstSlide(EMPTY_CTR_TITLE, false);

    expect(element.textContent).not.toContain("Presentation Title");
    expect(element.querySelector("[data-pptx-placeholder-empty]")).toBeNull();
  });

  it("skips metadata placeholders like sldNum", async () => {
    const element = await renderFirstSlide(EMPTY_SLD_NUM, true);

    expect(element.querySelector("[data-pptx-placeholder-empty]")).toBeNull();
    expect(element.querySelector("[data-pptx-placeholder-prompt]")).toBeNull();
  });

  it("leaves placeholders with real text untouched", async () => {
    const element = await renderFirstSlide(FILLED_TITLE, true);

    expect(element.textContent).toContain("Real Title");
    expect(element.textContent).not.toContain("Presentation Title");
    expect(element.querySelector("[data-pptx-placeholder-empty]")).toBeNull();
  });
});
