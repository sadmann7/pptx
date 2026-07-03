import { describe, expect, it } from "vitest";

import { buildPresentation } from "../model/presentation";
import { parseZip } from "../ooxml/zip-parser";
import { renderSlide } from "../renderer/slide-renderer";
import { buildPptxWithShapes } from "./minimal-pptx";

/** Renders a single text box with the given txBody paragraphs through the full pipeline. */
async function renderTextBox(paragraphsXml: string): Promise<HTMLElement> {
  const buffer = await buildPptxWithShapes(`<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="251751" y="3062425"/><a:ext cx="7788000" cy="485100"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
</p:spPr>
<p:txBody>
<a:bodyPr anchorCtr="0" anchor="t" bIns="91425" lIns="91425" spcFirstLastPara="1" rIns="91425" wrap="square" tIns="91425">
<a:noAutofit/>
</a:bodyPr>
<a:lstStyle/>
${paragraphsXml}
</p:txBody>
</p:sp>`);

  const files = await parseZip(buffer);
  const presentation = buildPresentation(files);
  const handle = renderSlide(presentation, presentation.slides[0]);
  return handle.element;
}

/** Paragraph divs are the direct children of the text container. */
function paragraphDivs(element: HTMLElement): HTMLElement[] {
  const withText = [...element.querySelectorAll("div")].filter((div) =>
    div.textContent?.includes("MARKER"),
  );
  // The innermost div containing the marker text is a paragraph div; its
  // parent is the text container holding all paragraphs.
  const paraDiv = withText[withText.length - 1];
  if (!paraDiv?.parentElement) throw new Error("marker paragraph not found");
  return [...paraDiv.parentElement.children] as HTMLElement[];
}

const EMPTY_PPR = `<a:pPr indent="0" lvl="0" marL="0" rtl="0" algn="l"><a:buNone/></a:pPr>`;

describe("empty paragraph rendering", () => {
  it("gives empty-text-run paragraphs a line box at the endParaRPr size (Google Slides spacers)", async () => {
    // Regression: 4 spacer paragraphs with an empty <a:t/> run used to collapse
    // to zero height, letting the following text overlap content above it.
    const spacer = `<a:p>${EMPTY_PPR}<a:r><a:t></a:t></a:r><a:endParaRPr b="1" sz="2000"/></a:p>`;
    const element = await renderTextBox(
      `${spacer.repeat(4)}<a:p>${EMPTY_PPR}<a:r><a:rPr b="1" lang="en" sz="2000"/><a:t>MARKER July 3, 2025</a:t></a:r><a:endParaRPr sz="2000"/></a:p>`,
    );

    const paragraphs = paragraphDivs(element);
    expect(paragraphs).toHaveLength(5);

    for (const spacerDiv of paragraphs.slice(0, 4)) {
      // Occupies a line…
      expect(spacerDiv.querySelector("br")).not.toBeNull();
      // …at the endParaRPr-declared 20pt, not an inherited default.
      expect(spacerDiv.style.fontSize).toBe("20pt");
    }
    expect(paragraphs[4].textContent).toContain("July 3, 2025");
  });

  it("keeps the line box for paragraphs with no runs at all", async () => {
    const element = await renderTextBox(
      `<a:p>${EMPTY_PPR}<a:endParaRPr sz="1400"/></a:p>
       <a:p>${EMPTY_PPR}<a:r><a:t>MARKER</a:t></a:r></a:p>`,
    );

    const paragraphs = paragraphDivs(element);
    expect(paragraphs).toHaveLength(2);
    expect(paragraphs[0].querySelector("br")).not.toBeNull();
    expect(paragraphs[0].style.fontSize).toBe("14pt");
  });

  it("does not double-count paragraphs that only contain an explicit line break", async () => {
    const element = await renderTextBox(
      `<a:p>${EMPTY_PPR}<a:br/></a:p>
       <a:p>${EMPTY_PPR}<a:r><a:t>MARKER</a:t></a:r></a:p>`,
    );

    const paragraphs = paragraphDivs(element);
    // The <a:br/> run produces its own line box; the empty-paragraph fallback
    // must not add a second one.
    expect(paragraphs[0].querySelectorAll("br")).toHaveLength(1);
  });

  it("adds no <br> to paragraphs with visible text", async () => {
    const element = await renderTextBox(
      `<a:p>${EMPTY_PPR}<a:r><a:t>MARKER visible</a:t></a:r></a:p>`,
    );

    const paragraphs = paragraphDivs(element);
    expect(paragraphs).toHaveLength(1);
    expect(paragraphs[0].querySelector("br")).toBeNull();
    expect(paragraphs[0].textContent).toContain("visible");
  });

  it("prefers the first run's explicit size over endParaRPr for the line strut", async () => {
    const element = await renderTextBox(
      `<a:p>${EMPTY_PPR}<a:r><a:rPr sz="3200"/><a:t></a:t></a:r><a:endParaRPr sz="1000"/></a:p>
       <a:p>${EMPTY_PPR}<a:r><a:t>MARKER</a:t></a:r></a:p>`,
    );

    const paragraphs = paragraphDivs(element);
    expect(paragraphs[0].style.fontSize).toBe("32pt");
  });
});
