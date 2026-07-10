import { describe, expect, it } from "vitest";

import { buildPresentation } from "../../model/presentation";
import { parseZip } from "../../ooxml/zip";
import { renderSlide } from "../../renderer/slide";
import { buildPptxWithShapes } from "../fixtures/minimal-pptx";

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

/** Returns all run spans inside the paragraph that contains `marker`. */
function runSpansContaining(element: HTMLElement, marker: string): HTMLElement[] {
  const divs = [...element.querySelectorAll<HTMLElement>("div")];
  const para = divs.find((d) => d.textContent?.includes(marker));
  if (!para) throw new Error(`paragraph containing "${marker}" not found`);
  return [...para.querySelectorAll<HTMLElement>("span")].filter(
    (s) => !s.querySelector("span"), // leaf spans only
  );
}

describe("leading whitespace preservation", () => {
  it("applies white-space:pre-wrap to a run that starts with a space", async () => {
    const element = await renderTextBox(
      `<a:p>${EMPTY_PPR}<a:r><a:t>  MARKER indented</a:t></a:r></a:p>`,
    );

    const [span] = runSpansContaining(element, "MARKER indented");
    expect(span).toBeDefined();
    expect(span.textContent).toBe("  MARKER indented");
    expect(span.style.whiteSpace).toBe("pre-wrap");
  });

  it("does not apply pre-wrap to a mid-line run that starts with a space", async () => {
    // Two runs on the same paragraph: the second begins with a space but is
    // not at line start — default whitespace handling applies.
    const element = await renderTextBox(
      `<a:p>${EMPTY_PPR}<a:r><a:t>MARKER</a:t></a:r><a:r><a:t> continues</a:t></a:r></a:p>`,
    );

    const spans = runSpansContaining(element, "MARKER");
    expect(spans).toHaveLength(2);
    // First run: no special whitespace needed.
    expect(spans[0].style.whiteSpace).toBe("");
    // Second run: space but NOT at line start — no pre-wrap.
    expect(spans[1].style.whiteSpace).toBe("");
  });

  it("applies pre-wrap to the first run after an explicit line break", async () => {
    // A <a:br/> resets the line-start flag; the run after it begins a new
    // visual line and its leading space must be preserved.
    const element = await renderTextBox(
      `<a:p>${EMPTY_PPR}<a:r><a:t>MARKER first</a:t></a:r><a:br/><a:r><a:t>  second line</a:t></a:r></a:p>`,
    );

    const spans = runSpansContaining(element, "MARKER first");
    // First run: no pre-wrap.
    expect(spans[0].style.whiteSpace).toBe("");
    // Run after <br>: leading space → pre-wrap.
    expect(spans[1].style.whiteSpace).toBe("pre-wrap");
    expect(spans[1].textContent).toBe("  second line");
  });

  it("still uses white-space:pre for runs that contain tabs", async () => {
    const element = await renderTextBox(
      `<a:p>${EMPTY_PPR}<a:r><a:t>&#x9;MARKER tabbed</a:t></a:r></a:p>`,
    );

    const [span] = runSpansContaining(element, "MARKER tabbed");
    expect(span).toBeDefined();
    expect(span.style.whiteSpace).toBe("pre");
  });
});
