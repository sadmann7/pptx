/**
 * Text body properties (a:bodyPr), field runs (a:fld) and hyperlink runs
 * (a:hlinkClick ppaction) through the real render pipeline.
 */
import { describe, expect, it, vi } from "vitest";

import {
  normalizeColor,
  paragraphsOf,
  renderTextBox,
  spansOf,
  textContainerOf,
} from "../fixtures/text-rendering-helpers";

const MARKER_PARAGRAPH = `<a:p><a:r><a:t>MARKER</a:t></a:r></a:p>`;

async function renderBody(bodyPrXml: string, paragraphsXml = MARKER_PARAGRAPH) {
  const element = await renderTextBox(paragraphsXml, { bodyPrXml });
  return textContainerOf(element);
}

describe("bodyPr anchor (vertical alignment)", () => {
  it.each([
    ["t", "flex-start"],
    ["ctr", "center"],
    ["b", "flex-end"],
  ] as const)("maps anchor=%s to justify-content %s", async (anchor, expected) => {
    const container = await renderBody(`<a:bodyPr anchor="${anchor}"><a:noAutofit/></a:bodyPr>`);
    expect(container.style.justifyContent).toBe(expected);
  });

  it("defaults to flex-start without an anchor", async () => {
    const container = await renderBody(`<a:bodyPr><a:noAutofit/></a:bodyPr>`);
    expect(container.style.justifyContent).toBe("flex-start");
  });
});

describe("bodyPr wrap", () => {
  it('maps wrap="none" to white-space nowrap', async () => {
    const container = await renderBody(`<a:bodyPr wrap="none"><a:noAutofit/></a:bodyPr>`);
    expect(container.style.whiteSpace).toBe("nowrap");
  });

  it('keeps normal wrapping for wrap="square"', async () => {
    const container = await renderBody(`<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>`);
    expect(container.style.whiteSpace).not.toBe("nowrap");
  });
});

describe("bodyPr insets", () => {
  it("maps lIns/tIns/rIns/bIns (EMU) to padding in px", async () => {
    const container = await renderBody(
      `<a:bodyPr lIns="182880" tIns="91440" rIns="0" bIns="45720"><a:noAutofit/></a:bodyPr>`,
    );
    expect(parseFloat(container.style.paddingLeft)).toBeCloseTo(19.2, 3);
    expect(parseFloat(container.style.paddingTop)).toBeCloseTo(9.6, 3);
    expect(parseFloat(container.style.paddingRight)).toBeCloseTo(0, 3);
    expect(parseFloat(container.style.paddingBottom)).toBeCloseTo(4.8, 3);
  });

  it("uses the OOXML default insets when none are given", async () => {
    const container = await renderBody(`<a:bodyPr><a:noAutofit/></a:bodyPr>`);
    // Defaults: 91440 EMU (0.1") left/right, 45720 EMU (0.05") top/bottom.
    expect(parseFloat(container.style.paddingLeft)).toBeCloseTo(9.6, 3);
    expect(parseFloat(container.style.paddingTop)).toBeCloseTo(4.8, 3);
    expect(parseFloat(container.style.paddingRight)).toBeCloseTo(9.6, 3);
    expect(parseFloat(container.style.paddingBottom)).toBeCloseTo(4.8, 3);
  });
});

describe("bodyPr vertical text (vert)", () => {
  it("renders eaVert as vertical-rl writing mode", async () => {
    const container = await renderBody(`<a:bodyPr vert="eaVert"><a:noAutofit/></a:bodyPr>`);
    expect(container.style.writingMode).toBe("vertical-rl");
    expect(container.style.justifyContent).toBe("center");
    // Paragraphs avoid mid-word breaks in vertical flow.
    const [paragraph] = paragraphsOf(container);
    expect(paragraph.style.wordBreak).toBe("keep-all");
  });

  it("renders vert270 as vertical-rl rotated 180 degrees", async () => {
    const container = await renderBody(`<a:bodyPr vert="vert270"><a:noAutofit/></a:bodyPr>`);
    expect(container.style.writingMode).toBe("vertical-rl");
    expect(container.style.transform).toContain("rotate(180deg)");
  });

  it("maps the anchor onto align-items in vertical flow", async () => {
    const container = await renderBody(
      `<a:bodyPr vert="eaVert" anchor="b"><a:noAutofit/></a:bodyPr>`,
    );
    expect(container.style.alignItems).toBe("flex-end");
  });

  it("renders wordArtVert as upright vertical-lr text", async () => {
    const container = await renderBody(`<a:bodyPr vert="wordArtVert"><a:noAutofit/></a:bodyPr>`);
    expect(container.style.writingMode).toBe("vertical-lr");
    expect(container.style.textOrientation).toBe("upright");
  });
});

describe("bodyPr normAutofit", () => {
  it("applies fontScale to run and paragraph font sizes", async () => {
    const container = await renderBody(
      `<a:bodyPr><a:normAutofit fontScale="62500"/></a:bodyPr>`,
      `<a:p><a:r><a:rPr sz="3200"/><a:t>MARKER</a:t></a:r></a:p>`,
    );
    const [paragraph] = paragraphsOf(container);
    const run = spansOf(paragraph).find((s) => s.textContent === "MARKER");
    // 32pt × 62.5% = 20pt.
    expect(parseFloat(run!.style.fontSize)).toBeCloseTo(20, 3);
    expect(parseFloat(paragraph.style.fontSize)).toBeCloseTo(20, 3);
  });

  it("applies lnSpcReduction to the paragraph line-height", async () => {
    const container = await renderBody(
      `<a:bodyPr><a:normAutofit fontScale="100000" lnSpcReduction="20000"/></a:bodyPr>`,
      `<a:p><a:pPr><a:lnSpc><a:spcPct val="100000"/></a:lnSpc></a:pPr><a:r><a:rPr sz="1800"/><a:t>MARKER</a:t></a:r></a:p>`,
    );
    const [paragraph] = paragraphsOf(container);
    // 1.0 single spacing reduced by 20%.
    expect(parseFloat(paragraph.style.lineHeight)).toBeCloseTo(0.8, 3);
  });

  it("scales bullets along with the text", async () => {
    const container = await renderBody(
      `<a:bodyPr><a:normAutofit fontScale="50000"/></a:bodyPr>`,
      `<a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:rPr sz="2000"/><a:t>MARKER</a:t></a:r></a:p>`,
    );
    const [paragraph] = paragraphsOf(container);
    const bullet = spansOf(paragraph).find((s) => s.textContent === "• ");
    expect(parseFloat(bullet!.style.fontSize)).toBeCloseTo(10, 3);
  });
});

describe("field runs (a:fld)", () => {
  // TODO(spec?): ECMA-376 fields like slidenum should show the current slide
  // number; the renderer emits the cached <a:t> literal instead of recomputing
  // it from the slide index. Asserting actual behavior here.
  it("renders the cached literal text of a slidenum field", async () => {
    const element = await renderTextBox(
      `<a:p><a:r><a:t>Slide </a:t></a:r><a:fld id="{6D2CF6E5-4568-4E3F-9D3E-42AB6B0C5E11}" type="slidenum"><a:rPr b="1"/><a:t>7</a:t></a:fld></a:p>`,
    );
    const [paragraph] = paragraphsOf(element);
    expect(paragraph.textContent).toBe("Slide 7");
    // Field run properties are applied like regular run properties.
    const fieldSpan = spansOf(paragraph).find((s) => s.textContent === "7");
    expect(fieldSpan?.style.fontWeight).toBe("bold");
  });

  it("keeps fields interleaved with runs in document order", async () => {
    const element = await renderTextBox(
      `<a:p><a:fld id="{6D2CF6E5-4568-4E3F-9D3E-42AB6B0C5E11}" type="slidenum"><a:t>1</a:t></a:fld><a:r><a:t> of 10</a:t></a:r></a:p>`,
    );
    expect(paragraphsOf(element)[0].textContent).toBe("1 of 10");
  });
});

describe("hyperlink runs (ppaction, no rels required)", () => {
  const LINK_PARAGRAPH = `<a:p><a:r><a:rPr><a:hlinkClick action="ppaction://hlinkshowjump?jump=firstslide"/></a:rPr><a:t>MARKER</a:t></a:r></a:p>`;

  it("renders a slide-jump action as an accessible link span", async () => {
    const element = await renderTextBox(LINK_PARAGRAPH, {
      renderOptions: { onNavigate: () => undefined },
    });
    const [paragraph] = paragraphsOf(element);
    const link = spansOf(paragraph).find((s) => s.textContent === "MARKER")!;
    expect(link.getAttribute("role")).toBe("link");
    expect(link.tabIndex).toBe(0);
    expect(link.title).toBe("Go to slide 1");
    expect(link.style.cursor).toBe("pointer");
    // Hyperlink runs default to underlined…
    expect(link.style.textDecoration).toContain("underline");
    // …in the theme hlink color (fixture hlink = 0563C1).
    expect(normalizeColor(link.style.color)).toBe("5,99,193,1");
  });

  it("invokes onNavigate with the target slide index on click", async () => {
    const onNavigate = vi.fn();
    const element = await renderTextBox(LINK_PARAGRAPH, { renderOptions: { onNavigate } });
    const link = spansOf(paragraphsOf(element)[0]).find((s) => s.textContent === "MARKER")!;
    link.click();
    expect(onNavigate).toHaveBeenCalledWith({ slideIndex: 0 });
  });

  it("invokes onNavigate on Enter for keyboard users", async () => {
    const onNavigate = vi.fn();
    const element = await renderTextBox(LINK_PARAGRAPH, { renderOptions: { onNavigate } });
    const link = spansOf(paragraphsOf(element)[0]).find((s) => s.textContent === "MARKER")!;
    link.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
    expect(onNavigate).toHaveBeenCalledWith({ slideIndex: 0 });
  });

  it("uses the hlinkClick tooltip as the link title when present", async () => {
    const element = await renderTextBox(
      `<a:p><a:r><a:rPr><a:hlinkClick action="ppaction://hlinkshowjump?jump=lastslide" tooltip="Jump to end"/></a:rPr><a:t>MARKER</a:t></a:r></a:p>`,
      { renderOptions: { onNavigate: () => undefined } },
    );
    const link = spansOf(paragraphsOf(element)[0]).find((s) => s.textContent === "MARKER")!;
    expect(link.title).toBe("Jump to end");
  });

  it("renders a plain span when no onNavigate handler is provided", async () => {
    const element = await renderTextBox(LINK_PARAGRAPH);
    const span = spansOf(paragraphsOf(element)[0]).find((s) => s.textContent === "MARKER")!;
    expect(span.getAttribute("role")).toBeNull();
    expect(span.style.cursor).not.toBe("pointer");
  });

  it("renders a plain span for unreachable jumps (nextslide on the last slide)", async () => {
    const element = await renderTextBox(
      `<a:p><a:r><a:rPr><a:hlinkClick action="ppaction://hlinkshowjump?jump=nextslide"/></a:rPr><a:t>MARKER</a:t></a:r></a:p>`,
      { renderOptions: { onNavigate: () => undefined } },
    );
    const span = spansOf(paragraphsOf(element)[0]).find((s) => s.textContent === "MARKER")!;
    // Single-slide deck: there is no next slide to jump to.
    expect(span.getAttribute("role")).toBeNull();
  });

  it("keeps an explicit run color instead of the hlink theme color", async () => {
    const element = await renderTextBox(
      `<a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:hlinkClick action="ppaction://hlinkshowjump?jump=firstslide"/></a:rPr><a:t>MARKER</a:t></a:r></a:p>`,
      { renderOptions: { onNavigate: () => undefined } },
    );
    const link = spansOf(paragraphsOf(element)[0]).find((s) => s.textContent === "MARKER")!;
    expect(normalizeColor(link.style.color)).toBe("255,0,0,1");
  });
});
