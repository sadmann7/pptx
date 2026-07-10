/**
 * Paragraph-level formatting (a:pPr) through the real render pipeline:
 * alignment, line spacing, space before/after, indentation, line breaks,
 * tabs and the trailing endParaRPr spacer.
 */
import { describe, expect, it } from "vitest";

import { paragraphsOf, renderTextBox, spansOf } from "../helpers/text-rendering-helpers";

/** Renders a single paragraph with the given pPr and returns its div. */
async function renderParagraph(pPrXml: string, runsXml?: string): Promise<HTMLElement> {
  const element = await renderTextBox(
    `<a:p>${pPrXml}${runsXml ?? "<a:r><a:t>MARKER</a:t></a:r>"}</a:p>`,
  );
  return paragraphsOf(element)[0];
}

describe("paragraph alignment", () => {
  it.each([
    ["l", "left"],
    ["ctr", "center"],
    ["r", "right"],
    ["just", "justify"],
  ] as const)("maps algn=%s to text-align %s", async (algn, expected) => {
    const paragraph = await renderParagraph(`<a:pPr algn="${algn}"/>`);
    expect(paragraph.style.textAlign).toBe(expected);
  });

  it("leaves text-align unset without algn", async () => {
    const paragraph = await renderParagraph(`<a:pPr/>`);
    expect(paragraph.style.textAlign).toBe("");
  });

  it('maps rtl="1" to direction rtl', async () => {
    const paragraph = await renderParagraph(`<a:pPr rtl="1"/>`);
    expect(paragraph.style.direction).toBe("rtl");
  });
});

describe("line spacing", () => {
  it("maps spcPct to a unitless line-height (150000 → 1.5)", async () => {
    const paragraph = await renderParagraph(
      `<a:pPr><a:lnSpc><a:spcPct val="150000"/></a:lnSpc></a:pPr>`,
    );
    expect(parseFloat(paragraph.style.lineHeight)).toBeCloseTo(1.5, 3);
    // Unitless — CSS percentage would be inherited as a fixed px value.
    expect(paragraph.style.lineHeight).not.toMatch(/%|pt|px/);
  });

  it("maps spcPct 100000 to single spacing (1.0)", async () => {
    const paragraph = await renderParagraph(
      `<a:pPr><a:lnSpc><a:spcPct val="100000"/></a:lnSpc></a:pPr>`,
    );
    expect(parseFloat(paragraph.style.lineHeight)).toBeCloseTo(1, 3);
  });

  it("maps spcPts (hundredths of pt) to an absolute pt line-height", async () => {
    const paragraph = await renderParagraph(
      `<a:pPr><a:lnSpc><a:spcPts val="2400"/></a:lnSpc></a:pPr>`,
    );
    expect(paragraph.style.lineHeight).toBe("24pt");
  });
});

describe("space before / after paragraphs", () => {
  it("maps spcBef spcPts to margin-top in pt", async () => {
    const paragraph = await renderParagraph(
      `<a:pPr><a:spcBef><a:spcPts val="1200"/></a:spcBef></a:pPr>`,
    );
    expect(paragraph.style.marginTop).toBe("12pt");
  });

  it("maps spcAft spcPts to margin-bottom in pt", async () => {
    const paragraph = await renderParagraph(
      `<a:pPr><a:spcAft><a:spcPts val="600"/></a:spcAft></a:pPr>`,
    );
    expect(paragraph.style.marginBottom).toBe("6pt");
  });

  it("resolves spcBef spcPct against the paragraph's effective font size", async () => {
    // 50% of the 20pt first run → 10pt.
    const paragraph = await renderParagraph(
      `<a:pPr><a:spcBef><a:spcPct val="50000"/></a:spcBef></a:pPr>`,
      `<a:r><a:rPr sz="2000"/><a:t>MARKER</a:t></a:r>`,
    );
    expect(parseFloat(paragraph.style.marginTop)).toBeCloseTo(10, 3);
    expect(paragraph.style.marginTop.endsWith("pt")).toBe(true);
  });

  it("resolves spcAft spcPct against the paragraph's effective font size", async () => {
    // 25% of the 24pt first run → 6pt.
    const paragraph = await renderParagraph(
      `<a:pPr><a:spcAft><a:spcPct val="25000"/></a:spcAft></a:pPr>`,
      `<a:r><a:rPr sz="2400"/><a:t>MARKER</a:t></a:r>`,
    );
    expect(parseFloat(paragraph.style.marginBottom)).toBeCloseTo(6, 3);
  });
});

describe("indentation", () => {
  it("maps marL (EMU) to padding-left in px", async () => {
    // 457200 EMU = 0.5 inch = 48px at 96dpi.
    const paragraph = await renderParagraph(`<a:pPr marL="457200"/>`);
    expect(parseFloat(paragraph.style.paddingLeft)).toBeCloseTo(48, 3);
  });

  it("maps positive indent to text-indent in px", async () => {
    const paragraph = await renderParagraph(`<a:pPr indent="228600"/>`);
    expect(parseFloat(paragraph.style.textIndent)).toBeCloseTo(24, 3);
  });

  it("keeps a negative (hanging) indent when there is no bullet", async () => {
    const paragraph = await renderParagraph(`<a:pPr marL="457200" indent="-228600"/>`);
    expect(parseFloat(paragraph.style.textIndent)).toBeCloseTo(-24, 3);
    expect(parseFloat(paragraph.style.paddingLeft)).toBeCloseTo(48, 3);
  });
});

describe("line breaks (a:br)", () => {
  it("renders a:br between runs as a <br>", async () => {
    const element = await renderTextBox(
      `<a:p><a:r><a:t>first</a:t></a:r><a:br/><a:r><a:t>second</a:t></a:r></a:p>`,
    );
    const [paragraph] = paragraphsOf(element);
    const spans = spansOf(paragraph);
    expect(spans.map((s) => s.textContent)).toEqual(["first", "second"]);
    expect(paragraph.querySelectorAll("br")).toHaveLength(1);
    // The <br> sits between the two run spans.
    const childTags = [...paragraph.children].map((child) => child.tagName.toLowerCase());
    expect(childTags).toEqual(["span", "br", "span"]);
  });

  it("wraps lines in fixed-height divs when line spacing is absolute (spcPts)", async () => {
    const element = await renderTextBox(
      `<a:p>
        <a:pPr><a:lnSpc><a:spcPts val="2400"/></a:lnSpc></a:pPr>
        <a:r><a:t>first</a:t></a:r><a:br/><a:r><a:t>second</a:t></a:r>
      </a:p>`,
    );
    const [paragraph] = paragraphsOf(element);
    // No <br>: each line lives in its own block wrapper with the exact height.
    expect(paragraph.querySelectorAll("br")).toHaveLength(0);
    const lines = [...paragraph.children].filter((c) => c.tagName.toLowerCase() === "div");
    expect(lines).toHaveLength(2);
    for (const line of lines as HTMLElement[]) {
      expect(line.style.height).toBe("24pt");
    }
    expect((lines[0] as HTMLElement).textContent).toBe("first");
    expect((lines[1] as HTMLElement).textContent).toBe("second");
  });

  it("adds an endParaRPr-sized spacer after a trailing line break", async () => {
    const element = await renderTextBox(
      `<a:p><a:r><a:rPr sz="1200"/><a:t>MARKER</a:t></a:r><a:br/><a:endParaRPr sz="3600"/></a:p>`,
    );
    const [paragraph] = paragraphsOf(element);
    const spacer = spansOf(paragraph).find((s) => s.textContent === "\u200B");
    expect(spacer).toBeDefined();
    expect(spacer!.style.fontSize).toBe("36pt");
  });

  it("adds no spacer when the paragraph does not end with a break", async () => {
    const element = await renderTextBox(
      `<a:p><a:r><a:rPr sz="1200"/><a:t>MARKER</a:t></a:r><a:endParaRPr sz="3600"/></a:p>`,
    );
    const [paragraph] = paragraphsOf(element);
    const spacer = spansOf(paragraph).find((s) => s.textContent === "\u200B");
    expect(spacer).toBeUndefined();
  });
});

describe("tabs", () => {
  it("renders runs containing tabs with white-space:pre and a default tab size", async () => {
    const element = await renderTextBox(`<a:p><a:r><a:t>a\tb</a:t></a:r></a:p>`);
    const [paragraph] = paragraphsOf(element);
    const [span] = spansOf(paragraph);
    expect(span.style.whiteSpace).toBe("pre");
    // Default OOXML tab = 914400 EMU = 96px.
    expect(parseFloat(paragraph.style.tabSize)).toBeCloseTo(96, 3);
  });

  it("honors pPr defTabSz for the tab size", async () => {
    const element = await renderTextBox(
      `<a:p><a:pPr defTabSz="457200"/><a:r><a:t>a\tb</a:t></a:r></a:p>`,
    );
    const [paragraph] = paragraphsOf(element);
    expect(parseFloat(paragraph.style.tabSize)).toBeCloseTo(48, 3);
  });
});

describe("paragraph line strut", () => {
  it("sets the paragraph font-size from the first run", async () => {
    const paragraph = await renderParagraph(
      `<a:pPr/>`,
      `<a:r><a:rPr sz="3000"/><a:t>MARKER</a:t></a:r>`,
    );
    expect(paragraph.style.fontSize).toBe("30pt");
  });

  it("preserves consecutive spaces with non-breaking spaces", async () => {
    const element = await renderTextBox(`<a:p><a:r><a:t>a  b</a:t></a:r></a:p>`);
    const [paragraph] = paragraphsOf(element);
    const [span] = spansOf(paragraph);
    // Space pairs become " \u00a0" so they survive whitespace collapse.
    expect(span.textContent).toBe("a \u00a0b");
  });
});
