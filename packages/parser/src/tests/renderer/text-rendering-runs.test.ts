/**
 * Run-level formatting (a:rPr) through the real render pipeline:
 * bold/italic/underline/strike, font size, fill color, and typefaces.
 */
import { describe, expect, it } from "vitest";

import {
  normalizeColor,
  paragraphsOf,
  renderTextBox,
  spansOf,
} from "../fixtures/text-rendering-helpers";

/** Renders one paragraph and returns the span for its single run. */
async function renderRun(rPrXml: string, text = "MARKER"): Promise<HTMLElement> {
  const element = await renderTextBox(`<a:p><a:r>${rPrXml}<a:t>${text}</a:t></a:r></a:p>`);
  const [paragraph] = paragraphsOf(element);
  const span = spansOf(paragraph).find((s) => s.textContent?.includes(text));
  if (!span) throw new Error("run span not found");
  return span;
}

describe("run text emphasis", () => {
  it('maps b="1" to font-weight bold', async () => {
    const span = await renderRun(`<a:rPr b="1"/>`);
    expect(span.style.fontWeight).toBe("bold");
  });

  it("leaves font-weight unset without b", async () => {
    const span = await renderRun(`<a:rPr/>`);
    expect(span.style.fontWeight).toBe("");
  });

  it('maps b="0" to no bold', async () => {
    const span = await renderRun(`<a:rPr b="0"/>`);
    expect(span.style.fontWeight).toBe("");
  });

  it('maps i="1" to font-style italic', async () => {
    const span = await renderRun(`<a:rPr i="1"/>`);
    expect(span.style.fontStyle).toBe("italic");
  });

  it('maps u="sng" to text-decoration underline', async () => {
    const span = await renderRun(`<a:rPr u="sng"/>`);
    expect(span.style.textDecoration).toContain("underline");
  });

  it('suppresses underline for u="none"', async () => {
    const span = await renderRun(`<a:rPr u="none"/>`);
    expect(span.style.textDecoration).not.toContain("underline");
  });

  it('maps strike="sngStrike" to line-through', async () => {
    const span = await renderRun(`<a:rPr strike="sngStrike"/>`);
    expect(span.style.textDecoration).toContain("line-through");
  });

  it('suppresses strike for strike="noStrike"', async () => {
    const span = await renderRun(`<a:rPr strike="noStrike"/>`);
    expect(span.style.textDecoration).not.toContain("line-through");
  });

  it("combines underline and strike into one text-decoration", async () => {
    const span = await renderRun(`<a:rPr u="sng" strike="sngStrike"/>`);
    expect(span.style.textDecoration).toMatch(/underline\s+line-through/);
  });
});

describe("run font size", () => {
  it("converts sz (hundredths of a point) to pt", async () => {
    const span = await renderRun(`<a:rPr sz="2400"/>`);
    expect(span.style.fontSize).toBe("24pt");
  });

  it("supports fractional point sizes", async () => {
    const span = await renderRun(`<a:rPr sz="1050"/>`);
    expect(parseFloat(span.style.fontSize)).toBeCloseTo(10.5, 3);
    expect(span.style.fontSize.endsWith("pt")).toBe(true);
  });

  it("defaults to 12pt when no size is specified anywhere", async () => {
    const span = await renderRun(`<a:rPr/>`);
    expect(span.style.fontSize).toBe("12pt");
  });
});

describe("run color", () => {
  it("applies solidFill srgbClr as the text color", async () => {
    const span = await renderRun(
      `<a:rPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr>`,
    );
    expect(normalizeColor(span.style.color)).toBe("255,0,0,1");
  });

  it("resolves solidFill schemeClr through the theme (accent1)", async () => {
    const span = await renderRun(
      `<a:rPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:rPr>`,
    );
    // Fixture theme accent1 = 4472C4.
    expect(normalizeColor(span.style.color)).toBe("68,114,196,1");
  });

  it("applies alpha as an rgba color", async () => {
    const span = await renderRun(
      `<a:rPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:rPr>`,
    );
    expect(span.style.color).toMatch(/rgba\(\s*255\s*,\s*0\s*,\s*0\s*,\s*0?\.50*\s*\)/);
  });

  it("defaults to black when no color is specified", async () => {
    const span = await renderRun(`<a:rPr/>`);
    expect(normalizeColor(span.style.color)).toBe("0,0,0,1");
  });
});

describe("run typeface", () => {
  it("applies a:latin typeface as the font family", async () => {
    const span = await renderRun(`<a:rPr><a:latin typeface="Georgia"/></a:rPr>`);
    expect(span.style.fontFamily).toContain("Georgia");
    // Serif face gets a serif generic fallback.
    expect(span.style.fontFamily).toMatch(/serif\s*$/);
  });

  it("resolves +mn-lt to the theme minor latin font", async () => {
    const span = await renderRun(`<a:rPr><a:latin typeface="+mn-lt"/></a:rPr>`);
    expect(span.style.fontFamily).toContain("Calibri");
  });

  it("falls back to the theme minor font when no typeface is given", async () => {
    const span = await renderRun(`<a:rPr/>`);
    expect(span.style.fontFamily).toContain("Calibri");
  });

  it("keeps latin + east-asian typefaces as a fallback stack", async () => {
    const span = await renderRun(
      `<a:rPr><a:latin typeface="Arial"/><a:ea typeface="Microsoft YaHei"/></a:rPr>`,
    );
    const families = span.style.fontFamily;
    expect(families).toContain("Arial");
    expect(families).toContain("Microsoft YaHei");
    expect(families.indexOf("Arial")).toBeLessThan(families.indexOf("Microsoft YaHei"));
  });
});

describe("additional run properties", () => {
  it("maps a:highlight to background-color", async () => {
    const span = await renderRun(
      `<a:rPr><a:highlight><a:srgbClr val="FFFF00"/></a:highlight></a:rPr>`,
    );
    expect(normalizeColor(span.style.backgroundColor)).toBe("255,255,0,1");
  });

  it("maps spc (hundredths of pt) to letter-spacing", async () => {
    const span = await renderRun(`<a:rPr spc="300"/>`);
    expect(span.style.letterSpacing).toBe("3pt");
  });

  it('maps cap="all" to text-transform uppercase', async () => {
    const span = await renderRun(`<a:rPr cap="all"/>`);
    expect(span.style.textTransform).toBe("uppercase");
  });

  it('maps cap="small" to font-variant small-caps', async () => {
    const span = await renderRun(`<a:rPr cap="small"/>`);
    expect(span.style.fontVariant).toBe("small-caps");
  });

  it("renders superscript baseline as vertical-align with a reduced font size", async () => {
    const span = await renderRun(`<a:rPr sz="2000" baseline="30000"/>`);
    expect(span.style.verticalAlign).toBe("30%");
    // Shift ≥ 20% also shrinks the glyphs to 65%.
    expect(parseFloat(span.style.fontSize)).toBeCloseTo(20 * 0.65, 3);
  });

  it("renders subscript baseline as a negative vertical-align", async () => {
    const span = await renderRun(`<a:rPr sz="2000" baseline="-25000"/>`);
    expect(span.style.verticalAlign).toBe("-25%");
  });
});

describe("multiple runs in one paragraph", () => {
  it("renders each run as its own span with independent formatting", async () => {
    const element = await renderTextBox(
      `<a:p>
        <a:r><a:rPr b="1" sz="2400"/><a:t>bold </a:t></a:r>
        <a:r><a:rPr i="1"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:rPr><a:t>green-italic </a:t></a:r>
        <a:r><a:rPr u="sng"/><a:t>underlined</a:t></a:r>
      </a:p>`,
    );
    const [paragraph] = paragraphsOf(element);
    const spans = spansOf(paragraph);
    expect(spans).toHaveLength(3);

    expect(spans[0].textContent).toBe("bold ");
    expect(spans[0].style.fontWeight).toBe("bold");
    expect(spans[0].style.fontSize).toBe("24pt");
    expect(spans[0].style.fontStyle).toBe("");

    expect(spans[1].textContent).toBe("green-italic ");
    expect(spans[1].style.fontStyle).toBe("italic");
    expect(normalizeColor(spans[1].style.color)).toBe("0,255,0,1");
    expect(spans[1].style.fontWeight).toBe("");

    expect(spans[2].textContent).toBe("underlined");
    expect(spans[2].style.textDecoration).toContain("underline");
    expect(spans[2].style.fontWeight).toBe("");
    // The full text is preserved in document order.
    expect(paragraph.textContent).toBe("bold green-italic underlined");
  });

  it("inherits pPr defRPr into runs without their own rPr", async () => {
    const element = await renderTextBox(
      `<a:p>
        <a:pPr><a:defRPr sz="2800" b="1"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:defRPr></a:pPr>
        <a:r><a:t>inherited</a:t></a:r>
        <a:r><a:rPr sz="1400" b="0"/><a:t>overridden</a:t></a:r>
      </a:p>`,
    );
    const [paragraph] = paragraphsOf(element);
    const spans = spansOf(paragraph);
    expect(spans[0].style.fontSize).toBe("28pt");
    expect(spans[0].style.fontWeight).toBe("bold");
    expect(normalizeColor(spans[0].style.color)).toBe("0,0,255,1");
    // Explicit run rPr wins over the paragraph default…
    expect(spans[1].style.fontSize).toBe("14pt");
    expect(spans[1].style.fontWeight).toBe("");
    // …but properties it does not set still inherit.
    expect(normalizeColor(spans[1].style.color)).toBe("0,0,255,1");
  });
});
