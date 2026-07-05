/**
 * Bullet rendering (a:buChar / a:buAutoNum / a:buNone / a:buClr / sizing)
 * through the real render pipeline.
 */
import { describe, expect, it } from "vitest";

import { normalizeColor, paragraphsOf, renderTextBox, spansOf } from "./text-rendering-helpers";

/** The bullet span is prepended before the run spans. */
function bulletOf(paragraph: HTMLElement): HTMLElement | undefined {
  const first = paragraph.firstElementChild as HTMLElement | null;
  if (!first || first.tagName.toLowerCase() !== "span") return undefined;
  // Run spans carry the paragraph text; the bullet span carries the marker + trailing space.
  return first.textContent?.endsWith(" ") && first.textContent !== "MARKER " ? first : undefined;
}

describe("buChar bullets", () => {
  it("prepends the custom bullet character with a trailing space", async () => {
    const element = await renderTextBox(
      `<a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:t>MARKER</a:t></a:r></a:p>`,
    );
    const [paragraph] = paragraphsOf(element);
    const bullet = bulletOf(paragraph);
    expect(bullet?.textContent).toBe("• ");
    expect(paragraph.textContent).toBe("• MARKER");
  });

  it("applies buFont to the bullet span only", async () => {
    const element = await renderTextBox(
      `<a:p><a:pPr><a:buFont typeface="Wingdings"/><a:buChar char="§"/></a:pPr><a:r><a:t>MARKER</a:t></a:r></a:p>`,
    );
    const [paragraph] = paragraphsOf(element);
    const bullet = bulletOf(paragraph);
    expect(bullet?.style.fontFamily).toContain("Wingdings");
    const runSpan = spansOf(paragraph).find((s) => s.textContent === "MARKER");
    expect(runSpan?.style.fontFamily).not.toContain("Wingdings");
  });

  it("sizes the bullet at the paragraph's effective font size by default", async () => {
    const element = await renderTextBox(
      `<a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:rPr sz="2000"/><a:t>MARKER</a:t></a:r></a:p>`,
    );
    const bullet = bulletOf(paragraphsOf(element)[0]);
    expect(bullet?.style.fontSize).toBe("20pt");
  });

  it("scales the bullet with buSzPct", async () => {
    const element = await renderTextBox(
      `<a:p><a:pPr><a:buSzPct val="75000"/><a:buChar char="•"/></a:pPr><a:r><a:rPr sz="2000"/><a:t>MARKER</a:t></a:r></a:p>`,
    );
    const bullet = bulletOf(paragraphsOf(element)[0]);
    expect(parseFloat(bullet!.style.fontSize)).toBeCloseTo(15, 3);
  });

  it("sizes the bullet absolutely with buSzPts", async () => {
    const element = await renderTextBox(
      `<a:p><a:pPr><a:buSzPts val="1800"/><a:buChar char="•"/></a:pPr><a:r><a:rPr sz="2000"/><a:t>MARKER</a:t></a:r></a:p>`,
    );
    const bullet = bulletOf(paragraphsOf(element)[0]);
    expect(bullet?.style.fontSize).toBe("18pt");
  });
});

describe("buAutoNum bullets", () => {
  it("increments arabicPeriod numbering across paragraphs", async () => {
    const p = (text: string) =>
      `<a:p><a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>${text}</a:t></a:r></a:p>`;
    const element = await renderTextBox(p("one") + p("two") + p("three"));
    const bullets = paragraphsOf(element).map((para) => bulletOf(para)?.textContent);
    expect(bullets).toEqual(["1. ", "2. ", "3. "]);
  });

  it("starts numbering at startAt and continues from there", async () => {
    const element = await renderTextBox(
      `<a:p><a:pPr><a:buAutoNum type="arabicPeriod" startAt="5"/></a:pPr><a:r><a:t>one</a:t></a:r></a:p>
       <a:p><a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>two</a:t></a:r></a:p>
       <a:p><a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>three</a:t></a:r></a:p>`,
    );
    const bullets = paragraphsOf(element).map((para) => bulletOf(para)?.textContent);
    expect(bullets).toEqual(["5. ", "6. ", "7. "]);
  });

  // TODO(spec?): PowerPoint repeats startAt on every paragraph of a list that
  // starts at N and still renders N, N+1, N+2… The renderer restarts the
  // counter whenever startAt is present, so repeated startAt yields repeated
  // numbers. Asserting actual behavior here.
  it("restarts the counter whenever startAt is present (actual behavior)", async () => {
    const p = (text: string) =>
      `<a:p><a:pPr><a:buAutoNum type="arabicPeriod" startAt="5"/></a:pPr><a:r><a:t>${text}</a:t></a:r></a:p>`;
    const element = await renderTextBox(p("one") + p("two"));
    const bullets = paragraphsOf(element).map((para) => bulletOf(para)?.textContent);
    expect(bullets).toEqual(["5. ", "5. "]);
  });

  it.each([
    ["arabicParenR", "1) "],
    ["arabicParenBoth", "(1) "],
    ["arabicPlain", "1 "],
    ["alphaUcPeriod", "A. "],
    ["alphaLcPeriod", "a. "],
    ["alphaUcParenR", "A) "],
    ["alphaLcParenR", "a) "],
    ["romanUcPeriod", "I. "],
    ["romanLcPeriod", "i. "],
  ] as const)("formats %s as %s", async (type, expected) => {
    const element = await renderTextBox(
      `<a:p><a:pPr><a:buAutoNum type="${type}"/></a:pPr><a:r><a:t>MARKER</a:t></a:r></a:p>`,
    );
    expect(bulletOf(paragraphsOf(element)[0])?.textContent).toBe(expected);
  });

  it("renders incrementing roman numerals", async () => {
    const p = (text: string) =>
      `<a:p><a:pPr><a:buAutoNum type="romanUcPeriod"/></a:pPr><a:r><a:t>${text}</a:t></a:r></a:p>`;
    const element = await renderTextBox(p("a") + p("b") + p("c") + p("d"));
    const bullets = paragraphsOf(element).map((para) => bulletOf(para)?.textContent);
    expect(bullets).toEqual(["I. ", "II. ", "III. ", "IV. "]);
  });

  it("keeps independent counters per indent level", async () => {
    const p = (lvl: number, text: string) =>
      `<a:p><a:pPr${lvl > 0 ? ` lvl="${lvl}"` : ""}><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>${text}</a:t></a:r></a:p>`;
    const element = await renderTextBox(p(0, "a") + p(1, "b") + p(0, "c"));
    const bullets = paragraphsOf(element).map((para) => bulletOf(para)?.textContent);
    expect(bullets).toEqual(["1. ", "1. ", "2. "]);
  });
});

describe("bullet suppression", () => {
  it("suppresses inherited bullets with buNone", async () => {
    const lstStyle = `<a:lstStyle><a:lvl1pPr><a:buChar char="•"/></a:lvl1pPr></a:lstStyle>`;
    const element = await renderTextBox(
      `<a:p><a:r><a:t>bulleted</a:t></a:r></a:p>
       <a:p><a:pPr><a:buNone/></a:pPr><a:r><a:t>plain</a:t></a:r></a:p>`,
      { lstStyleXml: lstStyle },
    );
    const [withBullet, withoutBullet] = paragraphsOf(element);
    expect(withBullet.textContent).toBe("• bulleted");
    expect(withoutBullet.textContent).toBe("plain");
  });

  it.each(["title", "ctrTitle", "subTitle"] as const)(
    "suppresses bullets on %s placeholders",
    async (phType) => {
      const element = await renderTextBox(
        `<a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:t>MARKER</a:t></a:r></a:p>`,
        { nvPrXml: `<p:ph type="${phType}"/>` },
      );
      expect(paragraphsOf(element)[0].textContent).toBe("MARKER");
    },
  );

  it("keeps bullets on body placeholders", async () => {
    const element = await renderTextBox(
      `<a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:t>MARKER</a:t></a:r></a:p>`,
      { nvPrXml: `<p:ph type="body" idx="1"/>` },
    );
    expect(paragraphsOf(element)[0].textContent).toBe("• MARKER");
  });

  it("shows no bullet on paragraphs without visible runs", async () => {
    const element = await renderTextBox(
      `<a:p><a:pPr><a:buChar char="•"/></a:pPr><a:endParaRPr/></a:p>
       <a:p><a:r><a:t>MARKER</a:t></a:r></a:p>`,
    );
    const [empty] = paragraphsOf(element);
    expect(empty.textContent).not.toContain("•");
    // It still occupies a line box.
    expect(empty.querySelector("br")).not.toBeNull();
  });
});

describe("bullet color", () => {
  it("uses an explicit buClr", async () => {
    const element = await renderTextBox(
      `<a:p><a:pPr><a:buClr><a:srgbClr val="FF0000"/></a:buClr><a:buChar char="•"/></a:pPr>
        <a:r><a:rPr><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:rPr><a:t>MARKER</a:t></a:r></a:p>`,
    );
    const paragraph = paragraphsOf(element)[0];
    expect(normalizeColor(bulletOf(paragraph)!.style.color)).toBe("255,0,0,1");
    // The run keeps its own color.
    const run = spansOf(paragraph).find((s) => s.textContent === "MARKER");
    expect(normalizeColor(run!.style.color)).toBe("0,0,255,1");
  });

  it("resolves buClr schemeClr through the theme (accent2)", async () => {
    const element = await renderTextBox(
      `<a:p><a:pPr><a:buClr><a:schemeClr val="accent2"/></a:buClr><a:buChar char="•"/></a:pPr><a:r><a:t>MARKER</a:t></a:r></a:p>`,
    );
    // Fixture theme accent2 = ED7D31.
    expect(normalizeColor(bulletOf(paragraphsOf(element)[0])!.style.color)).toBe("237,125,49,1");
  });

  it("follows the first visible run's color when no buClr is given", async () => {
    const element = await renderTextBox(
      `<a:p><a:pPr><a:buChar char="•"/></a:pPr>
        <a:r><a:rPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:rPr><a:t>MARKER</a:t></a:r></a:p>`,
    );
    expect(normalizeColor(bulletOf(paragraphsOf(element)[0])!.style.color)).toBe("0,255,0,1");
  });

  it("falls back to black when nothing specifies a color", async () => {
    const element = await renderTextBox(
      `<a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:t>MARKER</a:t></a:r></a:p>`,
    );
    expect(normalizeColor(bulletOf(paragraphsOf(element)[0])!.style.color)).toBe("0,0,0,1");
  });
});

describe("hanging bullet gutter (marL + negative indent)", () => {
  it("positions the bullet absolutely inside the gutter for left-aligned text", async () => {
    const element = await renderTextBox(
      `<a:p><a:pPr marL="457200" indent="-457200"><a:buChar char="•"/></a:pPr><a:r><a:t>MARKER</a:t></a:r></a:p>`,
    );
    const paragraph = paragraphsOf(element)[0];
    const bullet = bulletOf(paragraph)!;
    // The hanging indent is converted into an absolutely positioned marker…
    expect(paragraph.style.textIndent).toBe("0px");
    expect(paragraph.style.position).toBe("relative");
    expect(bullet.style.position).toBe("absolute");
    expect(parseFloat(bullet.style.left)).toBeCloseTo(0, 3);
    // …whose width is the marL gutter (457200 EMU = 48px).
    expect(parseFloat(bullet.style.width)).toBeCloseTo(48, 3);
    // Text stays padded by marL.
    expect(parseFloat(paragraph.style.paddingLeft)).toBeCloseTo(48, 3);
  });

  it("uses an inline-block marker for centered paragraphs", async () => {
    const element = await renderTextBox(
      `<a:p><a:pPr marL="457200" indent="-457200" algn="ctr"><a:buChar char="•"/></a:pPr><a:r><a:t>MARKER</a:t></a:r></a:p>`,
    );
    const paragraph = paragraphsOf(element)[0];
    const bullet = bulletOf(paragraph)!;
    expect(paragraph.style.paddingLeft).toBe("0px");
    expect(bullet.style.display).toBe("inline-block");
    expect(parseFloat(bullet.style.width)).toBeCloseTo(48, 3);
  });
});

describe("bullets from the shape lstStyle", () => {
  it("picks the bullet for the paragraph's indent level", async () => {
    const lstStyle = `<a:lstStyle>
      <a:lvl1pPr><a:buChar char="•"/></a:lvl1pPr>
      <a:lvl2pPr><a:buChar char="-"/></a:lvl2pPr>
    </a:lstStyle>`;
    const element = await renderTextBox(
      `<a:p><a:r><a:t>level0</a:t></a:r></a:p>
       <a:p><a:pPr lvl="1"/><a:r><a:t>level1</a:t></a:r></a:p>`,
      { lstStyleXml: lstStyle },
    );
    const [first, second] = paragraphsOf(element);
    expect(first.textContent).toBe("• level0");
    expect(second.textContent).toBe("- level1");
  });
});
