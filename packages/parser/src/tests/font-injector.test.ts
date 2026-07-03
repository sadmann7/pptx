import { describe, expect, it } from "vitest";

import { collectPriorityTypefaces, injectEmbeddedFonts } from "../fonts/font-injector";
import type { PresentationData } from "../model/presentation";

function presentationWithFonts(
  typefaces: string[],
  themeMajor = "",
  themeMinor = "",
): PresentationData {
  return {
    embeddedFonts: typefaces.map((typeface) => ({
      typeface,
      regular: { path: `ppt/fonts/${typeface}.fntdata` },
    })),
    fonts: new Map(),
    themes: new Map([
      [
        "ppt/theme/theme1.xml",
        {
          majorFont: { latin: themeMajor, ea: "", cs: "" },
          minorFont: { latin: themeMinor, ea: "", cs: "" },
        },
      ],
    ]),
  } as unknown as PresentationData;
}

describe("collectPriorityTypefaces", () => {
  const slideXml = `<p:sld><a:r><a:rPr><a:latin typeface="Inter"/></a:rPr><a:t>x</a:t></a:r></p:sld>`;

  it("returns only typefaces referenced by the given sources", () => {
    const pres = presentationWithFonts(["Inter", "Archivo", "Space Grotesk"]);
    const priority = collectPriorityTypefaces(pres, [slideXml]);
    expect(priority).toEqual(new Set(["Inter"]));
  });

  it("matches typefaces across multiple sources (slide + layout/master)", () => {
    const pres = presentationWithFonts(["Inter", "Archivo", "Space Grotesk"]);
    const layoutXml = `<p:sldLayout><a:defRPr><a:latin typeface="Archivo"/></a:defRPr></p:sldLayout>`;
    const priority = collectPriorityTypefaces(pres, [slideXml, layoutXml, undefined]);
    expect(priority).toEqual(new Set(["Inter", "Archivo"]));
  });

  it("always includes theme major/minor fonts", () => {
    const pres = presentationWithFonts(
      ["Inter", "Archivo", "Space Grotesk"],
      "Space Grotesk",
      "Archivo",
    );
    const priority = collectPriorityTypefaces(pres, [slideXml]);
    expect(priority).toEqual(new Set(["Inter", "Archivo", "Space Grotesk"]));
  });

  it("returns undefined without embedded fonts or sources", () => {
    expect(collectPriorityTypefaces(presentationWithFonts([]), [slideXml])).toBeUndefined();
    expect(collectPriorityTypefaces(presentationWithFonts(["Inter"]), [undefined])).toBeUndefined();
  });
});

describe("injectEmbeddedFonts handle", () => {
  it("returns a resolved noop handle when there are no embedded fonts", async () => {
    const handle = injectEmbeddedFonts(presentationWithFonts([]));
    await expect(handle.ready).resolves.toBeUndefined();
    await expect(handle.complete).resolves.toBeUndefined();
    expect(() => handle.dispose()).not.toThrow();
  });
});
