import { beforeAll, describe, expect, it } from "vitest";

import type { PresentationData } from "../model/presentation";
import { buildPresentation } from "../model/presentation";
import { parseZip } from "../parser/zip-parser";
import { buildTextIndex, searchPresentation, searchText } from "../search/text-search";
import { buildPptxWithShapes } from "./minimal-pptx";

function textShape(id: number, text: string): string {
  return `<p:sp>
<p:nvSpPr><p:cNvPr id="${id}" name="Text ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="914400" y="${id * 914400}"/><a:ext cx="4572000" cy="914400"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
</p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody>
</p:sp>`;
}

let presentation: PresentationData;

beforeAll(async () => {
  const buffer = await buildPptxWithShapes(
    textShape(2, "Revenue grew 40 percent") + textShape(3, "revenue outlook for Q3"),
  );
  presentation = buildPresentation(await parseZip(buffer));
});

describe("buildTextIndex", () => {
  it("indexes shape text with slide position and bounds", () => {
    const index = buildTextIndex(presentation);
    expect(index.length).toBe(2);
    expect(index[0].slideIndex).toBe(0);
    expect(index[0].textKind).toBe("shape");
    expect(index[0].text).toContain("Revenue");
    expect(index[0].bounds.w).toBeGreaterThan(0);
  });

  it("can exclude shapes", () => {
    expect(buildTextIndex(presentation, { includeShapes: false })).toHaveLength(0);
  });
});

describe("searchText / searchPresentation", () => {
  it("matches case-insensitively by default", () => {
    const results = searchPresentation(presentation, "revenue");
    expect(results).toHaveLength(2);
  });

  it("respects matchCase", () => {
    const results = searchPresentation(presentation, "revenue", { matchCase: true });
    expect(results).toHaveLength(1);
    expect(results[0].text).toContain("outlook");
  });

  it("respects wholeWord", () => {
    expect(searchPresentation(presentation, "grew", { wholeWord: true })).toHaveLength(1);
    expect(searchPresentation(presentation, "gre", { wholeWord: true })).toHaveLength(0);
  });

  it("supports regex queries with match offsets and snippets", () => {
    const results = searchPresentation(presentation, String.raw`\d+ percent`, { useRegex: true });
    expect(results).toHaveLength(1);
    const r = results[0];
    expect(r.text.slice(r.matchStart, r.matchEnd)).toBe("40 percent");
    expect(r.snippet).toContain("40 percent");
  });

  it("returns no results for an empty query", () => {
    const index = buildTextIndex(presentation);
    expect(searchText(index, "")).toHaveLength(0);
  });
});
