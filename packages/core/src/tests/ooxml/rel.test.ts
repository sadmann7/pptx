import { describe, expect, it } from "vitest";

import { isExternalTargetMode, parseRels, resolveRelTarget } from "../../ooxml/rel";

const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
  <Relationship Type="missing-id" Target="x.xml"/>
</Relationships>`;

describe("parseRels", () => {
  it("maps relationship ids to entries", () => {
    const rels = parseRels(RELS_XML);
    expect(rels.size).toBe(2);
    expect(rels.get("rId1")).toEqual({
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
      target: "../slideLayouts/slideLayout1.xml",
      targetMode: undefined,
    });
    expect(rels.get("rId2")?.targetMode).toBe("External");
  });

  it("skips relationships without an Id", () => {
    const rels = parseRels(RELS_XML);
    expect([...rels.keys()]).toEqual(["rId1", "rId2"]);
  });

  it("returns an empty map for empty or malformed input", () => {
    expect(parseRels("").size).toBe(0);
    expect(parseRels("<not-xml <<").size).toBe(0);
  });
});

describe("isExternalTargetMode", () => {
  it("matches 'External' case-insensitively and trims whitespace", () => {
    expect(isExternalTargetMode("External")).toBe(true);
    expect(isExternalTargetMode(" external ")).toBe(true);
    expect(isExternalTargetMode("Internal")).toBe(false);
    expect(isExternalTargetMode(undefined)).toBe(false);
  });
});

describe("resolveRelTarget", () => {
  it("resolves parent-relative targets", () => {
    expect(resolveRelTarget("ppt/slides", "../slideLayouts/slideLayout1.xml")).toBe(
      "ppt/slideLayouts/slideLayout1.xml",
    );
  });

  it("resolves child-relative targets", () => {
    expect(resolveRelTarget("ppt/slides", "media/image1.png")).toBe("ppt/slides/media/image1.png");
    expect(resolveRelTarget("ppt", "slides/slide1.xml")).toBe("ppt/slides/slide1.xml");
  });

  it("handles absolute targets by stripping the leading slash", () => {
    expect(resolveRelTarget("ppt/slides", "/ppt/media/image1.png")).toBe("ppt/media/image1.png");
  });

  it("ignores '.' segments and strips query/fragment suffixes", () => {
    expect(resolveRelTarget("ppt", "./slides/slide1.xml")).toBe("ppt/slides/slide1.xml");
    expect(resolveRelTarget("ppt", "slides/slide1.xml?rev=2#frag")).toBe("ppt/slides/slide1.xml");
  });

  it("decodes URI-encoded path segments", () => {
    expect(resolveRelTarget("ppt/slides", "media/my%20image.png")).toBe(
      "ppt/slides/media/my image.png",
    );
  });

  it("normalizes backslashes", () => {
    expect(resolveRelTarget("ppt\\slides", "media\\image1.png")).toBe(
      "ppt/slides/media/image1.png",
    );
  });
});
