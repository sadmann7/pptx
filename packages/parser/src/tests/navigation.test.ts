import { describe, expect, it } from "vitest";

import {
  resolveSlideJumpIndex,
  resolveSlideNavigationIndex,
  slideJumpTitle,
} from "../renderer/navigation";
import type { RenderContext } from "../renderer/render-context";

/** Minimal structural RenderContext: navigation only reads slide + presentation.slides. */
function ctxAt(currentIndex: number, slideCount = 4): RenderContext {
  const slides = Array.from({ length: slideCount }, (_, i) => ({
    id: `ppt/slides/slide${i + 1}.xml`,
    index: i,
  }));
  return {
    slide: slides[currentIndex],
    presentation: { slides },
  } as unknown as RenderContext;
}

describe("resolveSlideJumpIndex", () => {
  it("maps a slide relationship target to its presentation index", () => {
    const rel = { type: "slide", target: "slide3.xml" };
    expect(resolveSlideJumpIndex(ctxAt(0), rel)).toBe(2);
  });

  it("resolves relative targets against the current slide's folder", () => {
    const rel = { type: "slide", target: "../slides/slide2.xml" };
    expect(resolveSlideJumpIndex(ctxAt(3), rel)).toBe(1);
  });

  it("returns undefined for external targets", () => {
    const rel = { type: "hyperlink", target: "https://example.com", targetMode: "External" };
    expect(resolveSlideJumpIndex(ctxAt(0), rel)).toBeUndefined();
  });

  it("falls back to slideN.xml numbering for unknown paths", () => {
    const rel = { type: "slide", target: "slide9.xml" };
    expect(resolveSlideJumpIndex(ctxAt(0), rel)).toBe(8);
  });
});

describe("resolveSlideNavigationIndex", () => {
  it("handles hlinksldjump with a relationship", () => {
    const rel = { type: "slide", target: "slide4.xml" };
    expect(resolveSlideNavigationIndex(ctxAt(0), "ppaction://hlinksldjump", rel)).toBe(3);
  });

  it.each([
    ["ppaction://hlinkshowjump?jump=firstslide", 1, 0],
    ["ppaction://hlinkshowjump?jump=lastslide", 1, 3],
    ["ppaction://hlinkshowjump?jump=nextslide", 1, 2],
    ["ppaction://hlinkshowjump?jump=previousslide", 1, 0],
  ] as const)("handles %s", (action, from, expected) => {
    expect(resolveSlideNavigationIndex(ctxAt(from), action)).toBe(expected);
  });

  it("returns undefined at deck boundaries", () => {
    expect(
      resolveSlideNavigationIndex(ctxAt(3), "ppaction://hlinkshowjump?jump=nextslide"),
    ).toBeUndefined();
    expect(
      resolveSlideNavigationIndex(ctxAt(0), "ppaction://hlinkshowjump?jump=previousslide"),
    ).toBeUndefined();
  });

  it("returns undefined for unknown actions or missing input", () => {
    expect(resolveSlideNavigationIndex(ctxAt(0), "ppaction://customshow")).toBeUndefined();
    expect(resolveSlideNavigationIndex(ctxAt(0), undefined)).toBeUndefined();
  });
});

describe("slideJumpTitle", () => {
  it("formats a 1-based label", () => {
    expect(slideJumpTitle(0)).toBe("Go to slide 1");
    expect(slideJumpTitle(11)).toBe("Go to slide 12");
  });
});
