import type { SlideHandle } from "@diceui/pptx-core";
import { describe, expect, it, vi } from "vitest";

import type { CachedThumbnail } from "../thumbnail-list";
import {
  releaseOldThumbnails,
  THUMBNAIL_CACHE_LIMIT,
  markThumbnailRecent,
} from "../thumbnail-list";

const disposeMocks = new WeakMap<CachedThumbnail, ReturnType<typeof vi.fn>>();

function entry(attached = false): CachedThumbnail {
  const element = document.createElement("div");
  if (attached) document.createElement("div").appendChild(element);
  const dispose = vi.fn();
  const cached = {
    slideHandle: { element, dispose } as unknown as SlideHandle,
    revision: 0,
  };
  disposeMocks.set(cached, dispose);
  return cached;
}

function fill(
  count: number,
  attached: (index: number) => boolean = () => false,
) {
  const cache = new Map<string, CachedThumbnail>();
  for (let i = 0; i < count; i++) cache.set(`s${i}`, entry(attached(i)));
  return cache;
}

describe("thumbnail cache", () => {
  it("releases the least recently shown miniatures past the limit", () => {
    const cache = fill(THUMBNAIL_CACHE_LIMIT + 2);
    const oldest = cache.get("s0")!;
    markThumbnailRecent(cache, "s1");

    releaseOldThumbnails(cache, `s${THUMBNAIL_CACHE_LIMIT + 1}`);

    expect(cache.size).toBe(THUMBNAIL_CACHE_LIMIT);
    expect(cache.has("s0")).toBe(false);
    expect(disposeMocks.get(oldest)).toHaveBeenCalled();
    // Shown again, s1 is now the most recent, so s2 went instead.
    expect(cache.has("s1")).toBe(true);
    expect(cache.has("s2")).toBe(false);
  });

  it("keeps miniatures on screen and the one just added", () => {
    const cache = fill(
      THUMBNAIL_CACHE_LIMIT + 1,
      (i) => i < THUMBNAIL_CACHE_LIMIT,
    );
    const newest = `s${THUMBNAIL_CACHE_LIMIT}`;

    releaseOldThumbnails(cache, newest);

    expect(cache.size).toBe(THUMBNAIL_CACHE_LIMIT + 1);
    expect(disposeMocks.get(cache.get(newest)!)).not.toHaveBeenCalled();
  });
});
