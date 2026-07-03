import { describe, expect, it } from "vitest";

import type { MediaResolver } from "../utils/media";
import {
  findMediaByTarget,
  findMediaByTargetAsync,
  getMimeType,
  getOrCreateBlobUrl,
  resolveMediaPath,
  resolveMediaPathCandidates,
} from "../utils/media";

describe("getMimeType", () => {
  it("maps known extensions", () => {
    expect(getMimeType("ppt/media/image1.png")).toBe("image/png");
    expect(getMimeType("photo.jpg")).toBe("image/jpeg");
    expect(getMimeType("photo.jpeg")).toBe("image/jpeg");
    expect(getMimeType("anim.gif")).toBe("image/gif");
    expect(getMimeType("vector.svg")).toBe("image/svg+xml");
    expect(getMimeType("old.emf")).toBe("image/x-emf");
    expect(getMimeType("clip.mp4")).toBe("video/mp4");
    expect(getMimeType("audio.mp3")).toBe("audio/mpeg");
  });

  it("is case-insensitive on the extension", () => {
    expect(getMimeType("IMAGE1.PNG")).toBe("image/png");
  });

  it("falls back to octet-stream for unknown or missing extensions", () => {
    expect(getMimeType("file.xyz")).toBe("application/octet-stream");
    expect(getMimeType("noextension")).toBe("application/octet-stream");
    expect(getMimeType("")).toBe("application/octet-stream");
  });
});

describe("resolveMediaPath", () => {
  it("resolves relative rels targets to canonical ppt/media paths", () => {
    expect(resolveMediaPath("../media/image1.png")).toBe("ppt/media/image1.png");
    expect(resolveMediaPath("media/image1.png")).toBe("ppt/media/image1.png");
    expect(resolveMediaPath("/ppt/media/image1.png")).toBe("ppt/media/image1.png");
  });

  it("normalizes backslashes and dot segments", () => {
    expect(resolveMediaPath("..\\media\\image1.png")).toBe("ppt/media/image1.png");
    expect(resolveMediaPath("./media/./image1.png")).toBe("ppt/media/image1.png");
  });

  it("strips query strings and fragments", () => {
    expect(resolveMediaPath("../media/image1.png?rev=2")).toBe("ppt/media/image1.png");
    expect(resolveMediaPath("../media/image1.png#frag")).toBe("ppt/media/image1.png");
  });

  it("decodes percent-encoded segments", () => {
    expect(resolveMediaPath("../media/my%20image.png")).toBe("ppt/media/my image.png");
  });

  it("falls back to the basename when no media directory is present", () => {
    expect(resolveMediaPath("image1.png")).toBe("ppt/media/image1.png");
    expect(resolveMediaPath("some/dir/image1.png")).toBe("ppt/media/image1.png");
  });
});

describe("resolveMediaPathCandidates", () => {
  it("returns a single candidate for plain targets", () => {
    expect(resolveMediaPathCandidates("../media/image1.png")).toEqual(["ppt/media/image1.png"]);
  });

  it("prefers the decoded form but keeps the raw name as fallback", () => {
    expect(resolveMediaPathCandidates("../media/my%20image.png")).toEqual([
      "ppt/media/my image.png",
      "ppt/media/my%20image.png",
    ]);
  });
});

describe("findMediaByTarget", () => {
  const data = new Uint8Array([1, 2, 3]);

  it("finds media by decoded path", () => {
    const media = new Map([["ppt/media/image1.png", data]]);
    expect(findMediaByTarget("../media/image1.png", media)).toEqual({
      mediaPath: "ppt/media/image1.png",
      data,
    });
  });

  it("falls back to the raw percent-encoded zip entry name", () => {
    const media = new Map([["ppt/media/my%20image.png", data]]);
    expect(findMediaByTarget("../media/my%20image.png", media)).toEqual({
      mediaPath: "ppt/media/my%20image.png",
      data,
    });
  });

  it("returns undefined when nothing matches", () => {
    expect(findMediaByTarget("../media/missing.png", new Map())).toBeUndefined();
  });
});

describe("findMediaByTargetAsync", () => {
  const data = new Uint8Array([9, 9]);

  it("prefers the eager media map over the resolver", async () => {
    const media = new Map([["ppt/media/image1.png", data]]);
    const resolver: MediaResolver = {
      resolve: async () => {
        throw new Error("resolver must not be called for eager hits");
      },
    };
    await expect(findMediaByTargetAsync("../media/image1.png", media, resolver)).resolves.toEqual({
      mediaPath: "ppt/media/image1.png",
      data,
    });
  });

  it("delegates to the resolver on a miss", async () => {
    const resolver: MediaResolver = {
      resolve: async (target) => ({ mediaPath: resolveMediaPath(target), data }),
    };
    await expect(findMediaByTargetAsync("../media/lazy.png", new Map(), resolver)).resolves.toEqual(
      { mediaPath: "ppt/media/lazy.png", data },
    );
  });

  it("resolves undefined with no resolver", async () => {
    await expect(findMediaByTargetAsync("../media/lazy.png", new Map())).resolves.toBeUndefined();
  });
});

describe("getOrCreateBlobUrl", () => {
  it("creates a blob URL in happy-dom and caches it per media path", () => {
    // happy-dom implements URL.createObjectURL; this asserts we can rely on it.
    expect(typeof URL.createObjectURL).toBe("function");

    const cache = new Map<string, string>();
    const url = getOrCreateBlobUrl("ppt/media/image1.png", new Uint8Array([1, 2, 3]), cache);
    expect(typeof url).toBe("string");
    expect(url.length).toBeGreaterThan(0);
    expect(cache.get("ppt/media/image1.png")).toBe(url);

    // Second call with the same path hits the cache even with different data.
    const again = getOrCreateBlobUrl("ppt/media/image1.png", new Uint8Array([4, 5]), cache);
    expect(again).toBe(url);

    // Different path → different URL.
    const other = getOrCreateBlobUrl("ppt/media/image2.png", new Uint8Array([6]), cache);
    expect(other).not.toBe(url);
    expect(cache.size).toBe(2);
  });
});
