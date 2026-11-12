import { describe, expect, it } from "vitest";

import type { Rect } from "../selection/geometry";
import {
  getHandleCursor,
  getNodeBounds,
  getResizedRects,
  resizeRect,
  resizeRotatedRect,
} from "../selection/geometry";

// `u` and `v` are in the rect's own axes, 0 to 1 across it.
function toSlide(rect: Rect, rotation: number, u: number, v: number) {
  const radians = (rotation * Math.PI) / 180;
  const cx = rect.x + rect.w / 2;
  const cy = rect.y + rect.h / 2;
  const lx = (u - 0.5) * rect.w;
  const ly = (v - 0.5) * rect.h;
  return {
    x: cx + lx * Math.cos(radians) - ly * Math.sin(radians),
    y: cy + lx * Math.sin(radians) + ly * Math.cos(radians),
  };
}

const ORIGIN: Rect = { x: 100, y: 100, w: 200, h: 100 };

describe("resizeRotatedRect", () => {
  it("matches resizeRect when not rotated", () => {
    expect(resizeRotatedRect(ORIGIN, 0, "se", 30, 20)).toEqual(
      resizeRect(ORIGIN, "se", 30, 20),
    );
  });

  it("measures the drag along the shape's own axes", () => {
    // Turned 90°, the shape's width runs down the slide: dragging the east
    // handle down by 40 widens it by 40, and sideways does nothing.
    const down = resizeRotatedRect(ORIGIN, 90, "e", 0, 40);
    expect(down.w).toBeCloseTo(240);
    expect(down.h).toBeCloseTo(100);
    const sideways = resizeRotatedRect(ORIGIN, 90, "e", 40, 0);
    expect(sideways.w).toBeCloseTo(200);
  });

  it.each([
    ["e", 0, 0.5, 90, 0, 40],
    ["w", 1, 0.5, 30, -25, 10],
    ["se", 0, 0, 135, 20, -35],
    ["n", 0.5, 1, 210, 15, 15],
  ] as const)(
    "keeps the edge opposite the %s handle in place",
    (handle, u, v, rotation, dx, dy) => {
      const before = toSlide(ORIGIN, rotation, u, v);
      const after = toSlide(
        resizeRotatedRect(ORIGIN, rotation, handle, dx, dy),
        rotation,
        u,
        v,
      );
      expect(after.x).toBeCloseTo(before.x);
      expect(after.y).toBeCloseTo(before.y);
    },
  );

  it("keeps the aspect ratio with a locked corner drag", () => {
    const next = resizeRotatedRect(ORIGIN, 45, "se", 50, 50, true);
    expect(next.w / next.h).toBeCloseTo(2);
  });
});

describe("getHandleCursor", () => {
  it("turns the resize cursor with the shape", () => {
    expect(getHandleCursor("n", 0)).toBe("ns-resize");
    expect(getHandleCursor("n", 45)).toBe("nesw-resize");
    expect(getHandleCursor("n", 90)).toBe("ew-resize");
    expect(getHandleCursor("e", 90)).toBe("ns-resize");
    expect(getHandleCursor("nw", 180)).toBe("nwse-resize");
    expect(getHandleCursor("n", -90)).toBe("ew-resize");
  });
});

describe("getNodeBounds", () => {
  const node = (rotation: number) =>
    ({
      position: { x: 100, y: 100 },
      size: { w: 200, h: 100 },
      rotation,
    }) as never;

  it("is the node's box when not rotated", () => {
    expect(getNodeBounds(node(0))).toEqual(ORIGIN);
  });

  it("covers the rotated shape, about its center", () => {
    const bounds = getNodeBounds(node(90));
    expect(bounds.x).toBeCloseTo(150);
    expect(bounds.y).toBeCloseTo(50);
    expect(bounds.w).toBeCloseTo(100);
    expect(bounds.h).toBeCloseTo(200);
  });
});

describe("getResizedRects", () => {
  const node = (id: string, x: number, rotation: number) =>
    ({
      id,
      position: { x, y: 100 },
      size: { w: 200, h: 100 },
      rotation,
    }) as never;

  it("scales every shape by the dragged one's factors, along each shape's own axes", () => {
    const resized = node("a", 100, 0);
    const turned = node("b", 400, 90);
    const rects = getResizedRects(
      [resized, turned],
      resized,
      "e",
      100,
      0,
      false,
    );

    // The dragged shape grew 1.5× wide; the turned one grows 1.5× along its own width.
    expect(rects.get("a")).toEqual({ x: 100, y: 100, w: 300, h: 100 });
    const b = rects.get("b")!;
    expect(b.w).toBeCloseTo(300);
    expect(b.h).toBeCloseTo(100);

    // Its west edge (the top, turned 90°) stays put.
    const before = toSlide({ x: 400, y: 100, w: 200, h: 100 }, 90, 0, 0.5);
    const after = toSlide(b, 90, 0, 0.5);
    expect(after.x).toBeCloseTo(before.x);
    expect(after.y).toBeCloseTo(before.y);
  });
});
