import { describe, expect, it } from "vitest";

import { getUnionRect, snapMove, snapResize } from "../snap";

const SLIDE = { x: 0, y: 0, w: 960, h: 540 };
const BOX = { x: 100, y: 100, w: 100, h: 50 };

describe("snapMove", () => {
  it("snaps a center onto the slide's center and draws the guide across the slide", () => {
    // Center x lands at 150 + 327 = 477, 3px short of the slide's 480.
    const result = snapMove(BOX, 327, 0, [SLIDE], 6);
    expect(result.dx).toBe(330);
    expect(result.dy).toBe(0);
    expect(result.guides).toEqual([{ axis: "x", position: 480, start: 0, end: 540 }]);
  });

  it("snaps edges onto another shape's edges, both axes at once", () => {
    const other = { x: 400, y: 300, w: 80, h: 80 };
    // Left edge at 398 (→ 400), bottom edge at 296 (→ 300).
    const result = snapMove(BOX, 298, 146, [other], 6);
    expect(result.dx).toBe(300);
    expect(result.dy).toBe(150);
    expect(result.guides).toEqual([
      // Vertical guide at x=400 spanning the moved box (250-300) and the other (300-380).
      { axis: "x", position: 400, start: 250, end: 380 },
      { axis: "y", position: 300, start: 400, end: 500 },
    ]);
  });

  it("picks the closest anchor when several are in range", () => {
    const near = { x: 302, y: 0, w: 10, h: 10 };
    const nearer = { x: 299, y: 0, w: 10, h: 10 };
    // Left edge at 300: 2px from `near`, 1px from `nearer`.
    expect(snapMove(BOX, 200, 0, [near, nearer], 6).dx).toBe(199);
  });

  it("leaves the drag alone outside the threshold", () => {
    const result = snapMove(BOX, 320, 0, [SLIDE], 6);
    expect(result).toEqual({ dx: 320, dy: 0, guides: [] });
  });

  it("does not snap an axis that is locked", () => {
    // y would snap the box's center (125 + 143 = 268) onto 270.
    const result = snapMove(BOX, 0, 143, [SLIDE], 6, { x: true, y: false });
    expect(result.dy).toBe(143);
  });
});

describe("snapResize", () => {
  it("snaps only the edge the grip moves", () => {
    // Dragging the east edge from 200 to 477 snaps it to the slide center.
    const result = snapResize(BOX, "e", 277, 0, [SLIDE], 6);
    expect(result.dx).toBe(280);
    expect(result.guides).toEqual([{ axis: "x", position: 480, start: 0, end: 540 }]);
  });

  it("snaps a west edge and the guide follows the resized rect", () => {
    const other = { x: 40, y: 400, w: 20, h: 20 };
    const result = snapResize(BOX, "nw", -57, 0, [other], 6);
    expect(result.dx).toBe(-60);
    expect(result.guides[0]).toEqual({ axis: "x", position: 40, start: 100, end: 420 });
  });

  it("skips aspect-locked corner drags", () => {
    expect(snapResize(BOX, "se", 277, 3, [SLIDE], 6, true)).toEqual({
      dx: 277,
      dy: 3,
      guides: [],
    });
  });
});

describe("getUnionRect", () => {
  it("bounds a multi-selection", () => {
    expect(
      getUnionRect([
        { x: 10, y: 20, w: 30, h: 40 },
        { x: 50, y: 5, w: 10, h: 10 },
      ]),
    ).toEqual({ x: 10, y: 5, w: 50, h: 55 });
    expect(getUnionRect([])).toBeNull();
  });
});
