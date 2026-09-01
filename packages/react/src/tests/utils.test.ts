import { describe, expect, it } from "vitest";

import { clamp } from "../utils";

describe("clamp", () => {
  it("passes a value inside the range through", () => {
    expect(clamp(5, 0, 10)).toBe(5);
    expect(clamp(0, 0, 10)).toBe(0);
    expect(clamp(10, 0, 10)).toBe(10);
  });

  it("pins a value outside the range to the nearest bound", () => {
    expect(clamp(-3, 0, 10)).toBe(0);
    expect(clamp(42, 0, 10)).toBe(10);
    expect(clamp(0.05, 0.1, 4)).toBe(0.1);
  });

  it("resolves crossed bounds to min, so an empty list clamps to 0", () => {
    expect(clamp(2, 0, -1)).toBe(0);
    expect(clamp(-5, 0, -1)).toBe(0);
  });
});
