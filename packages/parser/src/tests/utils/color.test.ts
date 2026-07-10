import { describe, expect, it } from "vitest";

import {
  applyAlpha,
  applyColorModifiers,
  applyLumMod,
  applyLumOff,
  applySatMod,
  applyShade,
  applyTint,
  hexToRgb,
  hslToRgb,
  presetColorToHex,
  rgbToHex,
  rgbToHsl,
} from "../../utils/color";

describe("hex/rgb conversions", () => {
  it("parses 6-digit hex with and without #", () => {
    expect(hexToRgb("#FF8000")).toEqual({ r: 255, g: 128, b: 0 });
    expect(hexToRgb("ff8000")).toEqual({ r: 255, g: 128, b: 0 });
  });

  it("expands 3-digit hex", () => {
    expect(hexToRgb("#F80")).toEqual({ r: 255, g: 136, b: 0 });
  });

  it("returns black for invalid hex", () => {
    expect(hexToRgb("nope")).toEqual({ r: 0, g: 0, b: 0 });
    expect(hexToRgb("#12345")).toEqual({ r: 0, g: 0, b: 0 });
  });

  it("formats and clamps rgb to hex", () => {
    expect(rgbToHex(255, 128, 0)).toBe("#ff8000");
    expect(rgbToHex(300, -5, 12.4)).toBe("#ff000c");
  });
});

describe("hsl conversions", () => {
  it("round-trips primary colors", () => {
    for (const [r, g, b] of [
      [255, 0, 0],
      [0, 255, 0],
      [0, 0, 255],
      [128, 128, 128],
    ] as const) {
      const { h, s, l } = rgbToHsl(r, g, b);
      expect(hslToRgb(h, s, l)).toEqual({ r, g, b });
    }
  });

  it("normalizes hue outside 0-360", () => {
    expect(hslToRgb(360 + 120, 1, 0.5)).toEqual(hslToRgb(120, 1, 0.5));
    expect(hslToRgb(-240, 1, 0.5)).toEqual(hslToRgb(120, 1, 0.5));
  });

  it("treats zero saturation as grayscale", () => {
    expect(hslToRgb(200, 0, 0.5)).toEqual({ r: 128, g: 128, b: 128 });
  });
});

describe("OOXML modifiers", () => {
  it("tint 100000 keeps the color, tint 0 gives white", () => {
    expect(applyTint("#336699", 100000)).toBe("#336699");
    expect(applyTint("#336699", 0)).toBe("#ffffff");
  });

  it("shade 100000 keeps the color, shade 0 gives black", () => {
    expect(applyShade("#336699", 100000)).toBe("#336699");
    expect(applyShade("#336699", 0)).toBe("#000000");
  });

  it("lumMod scales luminance down", () => {
    // A common Office pattern: accent color at lumMod 50%.
    const darker = applyLumMod("#4472C4", 50000);
    const { l: lBefore } = rgbToHsl(
      ...(Object.values(hexToRgb("#4472C4")) as [number, number, number]),
    );
    const { l: lAfter } = rgbToHsl(
      ...(Object.values(hexToRgb(darker)) as [number, number, number]),
    );
    expect(lAfter).toBeCloseTo(lBefore * 0.5, 1);
  });

  it("lumOff shifts luminance up", () => {
    const lighter = applyLumOff("#4472C4", 20000);
    const before = rgbToHsl(68, 114, 196).l;
    const after = rgbToHsl(...(Object.values(hexToRgb(lighter)) as [number, number, number])).l;
    expect(after).toBeCloseTo(before + 0.2, 1);
  });

  it("satMod 0 desaturates fully", () => {
    const gray = applySatMod("#4472C4", 0);
    const { r, g, b } = hexToRgb(gray);
    expect(r).toBe(g);
    expect(g).toBe(b);
  });

  it("alpha converts OOXML 0-100000 to 0-1 with clamping", () => {
    expect(applyAlpha(100000)).toBe(1);
    expect(applyAlpha(50000)).toBe(0.5);
    expect(applyAlpha(-1)).toBe(0);
    expect(applyAlpha(200000)).toBe(1);
  });
});

describe("applyColorModifiers", () => {
  it("applies modifiers in document order", () => {
    // lumMod then lumOff: the standard Office "lighter variant" recipe.
    const viaComposite = applyColorModifiers("#4472C4", [
      { name: "lumMod", val: 60000 },
      { name: "lumOff", val: 40000 },
    ]).color;
    const manual = applyLumOff(applyLumMod("#4472C4", 60000), 40000);
    expect(viaComposite).toBe(manual);
  });

  it("strips a: prefixes from modifier names", () => {
    const withPrefix = applyColorModifiers("#336699", [{ name: "a:shade", val: 50000 }]);
    const without = applyColorModifiers("#336699", [{ name: "shade", val: 50000 }]);
    expect(withPrefix.color).toBe(without.color);
  });

  it("tracks alpha alongside color and composes alphaMod", () => {
    const { color, alpha } = applyColorModifiers("#FF0000", [
      { name: "alpha", val: 50000 },
      { name: "alphaMod", val: 50000 },
    ]);
    expect(color).toBe("#FF0000");
    expect(alpha).toBe(0.25);
  });

  it("ignores unknown modifiers", () => {
    const { color, alpha } = applyColorModifiers("#FF0000", [{ name: "frobnicate", val: 1 }]);
    expect(color).toBe("#FF0000");
    expect(alpha).toBe(1);
  });

  it("applies inv, gray, and comp", () => {
    expect(applyColorModifiers("#000000", [{ name: "inv", val: 0 }]).color).toBe("#ffffff");
    const gray = applyColorModifiers("#FF0000", [{ name: "gray", val: 0 }]).color;
    const { r, g, b } = hexToRgb(gray);
    expect(r).toBe(g);
    expect(g).toBe(b);
    // Complement of pure red is pure cyan.
    expect(applyColorModifiers("#FF0000", [{ name: "comp", val: 0 }]).color).toBe("#00ffff");
  });
});

describe("presetColorToHex", () => {
  it("resolves exact OOXML preset names", () => {
    expect(presetColorToHex("red")).toBe("#FF0000");
    expect(presetColorToHex("cornflowerBlue")).toBe("#6495ED");
  });

  it("falls back to case-insensitive lookup", () => {
    expect(presetColorToHex("CornflowerBlue")).toBe("#6495ED");
  });

  it("returns undefined for unknown names", () => {
    expect(presetColorToHex("notAColor")).toBeUndefined();
  });
});
