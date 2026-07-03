import { describe, expect, it } from "vitest";

import {
  angleToDeg,
  detectUnit,
  emuToPt,
  emuToPx,
  hundredthPtToPt,
  pctToDecimal,
  ptToPx,
  smartToPx,
} from "./units";

describe("unit conversions", () => {
  it("converts EMU to pixels at 96 DPI", () => {
    expect(emuToPx(914400)).toBe(96); // 1 inch
    expect(emuToPx(0)).toBe(0);
    expect(emuToPx(457200)).toBe(48);
  });

  it("converts EMU to points", () => {
    expect(emuToPt(12700)).toBe(1);
    expect(emuToPt(914400)).toBe(72); // 1 inch = 72pt
  });

  it("converts points to pixels", () => {
    expect(ptToPx(72)).toBe(96);
    expect(ptToPx(12)).toBe(16);
  });

  it("converts OOXML angles (60000ths of a degree) to degrees", () => {
    expect(angleToDeg(5400000)).toBe(90);
    expect(angleToDeg(21600000)).toBe(360);
    expect(angleToDeg(-1800000)).toBe(-30);
  });

  it("converts OOXML percentages (100000ths) to decimals", () => {
    expect(pctToDecimal(100000)).toBe(1);
    expect(pctToDecimal(50000)).toBe(0.5);
  });

  it("converts hundredths of a point to points", () => {
    expect(hundredthPtToPt(1800)).toBe(18);
  });
});

describe("smart unit detection", () => {
  it("treats large magnitudes as EMU", () => {
    expect(detectUnit(914400)).toBe("emu");
    expect(detectUnit(-914400)).toBe("emu");
    expect(smartToPx(914400)).toBe(96);
  });

  it("treats small magnitudes as points", () => {
    expect(detectUnit(72)).toBe("point");
    expect(smartToPx(72)).toBe(96);
  });
});
