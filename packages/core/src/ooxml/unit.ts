/** EMU to pixels (at 96 DPI). */
export function emuToPx(emu: number): number {
  return (emu / 914400) * 96;
}

/** Pixels (at 96 DPI) to EMU, rounded to an integer as required by OOXML. */
export function pxToEmu(px: number): number {
  return Math.round((px / 96) * 914400);
}

/** Degrees to OOXML angle (60000ths of a degree), rounded to an integer. */
export function degToAngle(deg: number): number {
  return Math.round(deg * 60000);
}

/** EMU to points. */
export function emuToPt(emu: number): number {
  return emu / 12700;
}

/** OOXML angle (60000ths of a degree) to degrees. */
export function angleToDeg(angle: number): number {
  return angle / 60000;
}

/** OOXML percentage (100000ths) to a decimal fraction (0..1 range for 0%..100%). */
export function pctToDecimal(pct: number): number {
  return pct / 100000;
}

/** Hundredths of a point to points (used for font sizes in OOXML). */
export function hundredthPtToPt(val: number): number {
  return val / 100;
}

/** Points to pixels (at 96 DPI). */
export function ptToPx(pt: number): number {
  return (pt * 96) / 72;
}

/** Values above 20000 are almost certainly EMU (one point is 12700 EMU). */
export function detectUnit(value: number): "emu" | "point" {
  return Math.abs(value) > 20000 ? "emu" : "point";
}

/** Converts EMU or points to pixels, guessing the unit with {@link detectUnit}. */
export function smartToPx(value: number): number {
  if (detectUnit(value) === "emu") {
    return emuToPx(value);
  }
  return ptToPx(value);
}
