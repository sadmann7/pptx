/** EMU to pixels (at 96 DPI). */
export function emuToPx(emu: number): number {
  return (emu / 914400) * 96;
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
