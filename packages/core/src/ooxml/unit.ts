// 914400 EMU per inch at 96 px per inch.
export function emuToPx(emu: number): number {
  return (emu / 914400) * 96;
}

// OOXML requires integer EMU.
export function pxToEmu(px: number): number {
  return Math.round((px / 96) * 914400);
}

// OOXML angles are integer 60000ths of a degree.
export function degToAngle(deg: number): number {
  return Math.round(deg * 60000);
}

export function emuToPt(emu: number): number {
  return emu / 12700;
}

export function angleToDeg(angle: number): number {
  return angle / 60000;
}

export function pctToDecimal(pct: number): number {
  return pct / 100000;
}

export function hundredthPtToPt(val: number): number {
  return val / 100;
}

export function ptToPx(pt: number): number {
  return (pt * 96) / 72;
}

// Values above 20000 are almost certainly EMU, since one point is 12700 EMU.
export function detectUnit(value: number): "emu" | "point" {
  return Math.abs(value) > 20000 ? "emu" : "point";
}

export function smartToPx(value: number): number {
  if (detectUnit(value) === "emu") {
    return emuToPx(value);
  }
  return ptToPx(value);
}
