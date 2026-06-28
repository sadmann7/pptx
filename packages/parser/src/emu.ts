/**
 * EMU (English Metric Units) conversion utilities.
 *
 * OOXML stores all measurements in EMU.
 *   1 inch  = 914400 EMU
 *   1 point = 12700 EMU
 *   1 cm    = 360000 EMU
 *
 * We normalize everything to points (pt) in the AST because:
 *   - Points are the native unit for typography
 *   - At 96 DPI, 1pt ≈ 1.33px, which is easy to work with in CSS
 *   - Avoids floating-point churn from double-conversion
 */

const EMU_PER_PT = 12700

export function emuToPoints(emu: number | string | undefined): number {
  if (emu === undefined || emu === '') return 0
  const n = typeof emu === 'string' ? Number.parseInt(emu, 10) : emu
  return n / EMU_PER_PT
}

/** Hundredths of a point → points (used in some font size attrs) */
export function hunPtToPoints(hunPt: number | string | undefined): number {
  if (hunPt === undefined || hunPt === '') return 0
  const n = typeof hunPt === 'string' ? Number.parseInt(hunPt, 10) : hunPt
  return n / 100
}

/**
 * Percent expressed as 1000x (e.g. 100000 = 100%).
 * Used for line spacing, indent, etc.
 */
export function perMilleToPercent(val: number | string | undefined): number {
  if (val === undefined || val === '') return 100
  const n = typeof val === 'string' ? Number.parseInt(val, 10) : val
  return n / 1000
}

/**
 * Angle in 60000ths of a degree → degrees.
 * Used for rotation, gradient angles.
 */
export function angleToDegs(val: number | string | undefined): number {
  if (val === undefined || val === '') return 0
  const n = typeof val === 'string' ? Number.parseInt(val, 10) : val
  return n / 60000
}
