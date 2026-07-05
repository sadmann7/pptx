/**
 * Evaluator for DrawingML geometry guide formulas (ECMA-376 §20.1.9.11).
 *
 * Guides appear in `a:avLst` and `a:gdLst` of both preset and custom
 * geometries. Each guide is `<a:gd name="..." fmla="op arg1 arg2 arg3"/>`
 * where arguments are numbers or references to previously defined guides
 * (including the built-in variables like `w`, `h`, `ss`, `hc`).
 *
 * Angles are expressed in 60000ths of a degree throughout.
 */

export interface GuideDefinition {
  name: string;
  fmla: string;
}

/** 60000ths-of-a-degree angle to radians. */
function guideAngleToRad(ang: number): number {
  return ((ang / 60000) * Math.PI) / 180;
}

/** Radians to 60000ths of a degree. */
function radToGuideAngle(rad: number): number {
  return ((rad * 180) / Math.PI) * 60000;
}

/**
 * Built-in guide variables (ECMA-376 §20.1.10.56 ST_GeomGuideFormula).
 * `w`/`h` are the shape extents in the geometry coordinate space.
 */
export function createGuideContext(w: number, h: number): Map<string, number> {
  const ss = Math.min(w, h);
  return new Map<string, number>([
    ["w", w],
    ["h", h],
    ["ss", ss],
    ["hc", w / 2],
    ["vc", h / 2],
    ["t", 0],
    ["l", 0],
    ["r", w],
    ["b", h],
    ["wd2", w / 2],
    ["wd3", w / 3],
    ["wd4", w / 4],
    ["wd5", w / 5],
    ["wd6", w / 6],
    ["wd8", w / 8],
    ["wd10", w / 10],
    ["wd12", w / 12],
    ["wd32", w / 32],
    ["hd2", h / 2],
    ["hd3", h / 3],
    ["hd4", h / 4],
    ["hd5", h / 5],
    ["hd6", h / 6],
    ["hd8", h / 8],
    ["hd10", h / 10],
    ["hd12", h / 12],
    ["ssd2", ss / 2],
    ["ssd4", ss / 4],
    ["ssd6", ss / 6],
    ["ssd8", ss / 8],
    ["ssd16", ss / 16],
    ["ssd32", ss / 32],
    // Angle constants in 60000ths of a degree: cdN = 360°/N
    ["cd2", 10800000],
    ["cd4", 5400000],
    ["cd8", 2700000],
    ["3cd4", 16200000],
    ["3cd8", 8100000],
    ["5cd8", 13500000],
    ["7cd8", 18900000],
  ]);
}

/** Resolve a formula argument: a literal number or a reference to a guide. */
export function resolveGuideOperand(
  token: string | undefined,
  guides: Map<string, number>,
): number {
  if (token === undefined) return 0;
  const n = Number(token);
  if (!Number.isNaN(n)) return n;
  return guides.get(token) ?? 0;
}

/**
 * Evaluate a single guide formula string against already-known guides.
 * Unknown operators evaluate to 0.
 */
export function evaluateGuideFormula(fmla: string, guides: Map<string, number>): number {
  const parts = fmla.trim().split(/\s+/);
  const op = parts[0];
  const x = resolveGuideOperand(parts[1], guides);
  const y = resolveGuideOperand(parts[2], guides);
  const z = resolveGuideOperand(parts[3], guides);

  switch (op) {
    case "val":
      return x;
    case "*/":
      return z === 0 ? 0 : (x * y) / z;
    case "+-":
      return x + y - z;
    case "+/":
      return z === 0 ? 0 : (x + y) / z;
    case "?:":
      return x > 0 ? y : z;
    case "abs":
      return Math.abs(x);
    case "min":
      return Math.min(x, y);
    case "max":
      return Math.max(x, y);
    case "sqrt":
      return Math.sqrt(Math.max(0, x));
    case "mod":
      return Math.sqrt(x * x + y * y + z * z);
    case "pin":
      // pin x y z: clamp y into [x, z]
      return y < x ? x : y > z ? z : y;
    case "at2":
      return radToGuideAngle(Math.atan2(y, x));
    case "cat2":
      return x * Math.cos(Math.atan2(z, y));
    case "sat2":
      return x * Math.sin(Math.atan2(z, y));
    case "cos":
      return x * Math.cos(guideAngleToRad(y));
    case "sin":
      return x * Math.sin(guideAngleToRad(y));
    case "tan":
      return x * Math.tan(guideAngleToRad(y));
    default:
      return 0;
  }
}

/**
 * Evaluate an ordered list of guide definitions (avLst first, then gdLst)
 * into a lookup of guide name → value. Guides may reference built-ins and
 * any guide defined before them.
 *
 * @param definitions - Guides in document order.
 * @param w - Shape width in the geometry coordinate space.
 * @param h - Shape height in the geometry coordinate space.
 * @param overrides - External adjustment values that take precedence over
 *   `avLst` defaults (e.g. a preset's instance `avLst`).
 */
export function evaluateGuides(
  definitions: GuideDefinition[],
  w: number,
  h: number,
  overrides?: Map<string, number>,
): Map<string, number> {
  const guides = createGuideContext(w, h);
  if (overrides) {
    for (const [name, value] of overrides) {
      guides.set(name, value);
    }
  }
  for (const def of definitions) {
    // An externally supplied adjustment wins over the avLst default formula.
    if (overrides?.has(def.name)) continue;
    guides.set(def.name, evaluateGuideFormula(def.fmla, guides));
  }
  return guides;
}
