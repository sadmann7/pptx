import { describe, expect, it } from "vitest";

import {
  createGuideContext,
  evaluateGuideFormula,
  evaluateGuides,
  resolveGuideOperand,
} from "../../geometry/guide-evaluator";

describe("createGuideContext", () => {
  it("defines the built-in variables from shape extents", () => {
    const ctx = createGuideContext(200, 100);
    expect(ctx.get("w")).toBe(200);
    expect(ctx.get("h")).toBe(100);
    expect(ctx.get("ss")).toBe(100);
    expect(ctx.get("hc")).toBe(100);
    expect(ctx.get("vc")).toBe(50);
    expect(ctx.get("l")).toBe(0);
    expect(ctx.get("t")).toBe(0);
    expect(ctx.get("r")).toBe(200);
    expect(ctx.get("b")).toBe(100);
    expect(ctx.get("wd32")).toBe(200 / 32);
    expect(ctx.get("ssd16")).toBe(100 / 16);
  });

  it("defines angle constants in 60000ths of a degree", () => {
    const ctx = createGuideContext(1, 1);
    expect(ctx.get("cd2")).toBe(10800000); // 180°
    expect(ctx.get("cd4")).toBe(5400000); // 90°
    expect(ctx.get("3cd4")).toBe(16200000); // 270°
  });
});

describe("resolveGuideOperand", () => {
  const ctx = new Map([["x1", 42]]);

  it("parses numeric literals", () => {
    expect(resolveGuideOperand("100", ctx)).toBe(100);
    expect(resolveGuideOperand("-5", ctx)).toBe(-5);
  });

  it("resolves guide references", () => {
    expect(resolveGuideOperand("x1", ctx)).toBe(42);
  });

  it("falls back to 0 for unknown references and missing tokens", () => {
    expect(resolveGuideOperand("nope", ctx)).toBe(0);
    expect(resolveGuideOperand(undefined, ctx)).toBe(0);
  });
});

describe("evaluateGuideFormula", () => {
  const ctx = createGuideContext(200, 100);

  it.each([
    ["val 12345", 12345],
    ["*/ 6 7 2", 21],
    ["+- 10 5 3", 12],
    ["+/ 10 6 4", 4],
    ["abs -42", 42],
    ["min 3 9", 3],
    ["max 3 9", 9],
    ["sqrt 81", 9],
    ["mod 3 4 0", 5],
  ] as const)("evaluates %s", (fmla, expected) => {
    expect(evaluateGuideFormula(fmla, ctx)).toBeCloseTo(expected, 9);
  });

  it("evaluates ?: as x > 0 ? y : z", () => {
    expect(evaluateGuideFormula("?: 1 10 20", ctx)).toBe(10);
    expect(evaluateGuideFormula("?: 0 10 20", ctx)).toBe(20);
    expect(evaluateGuideFormula("?: -1 10 20", ctx)).toBe(20);
  });

  it("evaluates pin as clamp of y into [x, z]", () => {
    expect(evaluateGuideFormula("pin 5 3 9", ctx)).toBe(5);
    expect(evaluateGuideFormula("pin 5 7 9", ctx)).toBe(7);
    expect(evaluateGuideFormula("pin 5 12 9", ctx)).toBe(9);
  });

  it("evaluates trig ops with angles in 60000ths of a degree", () => {
    expect(evaluateGuideFormula("cos 100 3600000", ctx)).toBeCloseTo(50, 9); // cos 60°
    expect(evaluateGuideFormula("sin 100 5400000", ctx)).toBeCloseTo(100, 9); // sin 90°
    expect(evaluateGuideFormula("tan 100 2700000", ctx)).toBeCloseTo(100, 9); // tan 45°
  });

  it("evaluates at2/cat2/sat2 with OOXML argument order", () => {
    expect(evaluateGuideFormula("at2 1 1", ctx)).toBeCloseTo(2700000, 3); // atan2 → 45°
    expect(evaluateGuideFormula("cat2 100 1 1", ctx)).toBeCloseTo(100 * Math.cos(Math.PI / 4), 9);
    expect(evaluateGuideFormula("sat2 100 1 1", ctx)).toBeCloseTo(100 * Math.sin(Math.PI / 4), 9);
  });

  it("returns 0 for division by zero and unknown operators", () => {
    expect(evaluateGuideFormula("*/ 6 7 0", ctx)).toBe(0);
    expect(evaluateGuideFormula("+/ 6 7 0", ctx)).toBe(0);
    expect(evaluateGuideFormula("frobnicate 1 2 3", ctx)).toBe(0);
  });
});

describe("evaluateGuides", () => {
  // The arrow geometry from the gdlst-test.pptx regression file: 3in x 2in.
  const W = 2743200;
  const H = 1828800;
  const defs = [
    { name: "adj", fmla: "val 40000" },
    { name: "headW", fmla: "*/ w adj 100000" },
    { name: "x1", fmla: "+- w 0 headW" },
    { name: "y1", fmla: "*/ h 25000 100000" },
    { name: "y2", fmla: "+- h 0 y1" },
    { name: "tailNotch", fmla: "*/ w 15000 100000" },
  ];

  it("evaluates guides in order, allowing references to earlier guides", () => {
    const guides = evaluateGuides(defs, W, H);
    expect(guides.get("adj")).toBe(40000);
    expect(guides.get("headW")).toBe(1097280);
    expect(guides.get("x1")).toBe(1645920);
    expect(guides.get("y1")).toBe(457200);
    expect(guides.get("y2")).toBe(1371600);
    expect(guides.get("tailNotch")).toBe(411480);
  });

  it("lets overrides take precedence over avLst default formulas", () => {
    const guides = evaluateGuides(defs, W, H, new Map([["adj", 20000]]));
    expect(guides.get("adj")).toBe(20000);
    expect(guides.get("headW")).toBe((W * 20000) / 100000);
  });
});
