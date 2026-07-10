import { describe, expect, it } from "vitest";

import { getMultiPathPreset, getPresetShapePath, presetShapes } from "../../geometry/presets";

/** All 187 preset names from ECMA-376 ST_ShapeType (dml-main.xsd, 5th edition). */
const ST_SHAPE_TYPE = [
  "accentBorderCallout1",
  "accentBorderCallout2",
  "accentBorderCallout3",
  "accentCallout1",
  "accentCallout2",
  "accentCallout3",
  "actionButtonBackPrevious",
  "actionButtonBeginning",
  "actionButtonBlank",
  "actionButtonDocument",
  "actionButtonEnd",
  "actionButtonForwardNext",
  "actionButtonHelp",
  "actionButtonHome",
  "actionButtonInformation",
  "actionButtonMovie",
  "actionButtonReturn",
  "actionButtonSound",
  "arc",
  "bentArrow",
  "bentConnector2",
  "bentConnector3",
  "bentConnector4",
  "bentConnector5",
  "bentUpArrow",
  "bevel",
  "blockArc",
  "borderCallout1",
  "borderCallout2",
  "borderCallout3",
  "bracePair",
  "bracketPair",
  "callout1",
  "callout2",
  "callout3",
  "can",
  "chartPlus",
  "chartStar",
  "chartX",
  "chevron",
  "chord",
  "circularArrow",
  "cloud",
  "cloudCallout",
  "corner",
  "cornerTabs",
  "cube",
  "curvedConnector2",
  "curvedConnector3",
  "curvedConnector4",
  "curvedConnector5",
  "curvedDownArrow",
  "curvedLeftArrow",
  "curvedRightArrow",
  "curvedUpArrow",
  "decagon",
  "diagStripe",
  "diamond",
  "dodecagon",
  "donut",
  "doubleWave",
  "downArrow",
  "downArrowCallout",
  "ellipse",
  "ellipseRibbon",
  "ellipseRibbon2",
  "flowChartAlternateProcess",
  "flowChartCollate",
  "flowChartConnector",
  "flowChartDecision",
  "flowChartDelay",
  "flowChartDisplay",
  "flowChartDocument",
  "flowChartExtract",
  "flowChartInputOutput",
  "flowChartInternalStorage",
  "flowChartMagneticDisk",
  "flowChartMagneticDrum",
  "flowChartMagneticTape",
  "flowChartManualInput",
  "flowChartManualOperation",
  "flowChartMerge",
  "flowChartMultidocument",
  "flowChartOfflineStorage",
  "flowChartOffpageConnector",
  "flowChartOnlineStorage",
  "flowChartOr",
  "flowChartPredefinedProcess",
  "flowChartPreparation",
  "flowChartProcess",
  "flowChartPunchedCard",
  "flowChartPunchedTape",
  "flowChartSort",
  "flowChartSummingJunction",
  "flowChartTerminator",
  "foldedCorner",
  "frame",
  "funnel",
  "gear6",
  "gear9",
  "halfFrame",
  "heart",
  "heptagon",
  "hexagon",
  "homePlate",
  "horizontalScroll",
  "irregularSeal1",
  "irregularSeal2",
  "leftArrow",
  "leftArrowCallout",
  "leftBrace",
  "leftBracket",
  "leftCircularArrow",
  "leftRightArrow",
  "leftRightArrowCallout",
  "leftRightCircularArrow",
  "leftRightRibbon",
  "leftRightUpArrow",
  "leftUpArrow",
  "lightningBolt",
  "line",
  "lineInv",
  "mathDivide",
  "mathEqual",
  "mathMinus",
  "mathMultiply",
  "mathNotEqual",
  "mathPlus",
  "moon",
  "nonIsoscelesTrapezoid",
  "noSmoking",
  "notchedRightArrow",
  "octagon",
  "parallelogram",
  "pentagon",
  "pie",
  "pieWedge",
  "plaque",
  "plaqueTabs",
  "plus",
  "quadArrow",
  "quadArrowCallout",
  "rect",
  "ribbon",
  "ribbon2",
  "rightArrow",
  "rightArrowCallout",
  "rightBrace",
  "rightBracket",
  "round1Rect",
  "round2DiagRect",
  "round2SameRect",
  "roundRect",
  "rtTriangle",
  "smileyFace",
  "snip1Rect",
  "snip2DiagRect",
  "snip2SameRect",
  "snipRoundRect",
  "squareTabs",
  "star10",
  "star12",
  "star16",
  "star24",
  "star32",
  "star4",
  "star5",
  "star6",
  "star7",
  "star8",
  "straightConnector1",
  "stripedRightArrow",
  "sun",
  "swooshArrow",
  "teardrop",
  "trapezoid",
  "triangle",
  "upArrow",
  "upArrowCallout",
  "upDownArrow",
  "upDownArrowCallout",
  "uturnArrow",
  "verticalScroll",
  "wave",
  "wedgeEllipseCallout",
  "wedgeRectCallout",
  "wedgeRoundRectCallout",
];

function expectValidPathData(d: string, context: string) {
  expect(d, context).toBeTruthy();
  expect(d, `${context} contains NaN`).not.toMatch(/NaN/);
  expect(d, `${context} contains Infinity`).not.toMatch(/Infinity/);
}

describe("preset shape coverage", () => {
  it.each(ST_SHAPE_TYPE)("renders %s at 200x100 without NaN", (name) => {
    const multi = getMultiPathPreset(name, 200, 100);
    if (multi) {
      expect(multi.length, name).toBeGreaterThan(0);
      for (const sub of multi) {
        expectValidPathData(sub.d, `${name} sub-path`);
      }
      return;
    }
    expectValidPathData(getPresetShapePath(name, 200, 100), name);
  });

  it.each(ST_SHAPE_TYPE)("renders %s at square/tall extents without NaN", (name) => {
    for (const [w, h] of [
      [100, 100],
      [50, 300],
    ] as const) {
      const multi = getMultiPathPreset(name, w, h);
      if (multi) {
        for (const sub of multi) {
          expectValidPathData(sub.d, `${name} sub-path at ${w}x${h}`);
        }
        continue;
      }
      expectValidPathData(getPresetShapePath(name, w, h), `${name} at ${w}x${h}`);
    }
  });
});

describe("preset registry", () => {
  it("falls back to a rectangle for unknown presets", () => {
    const d = getPresetShapePath("noSuchShape", 100, 50);
    expect(d).toBe("M0,0 L100,0 L100,50 L0,50 Z");
  });

  it("returns empty path for textNoShape", () => {
    expect(getPresetShapePath("textNoShape", 100, 50)).toBe("");
  });

  it("resolves spec-cased names for lowercase-keyed multi-path entries", () => {
    // multiPathPresets stores some keys lowercase (e.g. "bordercallout1");
    // the spec-cased prst value must still resolve via the lowercase lookup.
    expect(getMultiPathPreset("borderCallout1", 100, 50)).not.toBeNull();
    expect(getMultiPathPreset("curvedDownArrow", 100, 50)).not.toBeNull();
  });
});

describe("adjustment values", () => {
  it("mathPlus reads the spec-named adj1 adjustment (regression)", () => {
    const thin = getPresetShapePath("mathPlus", 200, 200, new Map([["adj1", 10000]]));
    const thick = getPresetShapePath("mathPlus", 200, 200, new Map([["adj1", 50000]]));
    const unset = getPresetShapePath("mathPlus", 200, 200);
    expect(thin).not.toBe(unset);
    expect(thick).not.toBe(unset);
    expect(thin).not.toBe(thick);
  });

  it("roundRect corner radius follows adj", () => {
    const sharp = getPresetShapePath("roundRect", 200, 100, new Map([["adj", 0]]));
    const round = getPresetShapePath("roundRect", 200, 100, new Map([["adj", 50000]]));
    expect(sharp).not.toBe(round);
  });

  it("presetShapes registry does not throw for any registered generator", () => {
    for (const [name, gen] of presetShapes) {
      const d = gen(120, 80);
      expect(typeof d, name).toBe("string");
    }
  });
});
