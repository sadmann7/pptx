import { describe, expect, it } from "vitest";

import { parseXml } from "../parser/xml-parser";
import { renderCustomGeometry } from "../shapes/custom-geometry";

const A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

function custGeom(inner: string): ReturnType<typeof parseXml> {
  return parseXml(`<a:custGeom ${A_NS}>${inner}</a:custGeom>`);
}

/** Parse "M1,2 L3,4 …" into [["M",1,2], ["L",3,4], …] for numeric assertions. */
function pathCommands(d: string): Array<[string, ...number[]]> {
  const out: Array<[string, ...number[]]> = [];
  for (const m of d.matchAll(/([MLQCAZ])([^MLQCAZ]*)/g)) {
    const nums = (m[2].match(/-?\d*\.?\d+(?:e[-+]?\d+)?/gi) ?? []).map(Number);
    out.push([m[1], ...nums]);
  }
  return out;
}

describe("renderCustomGeometry with numeric coordinates", () => {
  it("renders and scales a simple triangle from path space to pixels", () => {
    const geom = custGeom(`
      <a:pathLst>
        <a:path w="200" h="100">
          <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
          <a:lnTo><a:pt x="200" y="0"/></a:lnTo>
          <a:lnTo><a:pt x="100" y="100"/></a:lnTo>
          <a:close/>
        </a:path>
      </a:pathLst>`);

    // Render at 2x path width, 3x path height.
    const d = renderCustomGeometry(geom, 400, 300);
    expect(pathCommands(d)).toEqual([["M", 0, 0], ["L", 400, 0], ["L", 200, 300], ["Z"]]);
  });

  it("returns empty string when pathLst is missing", () => {
    expect(renderCustomGeometry(custGeom(""), 100, 100)).toBe("");
  });
});

describe("renderCustomGeometry with gdLst guides", () => {
  // Regression for the gdlst-test.pptx arrow: guide-referencing points used to
  // collapse to 0 because they were read with numAttr().
  const W_EMU = 2743200;
  const H_EMU = 1828800;

  const arrow = custGeom(`
    <a:avLst><a:gd name="adj" fmla="val 40000"/></a:avLst>
    <a:gdLst>
      <a:gd name="headW" fmla="*/ w adj 100000"/>
      <a:gd name="x1" fmla="+- w 0 headW"/>
      <a:gd name="vc2" fmla="+/ h 0 2"/>
      <a:gd name="y1" fmla="*/ h 25000 100000"/>
      <a:gd name="y2" fmla="+- h 0 y1"/>
      <a:gd name="tailNotch" fmla="*/ w 15000 100000"/>
    </a:gdLst>
    <a:pathLst>
      <a:path w="${W_EMU}" h="${H_EMU}">
        <a:moveTo><a:pt x="0" y="y1"/></a:moveTo>
        <a:lnTo><a:pt x="x1" y="y1"/></a:lnTo>
        <a:lnTo><a:pt x="x1" y="0"/></a:lnTo>
        <a:lnTo><a:pt x="w" y="vc2"/></a:lnTo>
        <a:lnTo><a:pt x="x1" y="h"/></a:lnTo>
        <a:lnTo><a:pt x="x1" y="y2"/></a:lnTo>
        <a:lnTo><a:pt x="0" y="y2"/></a:lnTo>
        <a:lnTo><a:pt x="tailNotch" y="vc2"/></a:lnTo>
        <a:close/>
      </a:path>
    </a:pathLst>`);

  it("resolves guide-referencing points to the correct geometry", () => {
    // Render at native EMU scale (width == pathW) so values map 1:1.
    const d = renderCustomGeometry(arrow, W_EMU, H_EMU, { w: W_EMU, h: H_EMU });
    expect(pathCommands(d)).toEqual([
      ["M", 0, 457200],
      ["L", 1645920, 457200],
      ["L", 1645920, 0],
      ["L", 2743200, 914400],
      ["L", 1645920, 1828800],
      ["L", 1645920, 1371600],
      ["L", 0, 1371600],
      ["L", 411480, 914400],
      ["Z"],
    ]);
  });

  it("scales guide-derived points into pixel space", () => {
    // 2743200x1828800 EMU at 96dpi = 288x192 px.
    const d = renderCustomGeometry(arrow, 288, 192, { w: W_EMU, h: H_EMU });
    const cmds = pathCommands(d);
    expect(cmds[0]).toEqual(["M", 0, 48]); // y1 = 25% of 192
    expect(cmds[3]).toEqual(["L", 288, 96]); // w, vc
    expect(cmds[7]).toEqual(["L", 43.2, 96]); // tailNotch = 15% of 288
  });

  it("converts guide values between shape space and a differing path space", () => {
    // Path declares a 21600 local grid while guides evaluate in shape EMU space.
    const geom = custGeom(`
      <a:gdLst><a:gd name="gx" fmla="*/ w 1 2"/></a:gdLst>
      <a:pathLst>
        <a:path w="21600" h="21600">
          <a:moveTo><a:pt x="gx" y="0"/></a:moveTo>
          <a:lnTo><a:pt x="21600" y="21600"/></a:lnTo>
          <a:close/>
        </a:path>
      </a:pathLst>`);

    // gx = w/2 in EMU; in path space that's 10800, and at 216px target = 108px.
    const d = renderCustomGeometry(geom, 216, 216, { w: 914400, h: 914400 });
    const cmds = pathCommands(d);
    expect(cmds[0]).toEqual(["M", 108, 0]);
    expect(cmds[1]).toEqual(["L", 216, 216]);
  });

  it("resolves guide-referencing arcTo radii and angles", () => {
    // Half-circle: start at (0, vc), sweep 180° with radii wd2/hd2.
    const geom = custGeom(`
      <a:gdLst>
        <a:gd name="stAngG" fmla="val 10800000"/>
        <a:gd name="swAngG" fmla="val 10800000"/>
      </a:gdLst>
      <a:pathLst>
        <a:path w="200" h="200">
          <a:moveTo><a:pt x="0" y="100"/></a:moveTo>
          <a:arcTo wR="wd2" hR="hd2" stAng="stAngG" swAng="swAngG"/>
          <a:close/>
        </a:path>
      </a:pathLst>`);

    const d = renderCustomGeometry(geom, 200, 200, { w: 200, h: 200 });
    const cmds = pathCommands(d);
    const arc = cmds[1];
    expect(arc[0]).toBe("A");
    expect(arc[1]).toBeCloseTo(100, 6); // rx = wd2
    expect(arc[2]).toBeCloseTo(100, 6); // ry = hd2
    // 180° sweep from (0,100) around center (100,100) ends at (200,100).
    expect(arc[6]).toBeCloseTo(200, 6);
    expect(arc[7]).toBeCloseTo(100, 6);
  });
});
