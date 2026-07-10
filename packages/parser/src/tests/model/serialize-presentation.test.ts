import { beforeAll, describe, expect, it } from "vitest";

import type { PresentationData } from "../../model/presentation";
import { buildPresentation } from "../../model/presentation";
import { serializePresentation } from "../../model/serialize";
import { parseZip } from "../../ooxml/zip-parser";
import { buildPptxWithShapes } from "../helpers/minimal-pptx";

const SHAPES = `<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Title Box"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="914400" y="914400"/><a:ext cx="4572000" cy="914400"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
</p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Serialized title</a:t></a:r></a:p></p:txBody>
</p:sp>`;

let presentation: PresentationData;

beforeAll(async () => {
  const buffer = await buildPptxWithShapes(SHAPES);
  presentation = buildPresentation(await parseZip(buffer));
});

describe("serializePresentation", () => {
  it("produces a JSON-safe structure (no DOM nodes, survives round-trip)", () => {
    const serialized = serializePresentation(presentation);
    const roundTripped = JSON.parse(JSON.stringify(serialized));
    expect(roundTripped.slides).toHaveLength(1);
    expect(roundTripped.width).toBeCloseTo(1280, 6);
    expect(roundTripped.height).toBeCloseTo(720, 6);
  });

  it("serializes shape nodes with geometry, bounds, and text", () => {
    const serialized = serializePresentation(presentation);
    const nodes = serialized.slides[0].nodes;
    expect(nodes.length).toBeGreaterThan(0);
    const shape = nodes.find((n) => n.nodeType === "shape");
    expect(shape).toBeDefined();
    expect(shape?.name).toBe("Title Box");
    expect(shape?.presetGeometry).toBe("rect");
    expect(shape?.textBody?.totalText).toContain("Serialized title");
    expect(shape?.position.x).toBeCloseTo(96, 6);
    expect(shape?.size.w).toBeCloseTo(480, 6);
  });

  it("reports slideCount consistent with slides", () => {
    const serialized = serializePresentation(presentation);
    expect(serialized.slideCount).toBe(serialized.slides.length);
  });
});
