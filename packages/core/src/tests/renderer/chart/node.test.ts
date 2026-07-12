import { beforeAll, describe, expect, it } from "vitest";

import type { ChartNodeData } from "../../../model/nodes/chart";
import { parseChartNode } from "../../../model/nodes/chart";
import type { PresentationData } from "../../../model/presentation";
import { buildPresentation, materializeAllSlides } from "../../../model/presentation";
import type { RelEntry } from "../../../ooxml/rel";
import { parseXml } from "../../../ooxml/xml";
import { readPptx } from "../../../ooxml/zip";
import { BAR_CHART_XML, buildPptxWithChart } from "../../fixtures/chart-pptx";

const CHART_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

function chartFrameXml(chartRef: string): string {
  return `<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
<p:nvGraphicFramePr><p:cNvPr id="7" name="My Chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="914400" y="0"/><a:ext cx="914400" cy="1828800"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
${chartRef}
</a:graphicData></a:graphic>
</p:graphicFrame>`;
}

describe("chart node via full pipeline", () => {
  let presentation: PresentationData;

  beforeAll(async () => {
    const buffer = await buildPptxWithChart(BAR_CHART_XML);
    presentation = buildPresentation(await readPptx(buffer));
    materializeAllSlides(presentation);
  });

  it("parses the chart graphicFrame into a chart node", () => {
    const chart = presentation.slides[0].nodes.find(
      (n): n is ChartNodeData => n.nodeType === "chart",
    );
    expect(chart).toBeDefined();
    expect(chart!.id).toBe("5");
    expect(chart!.name).toBe("Chart 1");
    // 914400 EMU = 96px, 6096000x4572000 EMU = 640x480 px
    expect(chart!.position).toEqual({ x: 96, y: 96 });
    expect(chart!.size).toEqual({ w: 640, h: 480 });
  });

  it("resolves the rId2 relationship to the chart part path", () => {
    const chart = presentation.slides[0].nodes.find(
      (n): n is ChartNodeData => n.nodeType === "chart",
    );
    expect(chart!.chartPath).toBe("ppt/charts/chart1.xml");
  });

  it("registers the chart part XML in the presentation charts map", () => {
    const chartXml = presentation.charts.get("ppt/charts/chart1.xml");
    expect(chartXml).toBeDefined();
    expect(chartXml!.localName).toBe("chartSpace");
    expect(chartXml!.child("chart").child("plotArea").child("barChart").exists()).toBe(true);
  });
});

describe("parseChartNode", () => {
  const rels = new Map<string, RelEntry>([
    ["rId5", { type: CHART_REL_TYPE, target: "../charts/chart3.xml" }],
  ]);

  it("parses base props and resolves the chart path relative to the slide", () => {
    const frame = parseXml(
      chartFrameXml(
        `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId5"/>`,
      ),
    );
    const node = parseChartNode(frame, rels, "ppt/slides/slide2.xml");
    expect(node).toBeDefined();
    expect(node!.nodeType).toBe("chart");
    expect(node!.chartPath).toBe("ppt/charts/chart3.xml");
    expect(node!.name).toBe("My Chart");
    expect(node!.position).toEqual({ x: 96, y: 0 });
    expect(node!.size).toEqual({ w: 96, h: 192 });
  });

  it("falls back to a plain id attribute when r:id is absent", () => {
    const frame = parseXml(
      chartFrameXml(
        `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" id="rId5"/>`,
      ),
    );
    const node = parseChartNode(frame, rels, "ppt/slides/slide2.xml");
    expect(node?.chartPath).toBe("ppt/charts/chart3.xml");
  });

  it("returns undefined when the graphicData has no chart element", () => {
    const frame = parseXml(chartFrameXml(""));
    expect(parseChartNode(frame, rels, "ppt/slides/slide1.xml")).toBeUndefined();
  });

  it("returns undefined when the relationship id is not in the slide rels", () => {
    const frame = parseXml(
      chartFrameXml(
        `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId99"/>`,
      ),
    );
    expect(parseChartNode(frame, rels, "ppt/slides/slide1.xml")).toBeUndefined();
  });
});
