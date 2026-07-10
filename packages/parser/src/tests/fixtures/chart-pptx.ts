/**
 * In-memory .pptx fixture builders for chart tests.
 *
 * Extends the minimal single-slide package from `minimal-pptx.ts` with a chart
 * part (ppt/charts/chart1.xml), the slide relationship pointing at it, and the
 * [Content_Types].xml override — so chart graphicFrames flow through the real
 * parse pipeline (`parseZip` → `buildPresentation` → `renderSlide`).
 */
import JSZip from "jszip";

import { buildPresentation } from "../../model/presentation";
import { parseXml, SafeXmlNode } from "../../ooxml/xml";
import { parseZip } from "../../ooxml/zip";
import { createRenderContext, RenderContext } from "../../renderer/context";
import { buildPptxWithShapes } from "./minimal-pptx";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

const CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const RELS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/**
 * A graphicFrame hosting a chart reference (rId2 → ../charts/chart1.xml).
 * Placed at 914400,914400 EMU (96,96 px) sized 6096000x4572000 EMU (640x480 px).
 */
export const CHART_GRAPHIC_FRAME = `<p:graphicFrame>
<p:nvGraphicFramePr><p:cNvPr id="5" name="Chart 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="914400" y="914400"/><a:ext cx="6096000" cy="4572000"/></p:xfrm>
<a:graphic><a:graphicData uri="${CHART_NS}">
<c:chart xmlns:c="${CHART_NS}" xmlns:r="${RELS_NS}" r:id="rId2"/>
</a:graphicData></a:graphic>
</p:graphicFrame>`;

const slideRelsWithChart =
  XML_DECL +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`;

const chartContentTypeOverride =
  '<Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>';

/**
 * Build a single-slide .pptx whose slide contains a chart graphicFrame and
 * whose package includes the given chart1.xml part.
 */
export async function buildPptxWithChart(
  chartXml: string,
  options: { omitChartPart?: boolean } = {},
): Promise<ArrayBuffer> {
  const base = await buildPptxWithShapes(CHART_GRAPHIC_FRAME);
  const zip = await JSZip.loadAsync(base);

  const contentTypes = await zip.file("[Content_Types].xml")!.async("string");
  zip.file(
    "[Content_Types].xml",
    contentTypes.replace("</Types>", `${chartContentTypeOverride}\n</Types>`),
  );
  zip.file("ppt/slides/_rels/slide1.xml.rels", slideRelsWithChart);
  if (!options.omitChartPart) {
    zip.file("ppt/charts/chart1.xml", chartXml);
  }
  return zip.generateAsync({ type: "arraybuffer" });
}

/** Wrap chart XML content in a complete c:chartSpace document. */
export function chartSpaceXml(chartChildren: string, chartSpaceExtras = ""): string {
  return (
    XML_DECL +
    `<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${DRAWING_NS}" xmlns:r="${RELS_NS}">
${chartSpaceExtras}
<c:chart>
${chartChildren}
</c:chart>
</c:chartSpace>`
  );
}

/** Build a c:ser element for cartesian charts (strRef categories + numRef values). */
export function seriesXml(
  idx: number,
  name: string,
  categories: string[],
  values: (number | null)[],
  extras = "",
): string {
  const catPts = categories.map((c, i) => `<c:pt idx="${i}"><c:v>${c}</c:v></c:pt>`).join("");
  const valPts = values
    .map((v, i) => (v === null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`))
    .join("");
  return `<c:ser>
<c:idx val="${idx}"/><c:order val="${idx}"/>
<c:tx><c:strRef><c:f>Sheet1!$${String.fromCharCode(66 + idx)}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache></c:strRef></c:tx>
${extras}
<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$${categories.length + 1}</c:f><c:strCache><c:ptCount val="${categories.length}"/>${catPts}</c:strCache></c:strRef></c:cat>
<c:val><c:numRef><c:f>Sheet1!$B$2:$B$${values.length + 1}</c:f><c:numCache><c:ptCount val="${values.length}"/>${valPts}</c:numCache></c:numRef></c:val>
</c:ser>`;
}

/** Standard catAx + valAx pair wired to the given axis ids. */
export function axesXml(catAxId = "111111111", valAxId = "222222222", valAxExtras = ""): string {
  return `<c:catAx>
<c:axId val="${catAxId}"/>
<c:scaling><c:orientation val="minMax"/></c:scaling>
<c:delete val="0"/>
<c:axPos val="b"/>
<c:crossAx val="${valAxId}"/>
</c:catAx>
<c:valAx>
<c:axId val="${valAxId}"/>
<c:scaling><c:orientation val="minMax"/></c:scaling>
<c:delete val="0"/>
<c:axPos val="l"/>
<c:majorGridlines/>
${valAxExtras}
<c:crossAx val="${catAxId}"/>
</c:valAx>`;
}

/**
 * A realistic minimal bar chart part: clustered column chart, 2 series with
 * 3 string categories and 3 numeric values each, catAx + valAx, legend, title.
 */
export const BAR_CHART_XML =
  chartSpaceXml(`<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Sales 2024</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
<c:autoTitleDeleted val="0"/>
<c:plotArea>
<c:layout/>
<c:barChart>
<c:barDir val="col"/>
<c:grouping val="clustered"/>
<c:varyColors val="0"/>
${seriesXml(0, "Series A", ["Q1", "Q2", "Q3"], [10, 20, 30])}
${seriesXml(1, "Series B", ["Q1", "Q2", "Q3"], [5, 15, 25])}
<c:axId val="111111111"/>
<c:axId val="222222222"/>
</c:barChart>
${axesXml()}
</c:plotArea>
<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>
<c:plotVisOnly val="1"/>`);

/**
 * Parse a chart XML fragment (using the standard c:/a:/r: prefixes) and return
 * the wrapper node; access parts through `.child(...)`.
 */
export function parseChartFragment(xml: string): SafeXmlNode {
  return parseXml(
    `<c:root xmlns:c="${CHART_NS}" xmlns:a="${DRAWING_NS}" xmlns:r="${RELS_NS}">${xml}</c:root>`,
  );
}

/**
 * Create a real RenderContext backed by the minimal single-slide package
 * (Office theme accents, blank layout/master) for chart submodule tests.
 */
export async function createChartTestContext(): Promise<RenderContext> {
  const buffer = await buildPptxWithShapes("");
  const presentation = buildPresentation(await parseZip(buffer));
  return createRenderContext(presentation, presentation.slides[0]);
}
