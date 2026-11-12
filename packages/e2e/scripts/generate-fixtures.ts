/**
 * Regenerates the committed fixture decks under fixtures/ (other files there
 * are left untouched):
 *
 *   pnpm fixtures
 *
 * - basic.pptx          3 slides of colored shapes/text (smoke + navigation)
 * - bom-rels.pptx       rels and [Content_Types] parts prefixed with a UTF-8 BOM
 * - nested-charts.pptx  charts under ppt/slides/charts/ with literal data
 * - tables-groups.pptx  unequal tblGrid widths; nested groups with chOff/chExt
 * - table-borders.pptx  styleless table grid; per-cell tcPr borders that disagree
 * - pictures.pptx       one embedded PNG picture (media blob URL lifecycle)
 */
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { buildCustomPptx } from "../../core/src/tests/fixtures/fixture-extras";

const FIXTURES_DIR = join(dirname(fileURLToPath(import.meta.url)), "..", "fixtures");

const BOM = "\uFEFF";

const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const REL_TYPE_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

// The shared fixture helper declares a "rel" Default instead of "rels", which
// PowerPoint rejects as corrupt, and the oracle needs decks PowerPoint opens.
function contentTypesXml(slideCount: number, chartParts: string[] = []): string {
  const slideOverrides = Array.from(
    { length: slideCount },
    (_, i) =>
      `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
  ).join("");
  const chartOverrides = chartParts
    .map(
      (part) =>
        `<Override PartName="${part}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
    )
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Types xmlns="${CT_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>${slideOverrides}${chartOverrides}<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`;
}

function rectWithText(id: number, color: string, text: string, offY = 914400): string {
  return `<p:sp>
<p:nvSpPr><p:cNvPr id="${id}" name="Rect ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="914400" y="${offY}"/><a:ext cx="6096000" cy="1828800"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>
</p:spPr>
<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>
<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="3200" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>${text}</a:t></a:r></a:p>
</p:txBody>
</p:sp>`;
}

function chartGraphicFrame(id: number, relId: string): string {
  return `<p:graphicFrame>
<p:nvGraphicFramePr><p:cNvPr id="${id}" name="Chart ${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="914400" y="914400"/><a:ext cx="9144000" cy="4572000"/></p:xfrm>
<a:graphic><a:graphicData uri="${CHART_NS}"><c:chart xmlns:c="${CHART_NS}" r:id="${relId}"/></a:graphicData></a:graphic>
</p:graphicFrame>`;
}

const PIXEL_PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==",
  "base64",
);

function picture(id: number, relId: string): string {
  return `<p:pic>
<p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
<p:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
</p:pic>`;
}

// An absolute chart target, as Open XML SDK writes it.
function slideRelsWithChart(chartTarget: string): string {
  return `<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="${RELS_NS}"><Relationship Type="${REL_TYPE_BASE}/slideLayout" Target="/ppt/slideLayouts/slideLayout1.xml" Id="rId1" /><Relationship Type="${REL_TYPE_BASE}/chart" Target="${chartTarget}" Id="rId2" /></Relationships>`;
}

function strLit(values: string[]): string {
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("");
  return `<c:strLit><c:ptCount val="${values.length}"/>${pts}</c:strLit>`;
}

function numLit(values: number[]): string {
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("");
  return `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${pts}</c:numLit>`;
}

const NO_FILL_LINE = `<a:ln w="0" xmlns:a="${DRAWING_NS}"><a:noFill/><a:prstDash val="solid"/></a:ln>`;

function doughnutChartXml(): string {
  const dPt = (idx: number, color: string) =>
    `<c:dPt><c:idx val="${idx}"/><c:spPr><a:solidFill xmlns:a="${DRAWING_NS}"><a:srgbClr val="${color}"/></a:solidFill></c:spPr></c:dPt>`;
  return `<?xml version="1.0" encoding="utf-8"?><c:chartSpace xmlns:c="${CHART_NS}"><c:chart><c:plotArea><c:doughnutChart><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Matches</c:v></c:tx>
${dPt(0, "1C1C1C")}${dPt(1, "91832A")}${dPt(2, "F5D200")}
<c:cat>${strLit(["GROUP", "KNOCKOUT", "FINAL"])}</c:cat>
<c:val>${numLit([72, 32, 1])}</c:val>
</c:ser><c:firstSliceAng val="270"/><c:holeSize val="68"/></c:doughnutChart></c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
}

function barChartXml(): string {
  const axSpPr = `<c:spPr>${NO_FILL_LINE}</c:spPr>`;
  return `<?xml version="1.0" encoding="utf-8"?><c:chartSpace xmlns:c="${CHART_NS}"><c:chart><c:plotArea><c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/><c:varyColors val="0"/>
<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Count</c:v></c:tx>
<c:spPr><a:solidFill xmlns:a="${DRAWING_NS}"><a:srgbClr val="F5D200"/></a:solidFill></c:spPr>
<c:cat>${strLit(["GROUP", "R16", "QF", "SF", "FINAL"])}</c:cat>
<c:val>${numLit([72, 8, 4, 2, 1])}</c:val>
</c:ser><c:gapWidth val="42"/><c:axId val="111"/><c:axId val="222"/></c:barChart>
<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorTickMark val="none"/>${axSpPr}<c:crossAx val="222"/></c:catAx>
<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:majorTickMark val="none"/>${axSpPr}<c:crossAx val="111"/></c:valAx>
<c:spPr><a:solidFill xmlns:a="${DRAWING_NS}"><a:srgbClr val="1C1C1C"/></a:solidFill>${NO_FILL_LINE}</c:spPr>
</c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
}

// Unequal column widths, so a renderer that ignores tblGrid and splits the frame
// evenly produces a visible diff.
function tableFrame(id: number): string {
  const cell = (text: string, fill: string) =>
    `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1400"/><a:t>${text}</a:t></a:r></a:p></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill></a:tcPr></a:tc>`;
  const header = `<a:tr h="914400">${cell("Stage", "4472C4")}${cell("Matches", "4472C4")}${cell("Days", "4472C4")}</a:tr>`;
  const data = `<a:tr h="914400">${cell("Group", "E7E6E6")}${cell("72", "E7E6E6")}${cell("16", "E7E6E6")}</a:tr>`;
  return `<p:graphicFrame>
<p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="914400" y="914400"/><a:ext cx="9144000" cy="1828800"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
<a:tbl>
<a:tblPr firstRow="1"/>
<a:tblGrid><a:gridCol w="4572000"/><a:gridCol w="3048000"/><a:gridCol w="1524000"/></a:tblGrid>
${header}
${data}
</a:tbl>
</a:graphicData></a:graphic>
</p:graphicFrame>`;
}

// Shared by both border slides so a spec can find a grid line by arithmetic. The
// bottom row and right column are textless, so a pixel scan along their centre
// lines crosses grid lines and cell fills only.
const BORDER_TABLE_GEOMETRY = {
  x: 96,
  y: 96,
  colWidth: 320,
  rowHeight: 96,
  cols: 3,
  rows: 3,
};

function borderTableFrame(id: number, rows: string[]): string {
  const { x, y, colWidth, rowHeight, cols } = BORDER_TABLE_GEOMETRY;
  const emu = (px: number) => Math.round((px / 96) * 914400);
  const gridCols = Array.from({ length: cols }, () => `<a:gridCol w="${emu(colWidth)}"/>`).join("");
  const body = rows.map((cells) => `<a:tr h="${emu(rowHeight)}">${cells}</a:tr>`).join("");
  return `<p:graphicFrame>
<p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(colWidth * cols)}" cy="${emu(rowHeight * rows.length)}"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
<a:tbl>
<a:tblPr/>
<a:tblGrid>${gridCols}</a:tblGrid>
${body}
</a:tbl>
</a:graphicData></a:graphic>
</p:graphicFrame>`;
}

function borderCell(text: string, fill: string, textColor: string, tcBdr = ""): string {
  const run = text
    ? `<a:r><a:rPr lang="en-US" sz="1400"><a:solidFill><a:srgbClr val="${textColor}"/></a:solidFill></a:rPr><a:t>${text}</a:t></a:r>`
    : "<a:endParaRPr/>";
  return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p>${run}</a:p></a:txBody><a:tcPr>${tcBdr}<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill></a:tcPr></a:tc>`;
}

// No a:tableStyleId, so the grid comes from the built-in fallback style. Fills
// alternate dark and light, so a grid line composited against the slide instead
// of the cells it separates is visibly wrong.
function styleFallbackGridSlide(): string {
  const dark = (text: string) => borderCell(text, "1C1C1C", "FFFFFF");
  const light = (text: string) => borderCell(text, "F2F2F2", "1C1C1C");
  const blank = () => borderCell("", "FFFFFF", "1C1C1C");
  return borderTableFrame(2, [
    `${dark("Stage")}${dark("Matches")}${dark("")}`,
    `${light("Group")}${light("72")}${light("")}`,
    `${blank()}${blank()}${blank()}`,
  ]);
}

function tcLine(side: "lnL" | "lnR" | "lnT" | "lnB", widthPt: number, color: string): string {
  return `<a:${side} w="${Math.round(widthPt * 12700)}" cap="flat"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:prstDash val="solid"/></a:${side}>`;
}

// The cells' top edges disagree, so the outline cannot be painted as one
// rectangle on the table and has to stay on the cells.
function explicitBorderSlide(): string {
  const cell = (text: string, top: string) =>
    borderCell(
      text,
      "FFFFFF",
      "1C1C1C",
      [
        tcLine("lnT", 1, top),
        tcLine("lnB", 1, "0070C0"),
        tcLine("lnL", 1, "0070C0"),
        tcLine("lnR", 1, "0070C0"),
      ].join(""),
    );
  return borderTableFrame(2, [
    `${cell("A1", "C00000")}${cell("B1", "00B050")}${cell("C1", "00B050")}`,
    `${cell("A2", "0070C0")}${cell("B2", "0070C0")}${cell("C2", "0070C0")}`,
    `${cell("", "0070C0")}${cell("", "0070C0")}${cell("", "0070C0")}`,
  ]);
}

function plainRect(
  id: number,
  x: number,
  y: number,
  cx: number,
  cy: number,
  color: string,
): string {
  return `<p:sp>
<p:nvSpPr><p:cNvPr id="${id}" name="Rect ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>
</p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>`;
}

// The child space is half the group's size, so children render at 2x scale.
function groupWithChildren(id: number): string {
  return `<p:grpSp>
<p:nvGrpSpPr><p:cNvPr id="${id}" name="Group ${id}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr>
<a:xfrm>
<a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1828800"/>
<a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="914400"/>
</a:xfrm>
</p:grpSpPr>
${plainRect(id + 1, 0, 0, 914400, 914400, "C00000")}
${plainRect(id + 2, 914400, 0, 914400, 914400, "00B050")}
</p:grpSp>`;
}

async function buildBasic(): Promise<ArrayBuffer> {
  return buildCustomPptx({
    slides: [
      rectWithText(2, "C00000", "Slide one"),
      rectWithText(2, "0070C0", "Slide two"),
      rectWithText(2, "00B050", "Slide three"),
    ],
    contentTypesXml: contentTypesXml(3),
  });
}

async function buildBomRels(): Promise<ArrayBuffer> {
  // Prefix every relationship part and [Content_Types].xml with a UTF-8 BOM,
  // as Open XML SDK does. JSZip encodes the leading U+FEFF as EF BB BF.
  const relTypeBase = REL_TYPE_BASE;
  const bomContentTypes =
    BOM +
    `<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`;
  const bomRootRels =
    BOM +
    `<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="${RELS_NS}"><Relationship Id="rId1" Type="${relTypeBase}/officeDocument" Target="/ppt/presentation.xml"/></Relationships>`;
  const bomPresentationRels =
    BOM +
    `<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="${RELS_NS}"><Relationship Id="rId1" Type="${relTypeBase}/slideMaster" Target="/ppt/slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="${relTypeBase}/slide" Target="/ppt/slides/slide1.xml"/><Relationship Id="rId3" Type="${relTypeBase}/theme" Target="/ppt/theme/theme1.xml"/></Relationships>`;
  const bomSlideRels =
    BOM +
    `<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="${RELS_NS}"><Relationship Id="rId1" Type="${relTypeBase}/slideLayout" Target="/ppt/slideLayouts/slideLayout1.xml"/></Relationships>`;

  return buildCustomPptx({
    slides: [rectWithText(2, "7030A0", "BOM deck renders")],
    extraFiles: {
      "[Content_Types].xml": bomContentTypes,
      "_rels/.rels": bomRootRels,
      "ppt/_rels/presentation.xml.rels": bomPresentationRels,
      "ppt/slides/_rels/slide1.xml.rels": bomSlideRels,
    },
  });
}

async function buildNestedCharts(): Promise<ArrayBuffer> {
  return buildCustomPptx({
    slides: [chartGraphicFrame(3, "rId2"), chartGraphicFrame(3, "rId2")],
    contentTypesXml: contentTypesXml(2, [
      "/ppt/slides/charts/chart1.xml",
      "/ppt/slides/charts/chart2.xml",
    ]),
    extraFiles: {
      "ppt/slides/charts/chart1.xml": doughnutChartXml(),
      "ppt/slides/charts/chart2.xml": barChartXml(),
      "ppt/slides/_rels/slide1.xml.rels": slideRelsWithChart("/ppt/slides/charts/chart1.xml"),
      "ppt/slides/_rels/slide2.xml.rels": slideRelsWithChart("/ppt/slides/charts/chart2.xml"),
    },
  });
}

async function buildTablesGroups(): Promise<ArrayBuffer> {
  return buildCustomPptx({
    slides: [tableFrame(2), groupWithChildren(2)],
    contentTypesXml: contentTypesXml(2),
  });
}

async function buildTableBorders(): Promise<ArrayBuffer> {
  return buildCustomPptx({
    slides: [styleFallbackGridSlide(), explicitBorderSlide()],
    contentTypesXml: contentTypesXml(2),
  });
}

async function buildPictures(): Promise<ArrayBuffer> {
  return buildCustomPptx({
    slides: [picture(2, "rId2")],
    contentTypesXml: contentTypesXml(1),
    extraFiles: {
      "ppt/media/image1.png": PIXEL_PNG,
      "ppt/slides/_rels/slide1.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${RELS_NS}"><Relationship Id="rId1" Type="${REL_TYPE_BASE}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL_TYPE_BASE}/image" Target="../media/image1.png"/></Relationships>`,
    },
  });
}

async function main(): Promise<void> {
  mkdirSync(FIXTURES_DIR, { recursive: true });

  const decks: Array<[string, Promise<ArrayBuffer>]> = [
    ["basic.pptx", buildBasic()],
    ["bom-rels.pptx", buildBomRels()],
    ["nested-charts.pptx", buildNestedCharts()],
    ["tables-groups.pptx", buildTablesGroups()],
    ["table-borders.pptx", buildTableBorders()],
    ["pictures.pptx", buildPictures()],
  ];

  for (const [name, bufferPromise] of decks) {
    const buffer = Buffer.from(await bufferPromise);
    writeFileSync(join(FIXTURES_DIR, name), buffer);
    console.log(`wrote fixtures/${name} (${buffer.byteLength} bytes)`);
  }
}

void main();
