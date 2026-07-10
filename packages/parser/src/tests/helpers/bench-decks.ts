/**
 * Parameterized deck generator for benchmarks and profiling.
 *
 * Produces realistic in-memory .pptx packages whose size and complexity are
 * controlled by the caller, so pipeline stages can be measured against
 * small/medium/large workloads reproducibly.
 */
import JSZip from "jszip";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

export interface DeckSpec {
  /** Number of slides. */
  slides: number;
  /** Text boxes per slide (3 paragraphs, 2 runs each). */
  textBoxes?: number;
  /** Preset shapes per slide (mix of geometries, gradient every 3rd). */
  shapes?: number;
  /** Include a 4x6 table on every slide. */
  withTable?: boolean;
}

const PRESETS = ["roundRect", "ellipse", "hexagon", "rightArrow", "star5", "pie"];

function textBoxXml(id: number, slot: number): string {
  const x = 457200 + (slot % 3) * 3048000;
  const y = 457200 + Math.floor(slot / 3) * 1371600;
  const paragraphs = Array.from(
    { length: 3 },
    (_, p) => `<a:p>
<a:pPr indent="0" lvl="${p % 2}" marL="${p * 228600}" algn="${p === 0 ? "ctr" : "l"}"><a:lnSpc><a:spcPct val="115000"/></a:lnSpc><a:buChar char="•"/></a:pPr>
<a:r><a:rPr sz="1800" b="${p === 0 ? 1 : 0}"><a:solidFill><a:schemeClr val="dk1"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>Revenue stream ${slot}.${p} grew substantially in Q${(p % 4) + 1}</a:t></a:r>
<a:r><a:rPr sz="1400" i="1"><a:solidFill><a:srgbClr val="4472C4"><a:lumMod val="75000"/></a:srgbClr></a:solidFill></a:rPr><a:t> with continued momentum across segments</a:t></a:r>
</a:p>`,
  ).join("");

  return `<p:sp>
<p:nvSpPr><p:cNvPr id="${id}" name="Text ${slot}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="2743200" cy="1143000"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
</p:spPr>
<p:txBody><a:bodyPr anchor="t" wrap="square"><a:normAutofit/></a:bodyPr><a:lstStyle/>${paragraphs}</p:txBody>
</p:sp>`;
}

function shapeXml(id: number, slot: number): string {
  const preset = PRESETS[slot % PRESETS.length];
  const x = 6096000 + (slot % 3) * 1828800;
  const y = 457200 + Math.floor(slot / 3) * 1143000;
  const fill =
    slot % 3 === 0
      ? `<a:gradFill><a:gsLst>
<a:gs pos="0"><a:schemeClr val="accent1"><a:tint val="60000"/></a:schemeClr></a:gs>
<a:gs pos="100000"><a:schemeClr val="accent2"><a:shade val="80000"/></a:schemeClr></a:gs>
</a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill>`
      : `<a:solidFill><a:schemeClr val="accent${(slot % 6) + 1}"/></a:solidFill>`;

  return `<p:sp>
<p:nvSpPr><p:cNvPr id="${id}" name="Shape ${slot}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr>
<a:xfrm rot="${(slot % 4) * 900000}"><a:off x="${x}" y="${y}"/><a:ext cx="1524000" cy="914400"/></a:xfrm>
<a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom>
${fill}
<a:ln w="12700"><a:solidFill><a:schemeClr val="dk2"/></a:solidFill></a:ln>
</p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"/><a:t>${preset}</a:t></a:r></a:p></p:txBody>
</p:sp>`;
}

function tableXml(id: number): string {
  const rows = Array.from({ length: 6 }, (_, r) => {
    const cells = Array.from(
      { length: 4 },
      (_, c) =>
        `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1100"/><a:t>Cell r${r}c${c} metric ${r * 4 + c}</a:t></a:r></a:p></a:txBody><a:tcPr><a:solidFill><a:schemeClr val="${r === 0 ? "accent1" : "lt2"}"/></a:solidFill></a:tcPr></a:tc>`,
    ).join("");
    return `<a:tr h="304800">${cells}</a:tr>`;
  }).join("");

  return `<p:graphicFrame>
<p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="457200" y="4114800"/><a:ext cx="5486400" cy="1828800"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
<a:tbl>
<a:tblPr firstRow="1" bandRow="1"/>
<a:tblGrid><a:gridCol w="1371600"/><a:gridCol w="1371600"/><a:gridCol w="1371600"/><a:gridCol w="1371600"/></a:tblGrid>
${rows}
</a:tbl>
</a:graphicData></a:graphic>
</p:graphicFrame>`;
}

function slideXml(index: number, spec: DeckSpec): string {
  const parts: string[] = [];
  let id = 2;
  for (let i = 0; i < (spec.textBoxes ?? 0); i++) parts.push(textBoxXml(id++, i));
  for (let i = 0; i < (spec.shapes ?? 0); i++) parts.push(shapeXml(id++, i + index));
  if (spec.withTable) parts.push(tableXml(id++));

  return (
    XML_DECL +
    `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld>
<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
${parts.join("\n")}
</p:spTree>
</p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`
  );
}

const THEME =
  XML_DECL +
  `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
<a:themeElements>
<a:clrScheme name="Office">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="44546A"/></a:dk2>
<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1>
<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
<a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
<a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink>
<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="Office">
<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
</a:fillStyleLst>
<a:lnStyleLst>
<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
</a:lnStyleLst>
<a:effectStyleLst>
<a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst/></a:effectStyle>
</a:effectStyleLst>
<a:bgFillStyleLst>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
</a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
</a:theme>`;

const SLIDE_MASTER =
  XML_DECL +
  `<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld>
<p:bg><p:bgPr><a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
</p:spTree>
</p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>`;

const SLIDE_LAYOUT =
  XML_DECL +
  `<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
<p:cSld name="Blank">
<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
</p:spTree>
</p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`;

/** Build a deck per spec; returns the zipped package. */
export async function generateDeck(spec: DeckSpec): Promise<ArrayBuffer> {
  const n = spec.slides;

  const contentTypes =
    XML_DECL +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
${Array.from({ length: n }, (_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join("\n")}
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`;

  const rootRels =
    XML_DECL +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;

  const presentation =
    XML_DECL +
    `<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
<p:sldIdLst>
${Array.from({ length: n }, (_, i) => `<p:sldId id="${256 + i}" r:id="rId${2 + i}"/>`).join("\n")}
</p:sldIdLst>
<p:sldSz cx="12192000" cy="6858000"/>
<p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`;

  const presentationRels =
    XML_DECL +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
${Array.from({ length: n }, (_, i) => `<Relationship Id="rId${2 + i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`).join("\n")}
<Relationship Id="rId${2 + n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`;

  const masterRels =
    XML_DECL +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`;

  const layoutRels =
    XML_DECL +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;

  const slideRels =
    XML_DECL +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;

  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes);
  zip.file("_rels/.rels", rootRels);
  zip.file("ppt/presentation.xml", presentation);
  zip.file("ppt/_rels/presentation.xml.rels", presentationRels);
  zip.file("ppt/theme/theme1.xml", THEME);
  zip.file("ppt/slideMasters/slideMaster1.xml", SLIDE_MASTER);
  zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", masterRels);
  zip.file("ppt/slideLayouts/slideLayout1.xml", SLIDE_LAYOUT);
  zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", layoutRels);
  for (let i = 1; i <= n; i++) {
    zip.file(`ppt/slides/slide${i}.xml`, slideXml(i, spec));
    zip.file(`ppt/slides/_rels/slide${i}.xml.rels`, slideRels);
  }
  return zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}

/** Standard specs shared by benchmarks and profiling so numbers are comparable. */
export const DECK_SPECS = {
  small: { slides: 5, textBoxes: 3, shapes: 3, withTable: false },
  medium: { slides: 20, textBoxes: 4, shapes: 6, withTable: true },
  large: { slides: 100, textBoxes: 4, shapes: 6, withTable: true },
} satisfies Record<string, DeckSpec>;
