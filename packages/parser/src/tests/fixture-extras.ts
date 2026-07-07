/**
 * Extended in-memory .pptx fixture builder for plumbing/utility tests.
 *
 * Complements ./minimal-pptx.ts (which stays untouched) with knobs the
 * package-plumbing tests need: multiple slides, custom master clrMap,
 * layout clrMapOvr, theme font typefaces, and arbitrary extra zip entries
 * (media, embedded fonts, tableStyles.xml, ...).
 */
import JSZip from "jszip";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

export interface CustomPptxOptions {
  /** Shapes XML (`p:sp`/`p:graphicFrame`/...) per slide; one entry per slide. */
  slides?: string[];
  /** Attribute string for the master's `p:clrMap` element. */
  masterClrMapAttrs?: string;
  /** Full `<p:bg>...</p:bg>` element for the slide master. */
  masterBgXml?: string;
  /** Full `<p:clrMapOvr>...</p:clrMapOvr>` element for the slide layout. */
  layoutClrMapOvrXml?: string;
  /** Shapes XML (`p:sp`/...) injected into the slide layout's spTree. */
  layoutShapesXml?: string;
  /** Latin typeface for the theme's majorFont (default "Calibri Light"). */
  majorLatin?: string;
  /** Latin typeface for the theme's minorFont (default "Calibri"). */
  minorLatin?: string;
  /** Extra XML injected into the theme's majorFont (e.g. script fonts). */
  majorFontExtraXml?: string;
  /** Extra zip entries: path → content (added verbatim). */
  extraFiles?: Record<string, string | Uint8Array>;
  /** Replace [Content_Types].xml wholesale (e.g. with malformed text). */
  contentTypesXml?: string;
}

const DEFAULT_CLR_MAP =
  'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"';

const DEFAULT_MASTER_BG = `<p:bg><p:bgPr><a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`;

function contentTypesFor(slideCount: number): string {
  const slideOverrides = Array.from(
    { length: slideCount },
    (_, i) =>
      `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
  ).join("\n");
  return (
    XML_DECL +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Default Extension="fntdata" ContentType="application/x-fontdata"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
${slideOverrides}
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`
  );
}

const rootRels =
  XML_DECL +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;

function presentationFor(slideCount: number): string {
  const sldIds = Array.from(
    { length: slideCount },
    (_, i) => `<p:sldId id="${256 + i}" r:id="rId${2 + i}"/>`,
  ).join("");
  return (
    XML_DECL +
    `<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
<p:sldIdLst>${sldIds}</p:sldIdLst>
<p:sldSz cx="12192000" cy="6858000"/>
<p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`
  );
}

function presentationRelsFor(slideCount: number): string {
  const slideRels = Array.from(
    { length: slideCount },
    (_, i) =>
      `<Relationship Id="rId${2 + i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`,
  ).join("\n");
  return (
    XML_DECL +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
${slideRels}
<Relationship Id="rId${2 + slideCount}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`
  );
}

function themeFor(options: CustomPptxOptions): string {
  const majorLatin = options.majorLatin ?? "Calibri Light";
  const minorLatin = options.minorLatin ?? "Calibri";
  const majorExtra = options.majorFontExtraXml ?? "";
  return (
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
<a:majorFont><a:latin typeface="${majorLatin}"/><a:ea typeface=""/><a:cs typeface=""/>${majorExtra}</a:majorFont>
<a:minorFont><a:latin typeface="${minorLatin}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
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
</a:theme>`
  );
}

function slideMasterFor(options: CustomPptxOptions): string {
  const clrMap = options.masterClrMapAttrs ?? DEFAULT_CLR_MAP;
  const bg = options.masterBgXml ?? DEFAULT_MASTER_BG;
  return (
    XML_DECL +
    `<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld>
${bg}
<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
</p:spTree>
</p:cSld>
<p:clrMap ${clrMap}/>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>`
  );
}

const slideMasterRels =
  XML_DECL +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`;

function slideLayoutFor(options: CustomPptxOptions): string {
  const clrMapOvr =
    options.layoutClrMapOvrXml ?? `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>`;
  return (
    XML_DECL +
    `<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
<p:cSld name="Blank">
<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
${options.layoutShapesXml ?? ""}
</p:spTree>
</p:cSld>
${clrMapOvr}
</p:sldLayout>`
  );
}

const slideLayoutRels =
  XML_DECL +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;

const slideRels =
  XML_DECL +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;

function slideFor(shapesXml: string): string {
  return (
    XML_DECL +
    `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld>
<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
${shapesXml}
</p:spTree>
</p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`
  );
}

/** Build a .pptx with the given customizations; defaults match minimal-pptx. */
export async function buildCustomPptx(options: CustomPptxOptions = {}): Promise<ArrayBuffer> {
  const slides = options.slides ?? [""];

  const zip = new JSZip();
  zip.file("[Content_Types].xml", options.contentTypesXml ?? contentTypesFor(slides.length));
  zip.file("_rels/.rels", rootRels);
  zip.file("ppt/presentation.xml", presentationFor(slides.length));
  zip.file("ppt/_rels/presentation.xml.rels", presentationRelsFor(slides.length));
  zip.file("ppt/theme/theme1.xml", themeFor(options));
  zip.file("ppt/slideMasters/slideMaster1.xml", slideMasterFor(options));
  zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMasterRels);
  zip.file("ppt/slideLayouts/slideLayout1.xml", slideLayoutFor(options));
  zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slideLayoutRels);
  for (let i = 0; i < slides.length; i++) {
    zip.file(`ppt/slides/slide${i + 1}.xml`, slideFor(slides[i]));
    zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, slideRels);
  }
  for (const [path, content] of Object.entries(options.extraFiles ?? {})) {
    zip.file(path, content);
  }
  return zip.generateAsync({ type: "arraybuffer" });
}

/** A tiny (but magic-byte-valid) PNG payload for media tests. */
export function fakePngBytes(length = 64): Uint8Array {
  const bytes = new Uint8Array(length);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  for (let i = 8; i < length; i++) bytes[i] = i % 251;
  return bytes;
}
