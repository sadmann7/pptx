import { describe, expect, it } from "vitest";

import { parseLayout } from "../../model/layout";
import { parseMaster } from "../../model/master";
import type { BaseNodeData } from "../../model/nodes/base-node";
import { buildPresentation, resolveNodePlaceholderInheritance } from "../../model/presentation";
import { parseTheme } from "../../model/theme";
import { parseXml } from "../../ooxml/xml-parser";
import { parseZip } from "../../ooxml/zip-parser";
import { renderSlide } from "../../renderer/slide-renderer";
import { buildCustomPptx, CustomPptxOptions } from "../helpers/fixture-extras";

const A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const P_NS =
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

function textShape(rPrChildrenXml: string, text: string): string {
  return `<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="4572000" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US">${rPrChildrenXml}</a:rPr><a:t>${text}</a:t></a:r></a:p></p:txBody>
</p:sp>`;
}

async function renderCustomSlide(options: CustomPptxOptions): Promise<HTMLElement> {
  const buffer = await buildCustomPptx(options);
  const files = await parseZip(buffer);
  const presentation = buildPresentation(files);
  return renderSlide(presentation, presentation.slides[0]).element;
}

function markerSpan(element: HTMLElement, marker: string): HTMLElement {
  const span = [...element.querySelectorAll("span")].find((s) => s.textContent?.includes(marker));
  if (!span) throw new Error(`span with "${marker}" not found`);
  return span;
}

describe("master clrMap remapping through the pipeline", () => {
  it("resolves bg1 through a swapped master clrMap for the slide background", async () => {
    const element = await renderCustomSlide({
      // Master background asks for scheme color bg1...
      masterBgXml: `<p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`,
      // ...and the clrMap sends bg1 to dk2 (theme: 44546A) instead of lt1.
      masterClrMapAttrs:
        'bg1="dk2" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"',
    });
    expect(element.style.backgroundColor.toUpperCase()).toBe("#44546A");
  });

  it("uses the standard bg1→lt1 mapping by default", async () => {
    const element = await renderCustomSlide({
      masterBgXml: `<p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`,
    });
    expect(element.style.backgroundColor.toUpperCase()).toBe("#FFFFFF");
  });
});

describe("layout clrMapOvr through the pipeline", () => {
  const shape = textShape(`<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>`, "OVERRIDE MARKER");

  it("overrideClrMapping in the layout swaps text scheme colors", async () => {
    const element = await renderCustomSlide({
      slides: [shape],
      layoutClrMapOvrXml: `<p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>`,
    });
    // tx1 → lt1 → window/FFFFFF instead of the usual dk1/000000.
    expect(markerSpan(element, "OVERRIDE MARKER").style.color.toUpperCase()).toBe("#FFFFFF");
  });

  it("masterClrMapping keeps the master's tx1→dk1 mapping", async () => {
    const element = await renderCustomSlide({ slides: [shape] });
    expect(markerSpan(element, "OVERRIDE MARKER").style.color.toUpperCase()).toBe("#000000");
  });
});

describe("theme font references through the pipeline", () => {
  it("resolves +mj-lt to the theme majorFont latin typeface", async () => {
    const element = await renderCustomSlide({
      slides: [textShape(`<a:latin typeface="+mj-lt"/>`, "MAJOR MARKER")],
      majorLatin: "Major Test Font",
    });
    expect(markerSpan(element, "MAJOR MARKER").style.fontFamily).toContain("Major Test Font");
  });

  it("resolves +mn-lt to the theme minorFont latin typeface", async () => {
    const element = await renderCustomSlide({
      slides: [textShape(`<a:latin typeface="+mn-lt"/>`, "MINOR MARKER")],
      minorLatin: "Minor Test Font",
    });
    expect(markerSpan(element, "MINOR MARKER").style.fontFamily).toContain("Minor Test Font");
  });
});

// ---------------------------------------------------------------------------
// Direct parser unit tests for uncovered branches
// ---------------------------------------------------------------------------

describe("parseTheme", () => {
  it("parses a clrScheme when the themeElements wrapper is missing", () => {
    const root = parseXml(`<a:theme ${A_NS}>
<a:clrScheme name="Bare"><a:dk1><a:srgbClr val="111111"/></a:dk1><a:accent1><a:srgbClr val="ABCDEF"/></a:accent1></a:clrScheme>
<a:fontScheme name="Bare"><a:majorFont><a:latin typeface="Zed"/></a:majorFont><a:minorFont><a:latin typeface="Yed"/></a:minorFont></a:fontScheme>
</a:theme>`);
    const theme = parseTheme(root);
    expect(theme.colorScheme.get("dk1")).toBe("111111");
    expect(theme.colorScheme.get("accent1")).toBe("ABCDEF");
    expect(theme.majorFont.latin).toBe("Zed");
    expect(theme.minorFont.latin).toBe("Yed");
  });

  it("falls back from sysClr lastClr to val", () => {
    const root = parseXml(`<a:theme ${A_NS}><a:themeElements>
<a:clrScheme name="T"><a:dk1><a:sysClr val="windowText"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FEFEFE"/></a:lt1></a:clrScheme>
<a:fontScheme name="T"><a:majorFont/><a:minorFont/></a:fontScheme>
</a:themeElements></a:theme>`);
    const theme = parseTheme(root);
    expect(theme.colorScheme.get("dk1")).toBe("windowText");
    expect(theme.colorScheme.get("lt1")).toBe("FEFEFE");
  });

  it("preserves script-specific fonts for east-asian resolution", () => {
    const root = parseXml(`<a:theme ${A_NS}><a:themeElements>
<a:clrScheme name="T"/>
<a:fontScheme name="T">
<a:majorFont><a:latin typeface="Latin"/><a:ea typeface=""/><a:font script="Hans" typeface="SimHei"/><a:font script="Jpan" typeface="Meiryo"/></a:majorFont>
<a:minorFont><a:latin typeface="Latin"/></a:minorFont>
</a:fontScheme>
</a:themeElements></a:theme>`);
    const theme = parseTheme(root);
    expect(theme.majorFont.scripts).toEqual({ Hans: "SimHei", Jpan: "Meiryo" });
    expect(theme.minorFont.scripts).toBeUndefined();
  });
});

describe("parseLayout", () => {
  it("extracts colorMapOverride only from overrideClrMapping", () => {
    const withOverride = parseLayout(
      parseXml(`<p:sldLayout ${P_NS}>
<p:cSld><p:spTree/></p:cSld>
<p:clrMapOvr><a:overrideClrMapping tx1="lt1" bg1="dk1"/></p:clrMapOvr>
</p:sldLayout>`),
    );
    expect(withOverride.colorMapOverride?.get("tx1")).toBe("lt1");
    expect(withOverride.colorMapOverride?.get("bg1")).toBe("dk1");

    const masterMapping = parseLayout(
      parseXml(`<p:sldLayout ${P_NS}>
<p:cSld><p:spTree/></p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`),
    );
    expect(masterMapping.colorMapOverride).toBeUndefined();
  });

  it("parses showMasterSp with a default of true", () => {
    const hidden = parseLayout(
      parseXml(`<p:sldLayout ${P_NS} showMasterSp="0"><p:cSld><p:spTree/></p:cSld></p:sldLayout>`),
    );
    expect(hidden.showMasterSp).toBe(false);

    const shown = parseLayout(
      parseXml(`<p:sldLayout ${P_NS}><p:cSld><p:spTree/></p:cSld></p:sldLayout>`),
    );
    expect(shown.showMasterSp).toBe(true);
  });
});

describe("parseMaster", () => {
  it("parses every clrMap attribute into the color map", () => {
    const master = parseMaster(
      parseXml(`<p:sldMaster ${P_NS}>
<p:cSld><p:spTree/></p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" hlink="hlink" folHlink="folHlink"/>
</p:sldMaster>`),
    );
    expect(master.colorMap.get("bg1")).toBe("lt1");
    expect(master.colorMap.get("tx1")).toBe("dk1");
    expect(master.colorMap.get("folHlink")).toBe("folHlink");
    expect(master.colorMap.size).toBe(7);
  });

  it("collects placeholders nested inside groups with absolute transforms", () => {
    // Group at (914400, 914400) EMU scaling children 2x; placeholder child at
    // (0,0) size 914400^2 in child space → absolute 96,96 px offset, 192px size.
    const master = parseMaster(
      parseXml(`<p:sldMaster ${P_NS}>
<p:cSld><p:spTree>
<p:grpSp>
<p:grpSpPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="1828800"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>
<p:sp>
<p:nvSpPr><p:cNvPr id="5" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr>
</p:sp>
</p:grpSp>
</p:spTree></p:cSld>
<p:clrMap bg1="lt1" tx1="dk1"/>
</p:sldMaster>`),
    );
    expect(master.placeholderEntries).toHaveLength(1);
    const entry = master.placeholderEntries![0];
    expect(entry.absoluteXfrm?.position).toEqual({ x: 96, y: 96 });
    expect(entry.absoluteXfrm?.size).toEqual({ w: 192, h: 192 });
  });
});

describe("placeholder inheritance type equivalence", () => {
  // Layout with a `title` placeholder and a `body` idx=1 placeholder,
  // mirroring decks where the slide uses ctrTitle/subTitle instead.
  const layout = parseLayout(
    parseXml(`<p:sldLayout ${P_NS}><p:cSld><p:spTree>
<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm></p:spPr>
</p:sp>
<p:sp>
<p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="914400" y="2743200"/><a:ext cx="1828800" cy="1828800"/></a:xfrm></p:spPr>
</p:sp>
</p:spTree></p:cSld></p:sldLayout>`),
  );

  function makeNode(type: string, idx?: number): BaseNodeData {
    return {
      id: "1",
      name: "n",
      nodeType: "shape",
      position: { x: 0, y: 0 },
      size: { w: 0, h: 0 },
      rotation: 0,
      flipH: false,
      flipV: false,
      placeholder: { type, idx },
      source: parseXml(`<p:sp ${P_NS}/>`),
    };
  }

  it("inherits layout `title` geometry for a slide `ctrTitle` placeholder", () => {
    const node = makeNode("ctrTitle");
    resolveNodePlaceholderInheritance(node, layout, undefined);
    expect(node.position).toEqual({ x: 96, y: 96 });
    expect(node.size).toEqual({ w: 192, h: 96 });
  });

  it("inherits layout `body` geometry for a slide `subTitle` placeholder with matching idx", () => {
    const node = makeNode("subTitle", 1);
    resolveNodePlaceholderInheritance(node, layout, undefined);
    expect(node.position).toEqual({ x: 96, y: 288 });
    expect(node.size).toEqual({ w: 192, h: 192 });
  });

  it("still prefers exact type+idx matches", () => {
    const node = makeNode("body", 1);
    resolveNodePlaceholderInheritance(node, layout, undefined);
    expect(node.position).toEqual({ x: 96, y: 288 });
    expect(node.size).toEqual({ w: 192, h: 192 });
  });

  it("inherits layout `title` lstStyle for a rendered `ctrTitle` placeholder", async () => {
    // Layout title placeholder styles runs at 44pt in a distinctive color.
    const layoutTitle = `<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="4572000" cy="914400"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/>
<a:lstStyle><a:lvl1pPr><a:defRPr sz="4400"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>
<a:p><a:r><a:t>Title prompt</a:t></a:r></a:p></p:txBody>
</p:sp>`;
    // Slide uses ctrTitle with no geometry, no run properties of its own.
    const slideCtrTitle = `<p:sp>
<p:nvSpPr><p:cNvPr id="10" name="Title"/><p:cNvSpPr txBox="1"/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr>
<p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>EQUIV MARKER</a:t></a:r></a:p></p:txBody>
</p:sp>`;

    const element = await renderCustomSlide({
      slides: [slideCtrTitle],
      layoutShapesXml: layoutTitle,
    });

    const span = markerSpan(element, "EQUIV MARKER");
    // Color and size come from the layout title's lstStyle via type equivalence.
    expect(span.style.color.toUpperCase()).toBe("#123456");
    expect(span.style.fontSize).toBe("44pt");
  });
});
