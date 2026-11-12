import type { MediaResolver } from "../media/resolve";
import type { PptxPackage } from "../ooxml/package";
import { parseRels, RelEntry, resolveRelTarget } from "../ooxml/rel";
import { emuToPx } from "../ooxml/unit";
import { parseXml, SafeXmlNode } from "../ooxml/xml";
import { PptxFiles } from "../ooxml/zip";
import { LayoutData, parseLayout, PlaceholderEntry } from "./layout";
import { MasterData, parseMaster } from "./master";
import { BaseNodeData, NodePosition, NodeSize, PlaceholderInfo } from "./nodes/base";
import type { GroupNodeData } from "./nodes/group";
import { createLazySlide, materializeSlideData, parseSlide, SlideData, SlideNode } from "./slide";
import { parseTheme, ThemeData } from "./theme";

export interface EmbeddedFontVariant {
  /** The font part's path inside the zip, like "ppt/fonts/font1.fntdata". */
  path: string;
  /** The GUID that XOR-deobfuscates an ODTTF font's first 32 bytes, absent for plain EOT. */
  fontKey?: string;
}

export interface EmbeddedFontEntry {
  typeface: string;
  pitchFamily?: number;
  charset?: string;
  regular?: EmbeddedFontVariant;
  bold?: EmbeddedFontVariant;
  italic?: EmbeddedFontVariant;
  boldItalic?: EmbeddedFontVariant;
}

export interface PresentationData {
  width: number;
  height: number;
  slides: SlideData[];
  layouts: Map<string, LayoutData>;
  masters: Map<string, MasterData>;
  themes: Map<string, ThemeData>;
  slideToLayout: Map<number, string>;
  layoutToMaster: Map<string, string>;
  masterToTheme: Map<string, string>;
  media: Map<string, Uint8Array>;
  mediaResolver?: MediaResolver;
  tableStyles?: SafeXmlNode;
  /** Presentation-wide default text style from ppt/presentation.xml. */
  defaultTextStyle?: SafeXmlNode;
  charts: Map<string, SafeXmlNode>;
  /** SmartArt fallback drawing parts keyed by part path (ppt/diagrams/drawing*.xml). */
  diagramDrawings?: Map<string, string>;
  /** Chart-local theme overrides keyed by chart part path. */
  chartThemes?: Map<string, ThemeData>;
  /** Chart style parts keyed by chart part path. */
  chartStyles?: Map<string, SafeXmlNode>;
  /** Chart color style parts keyed by chart part path. */
  chartColorStyles?: Map<string, SafeXmlNode>;
  isWps: boolean;
  /** Embedded font binary data keyed by zip path (ppt/fonts/*). */
  fonts: Map<string, Uint8Array>;
  /** The embedded font list from presentation.xml. */
  embeddedFonts?: EmbeddedFontEntry[];
  /**
   * The source package kept for round-trip saves with `writePptx()`, present
   * when parsed with `keepSourcePackage: true`.
   */
  sourcePackage?: PptxPackage;
}

export interface BuildPresentationOptions {
  /**
   * Whether to defer parsing a slide's nodes until it's rendered, searched, serialized, or
   * materialized.
   *
   * @default false
   */
  lazySlides?: boolean;
}

function basePath(filePath: string): string {
  const idx = filePath.lastIndexOf("/");
  return idx >= 0 ? filePath.substring(0, idx) : "";
}

function relsPathFor(filePath: string): string {
  const dir = basePath(filePath);
  const fileName = filePath.substring(filePath.lastIndexOf("/") + 1);
  return `${dir}/_rels/${fileName}.rels`;
}

function detectWps(presentationXml: string): boolean {
  // WPS adds its own namespace or processing instructions.
  return (
    presentationXml.includes("wps") ||
    presentationXml.includes("kso") ||
    presentationXml.includes("Kingsoft") ||
    presentationXml.includes("WPS")
  );
}

function findRelByType(rels: Map<string, RelEntry>, typeSubstring: string): RelEntry | undefined {
  for (const [, entry] of rels) {
    if (entry.type.includes(typeSubstring)) {
      return entry;
    }
  }
  return undefined;
}

function findRelsByType(rels: Map<string, RelEntry>, typeSubstring: string): [string, RelEntry][] {
  const results: [string, RelEntry][] = [];
  for (const [rId, entry] of rels) {
    if (entry.type.includes(typeSubstring)) {
      results.push([rId, entry]);
    }
  }
  return results;
}

/** Builds presentation data from extracted PPTX files, resolving each slide's layout, master, and theme. */
export function buildPresentation(
  files: PptxFiles,
  options: BuildPresentationOptions = {},
): PresentationData {
  const sourcePackage = files.sourcePackage;

  const presRoot = parseXml(files.presentation);
  sourcePackage?.registerXmlRoot("ppt/presentation.xml", presRoot);
  const presRels = parseRels(files.presentationRels);

  const sldSz = presRoot.child("sldSz");
  const width = emuToPx(sldSz.numAttr("cx") ?? 9144000); // 10 inches.
  const height = emuToPx(sldSz.numAttr("cy") ?? 6858000); // 7.5 inches.

  const isWps = detectWps(files.presentation);

  const defaultTextStyle = presRoot.child("defaultTextStyle");

  const themes = new Map<string, ThemeData>();
  for (const [themePath, themeXml] of files.themes) {
    const themeRoot = parseXml(themeXml);
    sourcePackage?.registerXmlRoot(themePath, themeRoot);
    themes.set(themePath, parseTheme(themeRoot));
  }

  const masters = new Map<string, MasterData>();
  const masterToTheme = new Map<string, string>();

  for (const [masterPath, masterXml] of files.slideMasters) {
    const masterRoot = parseXml(masterXml);
    sourcePackage?.registerXmlRoot(masterPath, masterRoot);
    const masterData = parseMaster(masterRoot);

    const masterRelsPath = relsPathFor(masterPath);
    const masterRelsXml = files.slideMasterRels.get(masterRelsPath);
    if (masterRelsXml) {
      const masterRels = parseRels(masterRelsXml);
      masterData.rels = masterRels;
      const themeRel = findRelByType(masterRels, "theme");
      if (themeRel) {
        const themePath = resolveRelTarget(basePath(masterPath), themeRel.target);
        masterToTheme.set(masterPath, themePath);
      }
    }
    masters.set(masterPath, masterData);
  }

  const layouts = new Map<string, LayoutData>();
  const layoutToMaster = new Map<string, string>();

  for (const [layoutPath, layoutXml] of files.slideLayouts) {
    const layoutRoot = parseXml(layoutXml);
    sourcePackage?.registerXmlRoot(layoutPath, layoutRoot);
    const layoutData = parseLayout(layoutRoot);

    const layoutRelsPath = relsPathFor(layoutPath);
    const layoutRelsXml = files.slideLayoutRels.get(layoutRelsPath);
    if (layoutRelsXml) {
      const layoutRels = parseRels(layoutRelsXml);
      layoutData.rels = layoutRels;
      const masterRel = findRelByType(layoutRels, "slideMaster");
      if (masterRel) {
        const masterPath = resolveRelTarget(basePath(layoutPath), masterRel.target);
        layoutToMaster.set(layoutPath, masterPath);
      }
    }
    layouts.set(layoutPath, layoutData);
  }

  const charts = new Map<string, SafeXmlNode>();
  const chartThemes = new Map<string, ThemeData>();
  const chartStyles = new Map<string, SafeXmlNode>();
  const chartColorStyles = new Map<string, SafeXmlNode>();
  for (const [chartPath, chartXml] of files.charts) {
    const chartRoot = parseXml(chartXml);
    sourcePackage?.registerXmlRoot(chartPath, chartRoot);
    if (chartRoot.exists()) {
      charts.set(chartPath, chartRoot);
    }

    const chartRelsPath = relsPathFor(chartPath);
    const chartRelsXml = files.chartRels?.get(chartRelsPath);
    if (!chartRelsXml) continue;
    const chartRels = parseRels(chartRelsXml);

    const chartStyleRel = findRelByType(chartRels, "chartStyle");
    if (chartStyleRel) {
      const chartStylePath = resolveRelTarget(basePath(chartPath), chartStyleRel.target);
      const chartStyleXml = files.chartStyles?.get(chartStylePath);
      if (chartStyleXml) {
        const chartStyleRoot = parseXml(chartStyleXml);
        sourcePackage?.registerXmlRoot(chartStylePath, chartStyleRoot);
        if (chartStyleRoot.exists()) chartStyles.set(chartPath, chartStyleRoot);
      }
    }

    const chartColorStyleRel = findRelByType(chartRels, "chartColorStyle");
    if (chartColorStyleRel) {
      const chartColorStylePath = resolveRelTarget(basePath(chartPath), chartColorStyleRel.target);
      const chartColorStyleXml = files.chartColors?.get(chartColorStylePath);
      if (chartColorStyleXml) {
        const chartColorStyleRoot = parseXml(chartColorStyleXml);
        sourcePackage?.registerXmlRoot(chartColorStylePath, chartColorStyleRoot);
        if (chartColorStyleRoot.exists()) chartColorStyles.set(chartPath, chartColorStyleRoot);
      }
    }

    const themeOverrideRel = findRelByType(chartRels, "themeOverride");
    if (!themeOverrideRel) continue;
    const themeOverridePath = resolveRelTarget(basePath(chartPath), themeOverrideRel.target);
    const themeOverrideXml =
      files.themeOverrides?.get(themeOverridePath) ?? files.themes.get(themeOverridePath);
    if (!themeOverrideXml) continue;
    const themeOverrideRoot = parseXml(themeOverrideXml);
    sourcePackage?.registerXmlRoot(themeOverridePath, themeOverrideRoot);
    if (themeOverrideRoot.exists()) {
      chartThemes.set(chartPath, parseTheme(themeOverrideRoot));
    }
  }

  // sldId references presentation.xml.rels through `r:id`, which
  // SafeXmlNode.attr() reads by its local name.
  const sldIdLst = presRoot.child("sldIdLst");
  const orderedSlideTargets: string[] = [];

  for (const sldId of sldIdLst.children("sldId")) {
    const rId = sldId.attr("r:id") ?? sldId.attr("id");
    if (rId) {
      const relEntry = presRels.get(rId);
      if (relEntry) {
        const slidePath = resolveRelTarget("ppt", relEntry.target);
        orderedSlideTargets.push(slidePath);
      }
    }
  }

  // Without a usable sldIdLst, slides come straight from presentation.xml.rels.
  if (orderedSlideTargets.length === 0) {
    const slideRels = findRelsByType(presRels, "slide");
    // rId number stands in for slide order.
    slideRels.sort((a, b) => {
      const numA = parseInt(a[0].replace(/\D/g, ""), 10) || 0;
      const numB = parseInt(b[0].replace(/\D/g, ""), 10) || 0;
      return numA - numB;
    });
    for (const [, entry] of slideRels) {
      // `/slide` also matches the slideLayout and slideMaster types.
      if (
        entry.type.includes("/slide") &&
        !entry.type.includes("slideLayout") &&
        !entry.type.includes("slideMaster")
      ) {
        const slidePath = resolveRelTarget("ppt", entry.target);
        orderedSlideTargets.push(slidePath);
      }
    }
  }

  const slides: SlideData[] = [];
  const slideToLayout = new Map<number, string>();

  for (let i = 0; i < orderedSlideTargets.length; i++) {
    const slidePath = orderedSlideTargets[i];
    const slideXml = files.slides.get(slidePath);
    if (!slideXml) continue;

    const slideRelsPath = relsPathFor(slidePath);
    const slideRelsXml = files.slideRels.get(slideRelsPath);
    const slideRels = slideRelsXml ? parseRels(slideRelsXml) : new Map<string, RelEntry>();

    let slideData: SlideData;
    if (options.lazySlides) {
      slideData = createLazySlide(slideXml, i, slideRels, slidePath);
    } else {
      const slideRoot = parseXml(slideXml);
      sourcePackage?.registerXmlRoot(slidePath, slideRoot);
      slideData = parseSlide(slideRoot, i, slideRels, slidePath, files.diagramDrawings);
    }

    if (slideData.layoutIndex) {
      const layoutPath = resolveRelTarget(basePath(slidePath), slideData.layoutIndex);
      slideData.layoutIndex = layoutPath;
      slideToLayout.set(i, layoutPath);
    }

    slides.push(slideData);
  }

  let tableStyles: SafeXmlNode | undefined;
  if (files.tableStyles) {
    const tsRoot = parseXml(files.tableStyles);
    sourcePackage?.registerXmlRoot("ppt/tableStyles.xml", tsRoot);
    if (tsRoot.exists()) {
      tableStyles = tsRoot;
    }
  }

  const embeddedFonts = parseEmbeddedFontList(presRoot, presRels);

  const result: PresentationData = {
    width,
    height,
    slides,
    layouts,
    masters,
    themes,
    slideToLayout,
    layoutToMaster,
    masterToTheme,
    media: files.media,
    mediaResolver: files.mediaResolver,
    tableStyles,
    defaultTextStyle: defaultTextStyle.exists() ? defaultTextStyle : undefined,
    charts,
    diagramDrawings: files.diagramDrawings,
    chartThemes,
    chartStyles,
    chartColorStyles,
    isWps,
    fonts: files.fonts,
    embeddedFonts: embeddedFonts.length > 0 ? embeddedFonts : undefined,
    sourcePackage,
  };

  if (!options.lazySlides) {
    resolvePlaceholderInheritance(result);
  }

  return result;
}

function resolveEmbeddedFontVariant(
  variantNode: SafeXmlNode,
  presRels: Map<string, RelEntry>,
): EmbeddedFontVariant | undefined {
  if (!variantNode.exists()) return undefined;
  const rId = variantNode.attr("r:id") ?? variantNode.attr("id");
  if (!rId) return undefined;
  const rel = presRels.get(rId);
  if (!rel) return undefined;
  const fontPath = resolveRelTarget("ppt", rel.target);
  const fontKey = variantNode.attr("fontKey") ?? undefined;
  return { path: fontPath, fontKey };
}

function parseEmbeddedFontList(
  presRoot: SafeXmlNode,
  presRels: Map<string, RelEntry>,
): EmbeddedFontEntry[] {
  const fontLst = presRoot.child("embeddedFontLst");
  if (!fontLst.exists()) return [];
  const entries: EmbeddedFontEntry[] = [];
  for (const embFont of fontLst.children("embeddedFont")) {
    const fontNode = embFont.child("font");
    const typeface = fontNode.attr("typeface");
    if (!typeface) continue;
    const pitchFamilyStr = fontNode.attr("pitchFamily");
    const entry: EmbeddedFontEntry = {
      typeface,
      pitchFamily: pitchFamilyStr !== undefined ? Number(pitchFamilyStr) : undefined,
      charset: fontNode.attr("charset") ?? undefined,
      regular: resolveEmbeddedFontVariant(embFont.child("regular"), presRels),
      bold: resolveEmbeddedFontVariant(embFont.child("bold"), presRels),
      italic: resolveEmbeddedFontVariant(embFont.child("italic"), presRels),
      boldItalic: resolveEmbeddedFontVariant(embFont.child("boldItalic"), presRels),
    };
    entries.push(entry);
  }
  return entries;
}

function getPhInfo(phNode: SafeXmlNode): { type?: string; idx?: number } {
  for (const wrapper of ["nvSpPr", "nvPicPr", "nvGrpSpPr", "nvGraphicFramePr", "nvCxnSpPr"]) {
    const nvWrapper = phNode.child(wrapper);
    if (nvWrapper.exists()) {
      const nvPr = nvWrapper.child("nvPr");
      const ph = nvPr.child("ph");
      if (ph.exists()) {
        const type = ph.attr("type");
        const idxStr = ph.attr("idx");
        const idx = idxStr !== undefined ? Number(idxStr) : undefined;
        return { type, idx: idx !== undefined && !isNaN(idx) ? idx : undefined };
      }
    }
  }
  return {};
}

function getPhXfrm(phNode: SafeXmlNode): { position: NodePosition; size: NodeSize } | undefined {
  // spPr > xfrm on most shapes, a direct xfrm on graphic frames.
  const spPrXfrm = phNode.child("spPr").child("xfrm");
  const xfrm = spPrXfrm.exists() ? spPrXfrm : phNode.child("xfrm");
  if (xfrm.exists()) {
    const off = xfrm.child("off");
    const ext = xfrm.child("ext");
    const x = off.numAttr("x");
    const cx = ext.numAttr("cx");
    if (x !== undefined && cx !== undefined) {
      return {
        position: { x: emuToPx(off.numAttr("x") ?? 0), y: emuToPx(off.numAttr("y") ?? 0) },
        size: { w: emuToPx(ext.numAttr("cx") ?? 0), h: emuToPx(ext.numAttr("cy") ?? 0) },
      };
    }
  }
  return undefined;
}

// PowerPoint matches a slide placeholder to a layout or master one of an
// equivalent type, like `ctrTitle` to `title` or `subTitle` to `body`.
const PLACEHOLDER_TYPE_EQUIVALENCE: Record<string, string> = {
  title: "title",
  ctrTitle: "title",
  body: "body",
  subTitle: "body",
  obj: "body",
};

export function normalizePlaceholderType(type: string | undefined): string | undefined {
  if (type === undefined) return undefined;
  return PLACEHOLDER_TYPE_EQUIVALENCE[type] ?? type;
}

// Best match first: exact type and idx, equivalent type and idx, exact type,
// equivalent type, idx with neither typed, then idx alone when both set it
// explicitly. A missing idx is 0, the OOXML default.
function findMatchingLayoutPlaceholder(
  placeholders: PlaceholderEntry[],
  type?: string,
  idx?: number,
): PlaceholderEntry | undefined {
  const typeKey = normalizePlaceholderType(type);
  const normIdx = idx ?? 0;

  let best: PlaceholderEntry | undefined;
  let bestScore = 0;

  for (const entry of placeholders) {
    const info = getPhInfo(entry.node);
    const sameType = type !== undefined && info.type === type;
    const equivType = typeKey !== undefined && normalizePlaceholderType(info.type) === typeKey;
    const sameIdx = (info.idx ?? 0) === normIdx;

    let score = 0;
    if (sameType && sameIdx) score = 6;
    else if (equivType && sameIdx) score = 5;
    else if (sameType) score = 4;
    else if (equivType) score = 3;
    else if (sameIdx && type === undefined && info.type === undefined) score = 2;
    else if (sameIdx && idx !== undefined && info.idx !== undefined) score = 1;

    if (score > bestScore) {
      best = entry;
      bestScore = score;
      if (score === 6) break;
    }
  }

  return best;
}

function getMasterPlaceholderEntries(master: MasterData): PlaceholderEntry[] {
  return master.placeholderEntries ?? master.placeholders.map((node) => ({ node }));
}

// The template node's txBody carries the "Click to add title" prompt runs edit mode shows.
export function findPlaceholderTemplateNode(
  placeholder: PlaceholderInfo,
  layout: LayoutData | undefined,
  master: MasterData | undefined,
): SafeXmlNode | undefined {
  const layoutMatch = layout
    ? findMatchingLayoutPlaceholder(layout.placeholders, placeholder.type, placeholder.idx)
    : undefined;
  if (layoutMatch) return layoutMatch.node;
  const masterMatch = master
    ? findMatchingLayoutPlaceholder(
        getMasterPlaceholderEntries(master),
        placeholder.type,
        placeholder.idx,
      )
    : undefined;
  return masterMatch?.node;
}

function resolvePlaceholderInheritance(pres: PresentationData): void {
  for (let i = 0; i < pres.slides.length; i++) {
    resolveSlidePlaceholderInheritance(pres, pres.slides[i]);
  }
}

function resolveSlidePlaceholderInheritance(pres: PresentationData, slide: SlideData): void {
  if (slide.placeholderInheritanceResolved) return;
  resolveSlideNodePlaceholders(pres, slide, slide.nodes);
  slide.placeholderInheritanceResolved = true;
}

export function resolveSlideNodePlaceholders(
  pres: PresentationData,
  slide: SlideData,
  nodes: SlideNode[],
): void {
  const layoutPath = pres.slideToLayout.get(slide.index) || slide.layoutIndex;
  const layout = layoutPath ? pres.layouts.get(layoutPath) : undefined;
  const masterPath = layoutPath ? pres.layoutToMaster.get(layoutPath) : undefined;
  const master = masterPath ? pres.masters.get(masterPath) : undefined;
  resolveNodesPlaceholders(nodes, layout, master);
}

export function materializeSlide(pres: PresentationData, slide: SlideData): void {
  const slideRoot = materializeSlideData(slide, pres.diagramDrawings);
  if (slideRoot) {
    pres.sourcePackage?.registerXmlRoot(slide.id, slideRoot);
  }
  resolveSlidePlaceholderInheritance(pres, slide);
}

export function materializeAllSlides(pres: PresentationData): void {
  for (const slide of pres.slides) {
    materializeSlide(pres, slide);
  }
}

function getPhBodyPr(phNode: SafeXmlNode): SafeXmlNode | undefined {
  const txBody = phNode.child("txBody");
  if (!txBody.exists()) return undefined;
  const bodyPr = txBody.child("bodyPr");
  return bodyPr.exists() ? bodyPr : undefined;
}

function inheritPlaceholderType(target: PlaceholderInfo, sourceNode: SafeXmlNode): void {
  if (target.type) return;

  const source = getPhInfo(sourceNode);
  if (source.type) {
    target.type = source.type;
  }
}

interface PlaceholderInheritanceOptions {
  // Placeholders are in slide space but group children render in the group's
  // child space, so an inherited xfrm is inverted through this group first.
  parentGroup?: GroupNodeData;
}

function toGroupChildXfrm(
  xfrm: { position: NodePosition; size: NodeSize },
  group: GroupNodeData,
): { position: NodePosition; size: NodeSize } {
  const scaleX = group.childExtent.w > 0 ? group.size.w / group.childExtent.w : 1;
  const scaleY = group.childExtent.h > 0 ? group.size.h / group.childExtent.h : 1;
  if (scaleX === 0 || scaleY === 0) return xfrm;

  return {
    position: {
      x: group.childOffset.x + (xfrm.position.x - group.position.x) / scaleX,
      y: group.childOffset.y + (xfrm.position.y - group.position.y) / scaleY,
    },
    size: {
      w: xfrm.size.w / scaleX,
      h: xfrm.size.h / scaleY,
    },
  };
}

function resolveInheritedXfrm(
  xfrm: { position: NodePosition; size: NodeSize },
  options: PlaceholderInheritanceOptions,
): { position: NodePosition; size: NodeSize } {
  return options.parentGroup ? toGroupChildXfrm(xfrm, options.parentGroup) : xfrm;
}

export function resolveNodePlaceholderInheritance(
  node: SlideNode | BaseNodeData,
  layout: LayoutData | undefined,
  master: MasterData | undefined,
  options: PlaceholderInheritanceOptions = {},
): void {
  if (!node.placeholder) return;

  const { type, idx } = node.placeholder;
  const findMasterMatch = (): PlaceholderEntry | undefined =>
    master
      ? findMatchingLayoutPlaceholder(
          getMasterPlaceholderEntries(master),
          node.placeholder?.type ?? type,
          idx,
        )
      : undefined;
  const sizeIsEmpty = node.size.w === 0 && node.size.h === 0;
  const positionLooksDefault = node.position.y < 5; // Near the top counts as unset.

  if (layout) {
    const layoutMatch = findMatchingLayoutPlaceholder(layout.placeholders, type, idx);
    if (layoutMatch) {
      inheritPlaceholderType(node.placeholder, layoutMatch.node);

      const rawXfrm = layoutMatch.absoluteXfrm ?? getPhXfrm(layoutMatch.node);
      if (rawXfrm) {
        const xfrm = resolveInheritedXfrm(rawXfrm, options);
        if (sizeIsEmpty) {
          node.position = xfrm.position;
          node.size = xfrm.size;
        } else if (positionLooksDefault) {
          node.position = xfrm.position;
        }
      }

      // A placeholder's bodyPr (anchor, insets) comes from its layout placeholder.
      if ("textBody" in node && node.textBody) {
        const layoutBodyPr = getPhBodyPr(layoutMatch.node);
        if (layoutBodyPr) {
          node.textBody.layoutBodyProperties = layoutBodyPr;
        }
      }

      if (rawXfrm) {
        const masterMatch = findMasterMatch();
        if (masterMatch) {
          inheritPlaceholderType(node.placeholder, masterMatch.node);
          if ("textBody" in node && node.textBody && !node.textBody.layoutBodyProperties) {
            const masterBodyPr = getPhBodyPr(masterMatch.node);
            if (masterBodyPr) {
              node.textBody.layoutBodyProperties = masterBodyPr;
            }
          }
        }
        return;
      }
    }
  }

  const masterMatch = findMasterMatch();
  if (masterMatch) {
    inheritPlaceholderType(node.placeholder, masterMatch.node);

    const rawXfrm = masterMatch.absoluteXfrm ?? getPhXfrm(masterMatch.node);
    if (rawXfrm) {
      const xfrm = resolveInheritedXfrm(rawXfrm, options);
      if (sizeIsEmpty) {
        node.position = xfrm.position;
        node.size = xfrm.size;
      } else if (positionLooksDefault) {
        node.position = xfrm.position;
      }
    }

    if ("textBody" in node && node.textBody && !node.textBody.layoutBodyProperties) {
      const masterBodyPr = getPhBodyPr(masterMatch.node);
      if (masterBodyPr) {
        node.textBody.layoutBodyProperties = masterBodyPr;
      }
    }
  }
}

function resolveNodesPlaceholders(
  nodes: SlideNode[],
  layout: LayoutData | undefined,
  master: MasterData | undefined,
): void {
  for (const node of nodes) {
    if (node.nodeType === "group" && "children" in node) {
      // Group children stay raw XML until the renderer parses them.
    }

    resolveNodePlaceholderInheritance(node, layout, master);
  }
}
