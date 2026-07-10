/**
 * Assembles parsed PPTX components (themes, masters, layouts, slides)
 * into a unified {@link PresentationData} structure.
 */

import type { MediaResolver } from "../media/resolve";
import type { PptxPackage } from "../ooxml/package";
import { parseRels, RelEntry, resolveRelTarget } from "../ooxml/rel";
import { emuToPx } from "../ooxml/unit";
import { parseXml, SafeXmlNode } from "../ooxml/xml";
import { PptxFiles } from "../ooxml/zip";
import { LayoutData, parseLayout, PlaceholderEntry } from "./layout";
import { MasterData, parseMaster } from "./master";
import { BaseNodeData, PlaceholderInfo, Position, Size } from "./nodes/base";
import type { GroupNodeData } from "./nodes/group";
import { createLazySlide, materializeSlideData, parseSlide, SlideData, SlideNode } from "./slide";
import { parseTheme, ThemeData } from "./theme";

export interface EmbeddedFontVariant {
  /** Resolved path inside the zip (e.g. "ppt/fonts/font1.fntdata"). */
  path: string;
  /** GUID used to XOR-deobfuscate the first 32 bytes (ODTTF). Absent for plain EOT. */
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
  /** Embedded font binary data keyed by zip path (ppt/fonts/*.fntdata). */
  fonts: Map<string, Uint8Array>;
  /** Parsed embedded font list from presentation.xml. */
  embeddedFonts?: EmbeddedFontEntry[];
  /**
   * Retained source package for round-trip save. Present when the zip was
   * parsed with `keepPackage: true`. Use `writePptx()` to produce a .pptx.
   */
  /** Retained source PPTX package for round-trip save. Present when parsed with `keepPackage: true`. Use `writePptx()` to produce a .pptx. */
  sourcePackage?: PptxPackage;
}

export interface BuildPresentationOptions {
  /**
   * When `true`, defers per-slide shape/table/chart node parsing until a
   * slide is rendered, searched, serialized, or explicitly materialized.
   *
   * @default false
   */
  lazy?: boolean;
}

/**
 * Derive the base directory from a file path.
 * E.g., "ppt/slides/slide1.xml" → "ppt/slides"
 */
function basePath(filePath: string): string {
  const idx = filePath.lastIndexOf("/");
  return idx >= 0 ? filePath.substring(0, idx) : "";
}

/**
 * For a given XML file path, find its corresponding .rels file path.
 * E.g., "ppt/slides/slide1.xml" → "ppt/slides/_rels/slide1.xml.rels"
 */
function relsPathFor(filePath: string): string {
  const dir = basePath(filePath);
  const fileName = filePath.substring(filePath.lastIndexOf("/") + 1);
  return `${dir}/_rels/${fileName}.rels`;
}

/**
 * Detect WPS (Kingsoft Office / WPS Office) by checking for known markers
 * in the presentation XML string.
 */
function detectWps(presentationXml: string): boolean {
  // WPS adds its own namespace or processing instructions
  return (
    presentationXml.includes("wps") ||
    presentationXml.includes("kso") ||
    presentationXml.includes("Kingsoft") ||
    presentationXml.includes("WPS")
  );
}

/**
 * Find a rels entry by type substring match.
 */
function findRelByType(rels: Map<string, RelEntry>, typeSubstring: string): RelEntry | undefined {
  for (const [, entry] of rels) {
    if (entry.type.includes(typeSubstring)) {
      return entry;
    }
  }
  return undefined;
}

/**
 * Find ALL rels entries matching a type substring, returning [rId, entry] pairs.
 */
function findRelsByType(rels: Map<string, RelEntry>, typeSubstring: string): [string, RelEntry][] {
  const results: [string, RelEntry][] = [];
  for (const [rId, entry] of rels) {
    if (entry.type.includes(typeSubstring)) {
      results.push([rId, entry]);
    }
  }
  return results;
}

/**
 * Build the complete PresentationData from extracted PPTX files.
 *
 * This is the main factory function that wires together all parsed components:
 * 1. Parses presentation.xml for slide ordering and size
 * 2. Resolves the full relationship chain: slide → layout → master → theme
 * 3. Parses each component and assembles the final structure
 */
export function buildPresentation(
  files: PptxFiles,
  options: BuildPresentationOptions = {},
): PresentationData {
  const sourcePackage = files.sourcePackage;

  // --- Parse presentation root ---
  const presRoot = parseXml(files.presentation);
  sourcePackage?.registerXmlRoot("ppt/presentation.xml", presRoot);
  const presRels = parseRels(files.presentationRels);

  // --- Slide size ---
  const sldSz = presRoot.child("sldSz");
  const width = emuToPx(sldSz.numAttr("cx") ?? 9144000); // default 10 inches
  const height = emuToPx(sldSz.numAttr("cy") ?? 6858000); // default 7.5 inches

  // --- WPS detection ---
  const isWps = detectWps(files.presentation);

  // --- Presentation default text style ---
  const defaultTextStyle = presRoot.child("defaultTextStyle");

  // --- Parse themes ---
  const themes = new Map<string, ThemeData>();
  for (const [themePath, themeXml] of files.themes) {
    const themeRoot = parseXml(themeXml);
    sourcePackage?.registerXmlRoot(themePath, themeRoot);
    themes.set(themePath, parseTheme(themeRoot));
  }

  // --- Parse slide masters and build master→theme mapping ---
  const masters = new Map<string, MasterData>();
  const masterToTheme = new Map<string, string>();

  for (const [masterPath, masterXml] of files.slideMasters) {
    const masterRoot = parseXml(masterXml);
    sourcePackage?.registerXmlRoot(masterPath, masterRoot);
    const masterData = parseMaster(masterRoot);

    // Find theme relationship for this master
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

  // --- Parse slide layouts and build layout→master mapping ---
  const layouts = new Map<string, LayoutData>();
  const layoutToMaster = new Map<string, string>();

  for (const [layoutPath, layoutXml] of files.slideLayouts) {
    const layoutRoot = parseXml(layoutXml);
    sourcePackage?.registerXmlRoot(layoutPath, layoutRoot);
    const layoutData = parseLayout(layoutRoot);

    // Find master relationship for this layout
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

  // --- Parse charts ---
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

  // --- Determine slide ordering ---
  // The sldIdLst contains sldId elements with r:id attributes that reference
  // presentation.xml.rels. We need to handle the fact that the attr might be
  // stored as 'r:id' in the original XML but SafeXmlNode.attr() uses localName.
  const sldIdLst = presRoot.child("sldIdLst");
  const orderedSlideTargets: string[] = [];

  for (const sldId of sldIdLst.children("sldId")) {
    // Try multiple attribute name patterns
    const rId = sldId.attr("r:id") ?? sldId.attr("id");
    if (rId) {
      const relEntry = presRels.get(rId);
      if (relEntry) {
        const slidePath = resolveRelTarget("ppt", relEntry.target);
        orderedSlideTargets.push(slidePath);
      }
    }
  }

  // Fallback: if sldIdLst parsing didn't yield results, use presRels directly
  if (orderedSlideTargets.length === 0) {
    const slideRels = findRelsByType(presRels, "slide");
    // Sort by rId number to maintain order
    slideRels.sort((a, b) => {
      const numA = parseInt(a[0].replace(/\D/g, ""), 10) || 0;
      const numB = parseInt(b[0].replace(/\D/g, ""), 10) || 0;
      return numA - numB;
    });
    for (const [, entry] of slideRels) {
      // Only include direct slide relationships, not slideLayout or slideMaster
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

  // --- Parse slides ---
  const slides: SlideData[] = [];
  const slideToLayout = new Map<number, string>();

  for (let i = 0; i < orderedSlideTargets.length; i++) {
    const slidePath = orderedSlideTargets[i];
    const slideXml = files.slides.get(slidePath);
    if (!slideXml) continue;

    // Parse slide rels
    const slideRelsPath = relsPathFor(slidePath);
    const slideRelsXml = files.slideRels.get(slideRelsPath);
    const slideRels = slideRelsXml ? parseRels(slideRelsXml) : new Map<string, RelEntry>();

    let slideData: SlideData;
    if (options.lazy) {
      slideData = createLazySlide(slideXml, i, slideRels, slidePath);
    } else {
      const slideRoot = parseXml(slideXml);
      sourcePackage?.registerXmlRoot(slidePath, slideRoot);
      slideData = parseSlide(slideRoot, i, slideRels, slidePath, files.diagramDrawings);
    }

    // Resolve layout path from the slide's layout relationship target
    if (slideData.layoutIndex) {
      const layoutPath = resolveRelTarget(basePath(slidePath), slideData.layoutIndex);
      slideData.layoutIndex = layoutPath;
      slideToLayout.set(i, layoutPath);
    }

    slides.push(slideData);
  }

  // --- Table styles ---
  let tableStyles: SafeXmlNode | undefined;
  if (files.tableStyles) {
    const tsRoot = parseXml(files.tableStyles);
    sourcePackage?.registerXmlRoot("ppt/tableStyles.xml", tsRoot);
    if (tsRoot.exists()) {
      tableStyles = tsRoot;
    }
  }

  // --- Parse embedded fonts ---
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

  // --- Resolve placeholder position inheritance ---
  if (!options.lazy) {
    resolvePlaceholderInheritance(result);
  }

  return result;
}

// ---------------------------------------------------------------------------
// Embedded Font Parsing
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// Placeholder Position Inheritance
// ---------------------------------------------------------------------------

/**
 * Extract placeholder info (type, idx) from a raw placeholder XML node
 * stored in layout/master.
 */
function getPhInfo(phNode: SafeXmlNode): { type?: string; idx?: number } {
  // Try nvSpPr > nvPr > ph, or nvPicPr > nvPr > ph
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

/**
 * Extract xfrm position/size from a raw placeholder XML node.
 */
function getPhXfrm(phNode: SafeXmlNode): { position: Position; size: Size } | undefined {
  // Try spPr > xfrm first (most shapes), then direct xfrm (graphic frames).
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

/**
 * Placeholder types that PowerPoint treats as equivalent when matching a slide
 * placeholder against layout/master placeholders. E.g. a slide `ctrTitle` can
 * inherit from a layout `title`, and a slide `subTitle` from a layout `body`.
 */
const PLACEHOLDER_TYPE_EQUIVALENCE: Record<string, string> = {
  title: "title",
  ctrTitle: "title",
  body: "body",
  subTitle: "body",
  obj: "body",
};

/** Normalize a placeholder type to its equivalence-class key. */
export function normalizePlaceholderType(type: string | undefined): string | undefined {
  if (type === undefined) return undefined;
  return PLACEHOLDER_TYPE_EQUIVALENCE[type] ?? type;
}

/**
 * Find a matching layout placeholder (PlaceholderEntry); use entry.absoluteXfrm when present.
 *
 * Match priority (best score wins):
 *   6. exact type + same idx
 *   5. equivalent type (title≡ctrTitle, body≡subTitle≡obj) + same idx
 *   4. exact type, any idx
 *   3. equivalent type, any idx
 *   2. same idx, both without type
 *   1. same idx with mismatched types (last-resort fallback; only when both
 *      placeholders explicitly specify idx, so defaulted idx 0 never triggers it)
 * A missing idx attribute is treated as idx 0 per OOXML defaults.
 */
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

/**
 * Find the layout (or master, as fallback) placeholder shape node that a slide
 * placeholder inherits from. Used for edit-mode prompt text: the template
 * node's txBody carries the "Click to add title"-style prompt runs.
 */
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

/**
 * Walk through all slide nodes (including group children recursively)
 * and fill in missing position/size from layout/master placeholders.
 */
function resolvePlaceholderInheritance(pres: PresentationData): void {
  for (let i = 0; i < pres.slides.length; i++) {
    resolveSlidePlaceholderInheritance(pres, pres.slides[i]);
  }
}

function resolveSlidePlaceholderInheritance(pres: PresentationData, slide: SlideData): void {
  if (slide.placeholderInheritanceResolved) return;

  const layoutPath = pres.slideToLayout.get(slide.index) || slide.layoutIndex;
  const layout = layoutPath ? pres.layouts.get(layoutPath) : undefined;
  const masterPath = layoutPath ? pres.layoutToMaster.get(layoutPath) : undefined;
  const master = masterPath ? pres.masters.get(masterPath) : undefined;

  resolveNodesPlaceholders(slide.nodes, layout, master);
  slide.placeholderInheritanceResolved = true;
}

export function materializeSlideNodes(pres: PresentationData, slide: SlideData): void {
  const slideRoot = materializeSlideData(slide, pres.diagramDrawings);
  if (slideRoot) {
    pres.sourcePackage?.registerXmlRoot(slide.id, slideRoot);
  }
  resolveSlidePlaceholderInheritance(pres, slide);
}

export function materializeAllSlideNodes(pres: PresentationData): void {
  for (const slide of pres.slides) {
    materializeSlideNodes(pres, slide);
  }
}

/** Extract bodyPr from a placeholder shape node (layout or master). */
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
  /**
   * Parent group for lazy-parsed group children. Layout/master placeholders are
   * stored in slide coordinates, while group children render in the group's
   * child coordinate space, so inherited xfrm must be inverted before remap.
   */
  parentGroup?: GroupNodeData;
}

function toGroupChildXfrm(
  xfrm: { position: Position; size: Size },
  group: GroupNodeData,
): { position: Position; size: Size } {
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
  xfrm: { position: Position; size: Size },
  options: PlaceholderInheritanceOptions,
): { position: Position; size: Size } {
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
  const positionLooksDefault = node.position.y < 5; // y=0 or near top → use layout position

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

      // Inherit bodyPr from layout placeholder for text rendering (anchor, insets, etc.)
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

    // Inherit bodyPr from master placeholder as fallback
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
    // Recursively handle group children
    if (node.nodeType === "group" && "children" in node) {
      // Group children are raw SafeXmlNode, not parsed yet: skip
      // (they get parsed during rendering in GroupRenderer)
    }

    resolveNodePlaceholderInheritance(node, layout, master);
  }
}
