import type { PptxFiles } from '../parser/ZipParser';
import type { MediaResolver } from '../utils/media';
import { parseXml, SafeXmlNode } from '../parser/XmlParser';
import { parseRels, resolveRelTarget } from '../parser/RelParser';
import type { RelEntry } from '../parser/RelParser';
import { emuToPx } from '../parser/units';
import { parseTheme } from './Theme';
import type { ThemeData } from './Theme';
import { parseSlide, createLazySlide, materializeSlideData } from './Slide';
import type { SlideData } from './Slide';

export interface PresentationData {
  width: number;
  height: number;
  slides: SlideData[];
  themes: Map<string, ThemeData>;
  slideToLayout: Map<number, string>;
  layoutToMaster: Map<string, string>;
  masterToTheme: Map<string, string>;
  media: Map<string, Uint8Array>;
  mediaResolver?: MediaResolver;
  tableStyles?: SafeXmlNode;
  defaultTextStyle?: SafeXmlNode;
  charts: Map<string, SafeXmlNode>;
  diagramDrawings?: Map<string, string>;
  isWps: boolean;
}

export interface BuildPresentationOptions {
  lazySlides?: boolean;
}

function basePath(filePath: string): string {
  const idx = filePath.lastIndexOf('/');
  return idx >= 0 ? filePath.substring(0, idx) : '';
}

function relsPathFor(filePath: string): string {
  const dir = basePath(filePath);
  const fileName = filePath.substring(filePath.lastIndexOf('/') + 1);
  return `${dir}/_rels/${fileName}.rels`;
}

function findRelByType(rels: Map<string, RelEntry>, typeSubstring: string): RelEntry | undefined {
  for (const [, entry] of rels) { if (entry.type.includes(typeSubstring)) return entry; }
  return undefined;
}

function findRelsByType(rels: Map<string, RelEntry>, typeSubstring: string): [string, RelEntry][] {
  const results: [string, RelEntry][] = [];
  for (const [rId, entry] of rels) { if (entry.type.includes(typeSubstring)) results.push([rId, entry]); }
  return results;
}

export function buildPresentation(files: PptxFiles, options: BuildPresentationOptions = {}): PresentationData {
  const presRoot = parseXml(files.presentation);
  const presRels = parseRels(files.presentationRels);

  const sldSz = presRoot.child('sldSz');
  const width = emuToPx(sldSz.numAttr('cx') ?? 9144000);
  const height = emuToPx(sldSz.numAttr('cy') ?? 6858000);
  const isWps = files.presentation.includes('wps') || files.presentation.includes('kso');

  const defaultTextStyle = presRoot.child('defaultTextStyle');

  // Themes
  const themes = new Map<string, ThemeData>();
  for (const [themePath, themeXml] of files.themes) {
    themes.set(themePath, parseTheme(parseXml(themeXml)));
  }

  // Master → theme mapping
  const masterToTheme = new Map<string, string>();
  for (const [masterPath] of files.slideMasters) {
    const masterRelsPath = relsPathFor(masterPath);
    const masterRelsXml = files.slideMasterRels.get(masterRelsPath);
    if (masterRelsXml) {
      const masterRels = parseRels(masterRelsXml);
      const themeRel = findRelByType(masterRels, 'theme');
      if (themeRel) masterToTheme.set(masterPath, resolveRelTarget(basePath(masterPath), themeRel.target));
    }
  }

  // Layout → master mapping
  const layoutToMaster = new Map<string, string>();
  for (const [layoutPath] of files.slideLayouts) {
    const layoutRelsPath = relsPathFor(layoutPath);
    const layoutRelsXml = files.slideLayoutRels.get(layoutRelsPath);
    if (layoutRelsXml) {
      const layoutRels = parseRels(layoutRelsXml);
      const masterRel = findRelByType(layoutRels, 'slideMaster');
      if (masterRel) layoutToMaster.set(layoutPath, resolveRelTarget(basePath(layoutPath), masterRel.target));
    }
  }

  // Charts
  const charts = new Map<string, SafeXmlNode>();
  for (const [chartPath, chartXml] of files.charts) {
    const chartRoot = parseXml(chartXml);
    if (chartRoot.exists()) charts.set(chartPath, chartRoot);
  }

  // Slide ordering from presentation.xml
  const sldIdLst = presRoot.child('sldIdLst');
  const orderedSlideTargets: string[] = [];
  for (const sldId of sldIdLst.children('sldId')) {
    const rId = sldId.attr('r:id') ?? sldId.attr('id');
    if (rId) {
      const relEntry = presRels.get(rId);
      if (relEntry) orderedSlideTargets.push(resolveRelTarget('ppt', relEntry.target));
    }
  }
  if (orderedSlideTargets.length === 0) {
    const slideRels = findRelsByType(presRels, 'slide')
      .filter(([, e]) => !e.type.includes('slideLayout') && !e.type.includes('slideMaster'));
    slideRels.sort((a, b) => {
      const numA = parseInt(a[0].replace(/\D/g, ''), 10) || 0;
      const numB = parseInt(b[0].replace(/\D/g, ''), 10) || 0;
      return numA - numB;
    });
    for (const [, entry] of slideRels) orderedSlideTargets.push(resolveRelTarget('ppt', entry.target));
  }

  // Parse slides
  const slides: SlideData[] = [];
  const slideToLayout = new Map<number, string>();
  for (let i = 0; i < orderedSlideTargets.length; i++) {
    const slidePath = orderedSlideTargets[i]!;
    const slideXml = files.slides.get(slidePath);
    if (!slideXml) continue;

    const slideRelsPath = relsPathFor(slidePath);
    const slideRelsXml = files.slideRels.get(slideRelsPath);
    const slideRels = slideRelsXml ? parseRels(slideRelsXml) : new Map<string, RelEntry>();

    const slideData = options.lazySlides
      ? createLazySlide(slideXml, i, slideRels, slidePath)
      : parseSlide(parseXml(slideXml), i, slideRels, slidePath, files.diagramDrawings);

    if (slideData.layoutIndex) {
      const layoutPath = resolveRelTarget(basePath(slidePath), slideData.layoutIndex);
      slideData.layoutIndex = layoutPath;
      slideToLayout.set(i, layoutPath);
    }
    slides.push(slideData);
  }

  // Table styles
  let tableStyles: SafeXmlNode | undefined;
  if (files.tableStyles) {
    const tsRoot = parseXml(files.tableStyles);
    if (tsRoot.exists()) tableStyles = tsRoot;
  }

  return {
    width, height, slides, themes, slideToLayout, layoutToMaster, masterToTheme,
    media: files.media, mediaResolver: files.mediaResolver, tableStyles,
    defaultTextStyle: defaultTextStyle.exists() ? defaultTextStyle : undefined,
    charts, diagramDrawings: files.diagramDrawings, isWps,
  };
}

export function materializeSlideNodes(pres: PresentationData, slide: SlideData): void {
  materializeSlideData(slide, pres.diagramDrawings);
}

export function materializeAllSlideNodes(pres: PresentationData): void {
  for (const slide of pres.slides) materializeSlideNodes(pres, slide);
}
