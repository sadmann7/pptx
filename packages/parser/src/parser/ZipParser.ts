import JSZip from 'jszip';
import type { MediaResolver } from '../utils/media';

export interface PptxFiles {
  contentTypes: string;
  presentation: string;
  presentationRels: string;
  slides: Map<string, string>;
  slideRels: Map<string, string>;
  slideLayouts: Map<string, string>;
  slideLayoutRels: Map<string, string>;
  slideMasters: Map<string, string>;
  slideMasterRels: Map<string, string>;
  themes: Map<string, string>;
  themeOverrides?: Map<string, string>;
  media: Map<string, Uint8Array>;
  mediaResolver?: MediaResolver;
  tableStyles?: string;
  charts: Map<string, string>;
  chartRels?: Map<string, string>;
  chartStyles: Map<string, string>;
  chartColors: Map<string, string>;
  diagramDrawings: Map<string, string>;
}

export interface ZipParseLimits {
  maxEntries?: number;
  maxEntryUncompressedBytes?: number;
  maxTotalUncompressedBytes?: number;
  maxMediaBytes?: number;
  maxConcurrency?: number;
}

export const RECOMMENDED_ZIP_LIMITS: Required<ZipParseLimits> = Object.freeze({
  maxEntries: 4_000,
  maxEntryUncompressedBytes: 32 * 1024 * 1024,
  maxTotalUncompressedBytes: 256 * 1024 * 1024,
  maxMediaBytes: 192 * 1024 * 1024,
  maxConcurrency: 8,
});

function isMediaPath(path: string): boolean {
  return path.startsWith('ppt/media/');
}

function decodeZipPath(path: string): string {
  return path
    .split('/')
    .map((segment) => {
      try {
        return decodeURIComponent(segment);
      } catch {
        return segment;
      }
    })
    .join('/');
}

function setPathMapEntry<T>(map: Map<string, T>, path: string, value: T): void {
  map.set(path, value);
  const decodedPath = decodeZipPath(path);
  if (decodedPath !== path && !map.has(decodedPath)) {
    map.set(decodedPath, value);
  }
}

export async function parseZip(
  buffer: ArrayBuffer,
  _limits: ZipParseLimits = {},
): Promise<PptxFiles> {
  const zip = await JSZip.loadAsync(buffer);
  const entries = Object.entries(zip.files).filter(([, file]) => !file.dir);

  const result: PptxFiles = {
    contentTypes: '',
    presentation: '',
    presentationRels: '',
    slides: new Map(),
    slideRels: new Map(),
    slideLayouts: new Map(),
    slideLayoutRels: new Map(),
    slideMasters: new Map(),
    slideMasterRels: new Map(),
    themes: new Map(),
    themeOverrides: new Map(),
    media: new Map(),
    charts: new Map(),
    chartRels: new Map(),
    chartStyles: new Map(),
    chartColors: new Map(),
    diagramDrawings: new Map(),
  };

  for (const [path, file] of entries) {
    const normalizedPath = path.replace(/\\/g, '/');

    if (normalizedPath === '[Content_Types].xml') {
      result.contentTypes = await file.async('string');
      continue;
    }
    if (normalizedPath === 'ppt/presentation.xml') {
      result.presentation = await file.async('string');
      continue;
    }
    if (normalizedPath === 'ppt/_rels/presentation.xml.rels') {
      result.presentationRels = await file.async('string');
      continue;
    }
    if (normalizedPath === 'ppt/tableStyles.xml') {
      result.tableStyles = await file.async('string');
      continue;
    }

    if (isMediaPath(normalizedPath)) {
      const bytes = await file.async('uint8array');
      setPathMapEntry(result.media, normalizedPath, bytes);
      continue;
    }

    if (/^ppt\/slides\/_rels\/[^/]+\.xml\.rels$/.test(normalizedPath)) {
      setPathMapEntry(result.slideRels, normalizedPath, await file.async('string'));
      continue;
    }
    if (/^ppt\/slides\/[^/]+\.xml$/.test(normalizedPath)) {
      setPathMapEntry(result.slides, normalizedPath, await file.async('string'));
      continue;
    }
    if (/^ppt\/slideLayouts\/_rels\/[^/]+\.xml\.rels$/.test(normalizedPath)) {
      setPathMapEntry(result.slideLayoutRels, normalizedPath, await file.async('string'));
      continue;
    }
    if (/^ppt\/slideLayouts\/[^/]+\.xml$/.test(normalizedPath)) {
      setPathMapEntry(result.slideLayouts, normalizedPath, await file.async('string'));
      continue;
    }
    if (/^ppt\/slideMasters\/_rels\/[^/]+\.xml\.rels$/.test(normalizedPath)) {
      setPathMapEntry(result.slideMasterRels, normalizedPath, await file.async('string'));
      continue;
    }
    if (/^ppt\/slideMasters\/[^/]+\.xml$/.test(normalizedPath)) {
      setPathMapEntry(result.slideMasters, normalizedPath, await file.async('string'));
      continue;
    }
    if (/^ppt\/theme\/(?!themeOverride)[^/]+\.xml$/.test(normalizedPath)) {
      setPathMapEntry(result.themes, normalizedPath, await file.async('string'));
      continue;
    }
    if (/^ppt\/theme\/themeOverride[^/]*\.xml$/.test(normalizedPath)) {
      if (result.themeOverrides) {
        setPathMapEntry(result.themeOverrides, normalizedPath, await file.async('string'));
      }
      continue;
    }
    if (/^ppt\/charts\/_rels\/[^/]+\.xml\.rels$/.test(normalizedPath)) {
      if (result.chartRels) {
        setPathMapEntry(result.chartRels, normalizedPath, await file.async('string'));
      }
      continue;
    }
    if (/^ppt\/charts\/(?!style)[^/]+\.xml$/.test(normalizedPath)) {
      setPathMapEntry(result.charts, normalizedPath, await file.async('string'));
      continue;
    }
    if (/^ppt\/charts\/style[^/]*\.xml$/.test(normalizedPath)) {
      setPathMapEntry(result.chartStyles, normalizedPath, await file.async('string'));
      continue;
    }
    if (/^ppt\/charts\/colors[^/]*\.xml$/.test(normalizedPath)) {
      setPathMapEntry(result.chartColors, normalizedPath, await file.async('string'));
      continue;
    }
    if (/^ppt\/diagrams\/[^/]+\.xml$/.test(normalizedPath)) {
      setPathMapEntry(result.diagramDrawings, normalizedPath, await file.async('string'));
      continue;
    }
  }

  return result;
}

export async function parseZipLazyMedia(
  buffer: ArrayBuffer,
  limits: ZipParseLimits = {},
): Promise<PptxFiles> {
  // For now, just use the eager parser. Lazy media can be added later.
  return parseZip(buffer, limits);
}
