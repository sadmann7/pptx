import type { ECharts } from "echarts";

import type { PdfjsConfig } from "../media/pdf";
import { LayoutData } from "../model/layout";
import { MasterData } from "../model/master";
import { PresentationData } from "../model/presentation";
import { SlideData } from "../model/slide";
import { ThemeData } from "../model/theme";
import { SafeXmlNode } from "../ooxml/xml";

export interface RenderContext {
  presentation: PresentationData;
  slide: SlideData;
  theme: ThemeData;
  master: MasterData;
  layout: LayoutData;
  partPath?: string;
  layoutPath?: string;
  masterPath?: string;
  mediaUrlCache: Map<string, string>;
  colorCache: Map<string, { color: string; alpha: number }>;
  asyncTasks?: Promise<void>[];
  // Run by `SlideHandle.dispose()`, which owns everything the render created.
  cleanups?: (() => void)[];
  pdfjs?: PdfjsConfig;
  chartInstances?: Set<ECharts>;
  // Parent group's fill, for children using `a:grpFill`.
  groupFillNode?: SafeXmlNode;
  // Connected root for text measurement while slide nodes are still detached.
  measurementRoot?: HTMLElement;
  skipPlaceholderChildren?: boolean;
  placeholderPrompts?: boolean;
  onNavigate?: (target: { slideIndex?: number; url?: string }) => void;
  // Skips work imperceptible at thumbnail scale: autofit measurement, ECharts
  // initialisation, and async EMF/PDF fallback rendering.
  thumbnail?: boolean;
}

export function createRenderContext(
  presentation: PresentationData,
  slide: SlideData,
  mediaUrlCache?: Map<string, string>,
  chartInstances?: Set<ECharts>,
  pdfjs?: PdfjsConfig,
): RenderContext {
  const layoutPath = presentation.slideToLayout.get(slide.index) || "";
  const masterPath = presentation.layoutToMaster.get(layoutPath) || "";
  const themePath = presentation.masterToTheme.get(masterPath) || "";

  const layout: LayoutData = presentation.layouts.get(layoutPath) || {
    placeholders: [],
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    spTree: {} as any,
    rels: new Map(),
    showMasterSp: true,
  };

  const master: MasterData = presentation.masters.get(masterPath) || {
    colorMap: new Map(),
    textStyles: {},
    placeholders: [],
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    spTree: {} as any,
    rels: new Map(),
  };

  const theme: ThemeData = presentation.themes.get(themePath) || {
    colorScheme: new Map(),
    majorFont: { latin: "Calibri", ea: "", cs: "" },
    minorFont: { latin: "Calibri", ea: "", cs: "" },
    fillStyles: [],
    bgFillStyles: [],
    lineStyles: [],
    effectStyles: [],
  };

  return {
    presentation,
    slide,
    theme,
    master,
    layout,
    partPath: slide.id,
    layoutPath,
    masterPath,
    mediaUrlCache: mediaUrlCache ?? new Map(),
    colorCache: new Map(),
    pdfjs,
    chartInstances,
  };
}
