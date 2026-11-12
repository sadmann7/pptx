import type { ECharts } from "echarts";

import type { PdfjsConfig } from "../media/pdf";
import { BaseNodeData } from "../model/nodes/base";
import { ChartNodeData } from "../model/nodes/chart";
import { GroupNodeData } from "../model/nodes/group";
import { isPlaceholderNode, parseRenderableChild } from "../model/nodes/parser";
import { PictureNodeData } from "../model/nodes/picture";
import { ShapeNodeData } from "../model/nodes/shape";
import { TableNodeData } from "../model/nodes/table";
import { materializeSlide, PresentationData } from "../model/presentation";
import { SlideData } from "../model/slide";
import type { RelEntry } from "../ooxml/rel";
import { SafeXmlNode } from "../ooxml/xml";
import { renderBackground } from "./background";
import { renderChart } from "./chart";
import { createRenderContext, RenderContext } from "./context";
import { PPTX_ATTRS } from "./dom-attributes";
import { renderGroup } from "./group";
import { renderImage } from "./image";
import { renderShape } from "./shape";
import { renderTable } from "./table";

export interface SlideRendererOptions {
  /** Called when a single node fails to render. */
  onNodeError?: (nodeId: string, error: unknown) => void;
  /** Called when a shape's hyperlink action targets a slide (zero-based index) or an external URL. */
  onNavigate?: (target: { slideIndex?: number; url?: string }) => void;
  /** Shared media URL cache for blob URL reuse across slides. */
  mediaUrlCache?: Map<string, string>;
  /** Optional pdfjs URLs for EMF-embedded PDF fallback rendering. */
  pdfjs?: PdfjsConfig;
  /** Shared set of live ECharts instances for explicit disposal. */
  chartInstances?: Set<ECharts>;
  /**
   * Renders empty placeholders with a dashed outline and their prompt text, like PowerPoint's editing view.
   * @default false
   */
  placeholderPrompts?: boolean;
  /**
   * Clips node content to the slide bounds; set `false` to keep shapes past the edge visible, like PowerPoint's pasteboard.
   * @default true
   */
  clipContent?: boolean;
}

/** Handle returned by `renderSlide()` for disposing one slide's resources without tearing down the viewer. */
export interface SlideHandle {
  /** The rendered slide DOM element. */
  readonly element: HTMLElement;
  /** Resolves when asynchronous slide resources (for example EMF-PDF fallbacks) finish. */
  readonly ready: Promise<void>;
  /** Disposes this slide's charts, and its blob URLs when no shared cache was passed. */
  dispose(): void;
  /** Supports `using` declarations. */
  [Symbol.dispose](): void;
}

function renderNode(node: BaseNodeData, ctx: RenderContext): HTMLElement {
  switch (node.nodeType) {
    case "shape":
      return renderShape(node as ShapeNodeData, ctx);
    case "picture":
      return renderImage(node as PictureNodeData, ctx);
    case "table":
      return renderTable(node as TableNodeData, ctx);
    case "group":
      return renderGroup(node as GroupNodeData, ctx, renderNode);
    case "chart":
      return renderChart(node as ChartNodeData, ctx);
    default: {
      const el = document.createElement("div");
      el.style.position = "absolute";
      el.style.left = `${node.position.x}px`;
      el.style.top = `${node.position.y}px`;
      el.style.width = `${node.size.w}px`;
      el.style.height = `${node.size.h}px`;
      return el;
    }
  }
}

function createErrorPlaceholder(node: BaseNodeData): HTMLElement {
  const el = document.createElement("div");
  el.style.position = "absolute";
  el.style.left = `${node.position.x}px`;
  el.style.top = `${node.position.y}px`;
  el.style.width = `${node.size.w}px`;
  el.style.height = `${node.size.h}px`;
  el.style.border = "2px dashed #ff4444";
  el.style.backgroundColor = "rgba(255,68,68,0.08)";
  el.style.display = "flex";
  el.style.alignItems = "center";
  el.style.justifyContent = "center";
  el.style.color = "#cc0000";
  el.style.fontSize = "11px";
  el.style.fontFamily = "monospace";
  el.style.overflow = "hidden";
  el.style.boxSizing = "border-box";
  el.style.padding = "4px";
  el.textContent = `Render Error`;
  el.title = `Failed to render node: ${node.id} (${node.name})`;
  return el;
}

interface TemplateShapeCacheEntry {
  nodes: BaseNodeData[];
  rels?: Map<string, RelEntry>;
  partPath?: string;
  diagramDrawings?: Map<string, string>;
}

const templateShapeCache = new WeakMap<SafeXmlNode, TemplateShapeCacheEntry>();

// Master and layout placeholders only serve as inheritance templates and never render.
function parseTemplateShapes(
  spTree: SafeXmlNode,
  rels?: Map<string, RelEntry>,
  partPath?: string,
  diagramDrawings?: Map<string, string>,
): BaseNodeData[] {
  const nodes: BaseNodeData[] = [];
  if (!spTree || !spTree.exists || !spTree.exists()) return nodes;
  const parseContext = {
    rels: rels ?? new Map<string, RelEntry>(),
    partPath,
    diagramDrawings,
  };

  for (const child of spTree.allChildren()) {
    if (isPlaceholderNode(child)) continue;

    try {
      const node = parseRenderableChild(child, parseContext);
      if (node && (node.size.w > 0 || node.size.h > 0)) {
        nodes.push(node);
      }
    } catch {
      // An unparseable template shape is left out rather than failing the slide.
    }
  }
  return nodes;
}

function getTemplateShapes(
  spTree: SafeXmlNode,
  rels?: Map<string, RelEntry>,
  partPath?: string,
  diagramDrawings?: Map<string, string>,
): BaseNodeData[] {
  const cached = templateShapeCache.get(spTree);
  if (
    cached &&
    cached.rels === rels &&
    cached.partPath === partPath &&
    cached.diagramDrawings === diagramDrawings
  ) {
    return cached.nodes;
  }

  const nodes = parseTemplateShapes(spTree, rels, partPath, diagramDrawings);
  templateShapeCache.set(spTree, {
    nodes,
    rels,
    partPath,
    diagramDrawings,
  });
  return nodes;
}

// Connects slide containers to the document while they render so text autofit can measure.
let measurementHost: HTMLElement | null = null;

function getMeasurementHost(): HTMLElement {
  if (measurementHost?.isConnected) return measurementHost;
  const host = document.createElement("div");
  host.setAttribute("aria-hidden", "true");
  host.setAttribute("data-pptx-measurement-host", "");
  host.style.position = "fixed";
  host.style.left = "-100000px";
  host.style.top = "0";
  host.style.width = "0";
  host.style.height = "0";
  host.style.overflow = "hidden";
  host.style.visibility = "hidden";
  host.style.pointerEvents = "none";
  // Full containment: mutations inside the host cannot invalidate layout,
  // style, or paint of the rest of the document.
  host.style.contain = "strict";
  document.body.appendChild(host);
  measurementHost = host;
  return host;
}

// Inherited typography such as Tailwind's base styles or `.prose` would skew autofit metrics.
function resetTypography(container: HTMLElement): void {
  container.style.fontSize = "initial";
  container.style.fontFamily = "initial";
  container.style.fontWeight = "normal";
  container.style.lineHeight = "normal";
  container.style.color = "initial";
  container.style.letterSpacing = "normal";
  container.style.textDecoration = "none";
  container.style.textTransform = "none";
}

function temporarilyConnectForMeasurement(container: HTMLElement): () => void {
  if (container.isConnected) return () => undefined;

  const host = getMeasurementHost();
  host.appendChild(container);

  return () => {
    if (container.parentNode === host) {
      host.removeChild(container);
    }
  };
}

export interface ThumbnailRendererOptions {
  /** Shared media URL cache for blob URL reuse across slides. */
  mediaUrlCache?: Map<string, string>;
}

/** Renders a slide like `renderSlide()` but skips autofit, ECharts, and EMF decoding, which are invisible at thumbnail size. */
export function renderThumbnail(
  presentation: PresentationData,
  slide: SlideData,
  options?: ThumbnailRendererOptions,
): SlideHandle {
  materializeSlide(presentation, slide);

  const isSharedCache = !!options?.mediaUrlCache;

  const ctx = createRenderContext(presentation, slide, options?.mediaUrlCache);
  ctx.thumbnail = true;

  const container = document.createElement("div");
  container.style.position = "relative";
  container.style.width = `${presentation.width}px`;
  container.style.height = `${presentation.height}px`;
  container.style.overflow = "hidden";
  container.style.backgroundColor = "#FFFFFF";
  resetTypography(container);

  try {
    try {
      renderBackground(ctx, container);
    } catch {
      // Non-fatal.
    }

    if (slide.showMasterSp && ctx.layout.showMasterSp) {
      const masterCtx: RenderContext = {
        ...ctx,
        slide: { ...ctx.slide, rels: ctx.master.rels },
        partPath: ctx.masterPath,
        skipPlaceholderChildren: true,
      };
      const masterShapes = getTemplateShapes(
        ctx.master.spTree,
        ctx.master.rels,
        ctx.masterPath,
        presentation.diagramDrawings,
      );
      for (const node of masterShapes) {
        try {
          container.appendChild(renderNode(node, masterCtx));
        } catch {
          // Non-fatal.
        }
      }
    }

    if (slide.showMasterSp) {
      const layoutCtx: RenderContext = {
        ...ctx,
        slide: { ...ctx.slide, rels: ctx.layout.rels },
        partPath: ctx.layoutPath,
        skipPlaceholderChildren: true,
      };
      const layoutShapes = getTemplateShapes(
        ctx.layout.spTree,
        ctx.layout.rels,
        ctx.layoutPath,
        presentation.diagramDrawings,
      );
      for (const node of layoutShapes) {
        try {
          container.appendChild(renderNode(node, layoutCtx));
        } catch {
          // Non-fatal.
        }
      }
    }

    for (const node of slide.nodes) {
      try {
        container.appendChild(renderNode(node, ctx));
      } catch {
        // Non-fatal; skip failed nodes silently in thumbnail mode.
      }
    }
  } catch {
    // Outer guard for any unexpected renderer failure.
  }

  const mediaUrlCache = ctx.mediaUrlCache;
  let disposed = false;

  const dispose = (): void => {
    if (disposed) return;
    disposed = true;
    if (!isSharedCache) {
      for (const url of mediaUrlCache.values()) URL.revokeObjectURL(url);
      mediaUrlCache.clear();
    }
  };

  return {
    element: container,
    ready: Promise.resolve(),
    dispose,
    [Symbol.dispose](): void {
      dispose();
    },
  };
}

/** Renders a slide, with its layout and master shapes, into an HTML element. */
export function renderSlide(
  presentation: PresentationData,
  slide: SlideData,
  options?: SlideRendererOptions,
): SlideHandle {
  materializeSlide(presentation, slide);

  const isSharedCache = !!options?.mediaUrlCache;
  const chartInstances = options?.chartInstances ?? new Set<ECharts>();
  const asyncTasks: Promise<void>[] = [];
  const cleanups: (() => void)[] = [];

  const ctx = createRenderContext(
    presentation,
    slide,
    options?.mediaUrlCache,
    chartInstances,
    options?.pdfjs,
  );
  ctx.asyncTasks = asyncTasks;
  ctx.cleanups = cleanups;
  if (options?.onNavigate) {
    ctx.onNavigate = options.onNavigate;
  }
  if (options?.placeholderPrompts) {
    ctx.placeholderPrompts = true;
  }

  const container = document.createElement("div");
  container.style.position = "relative";
  container.style.width = `${presentation.width}px`;
  container.style.height = `${presentation.height}px`;
  container.style.overflow = options?.clipContent === false ? "visible" : "hidden";
  container.style.backgroundColor = "#FFFFFF";
  resetTypography(container);

  ctx.measurementRoot = container;
  const restoreMeasurementMount = temporarilyConnectForMeasurement(container);

  try {
    try {
      renderBackground(ctx, container);
    } catch (error) {
      options?.onNodeError?.("__background__", error);
    }

    // A slide's `showMasterSp="0"` hides layout shapes too; a layout's hides only master shapes.
    if (slide.showMasterSp && ctx.layout.showMasterSp) {
      const masterCtx: RenderContext = {
        ...ctx,
        slide: { ...ctx.slide, rels: ctx.master.rels },
        partPath: ctx.masterPath,
        skipPlaceholderChildren: true,
      };
      const masterShapes = getTemplateShapes(
        ctx.master.spTree,
        ctx.master.rels,
        ctx.masterPath,
        presentation.diagramDrawings,
      );
      for (const node of masterShapes) {
        try {
          const el = renderNode(node, masterCtx);
          container.appendChild(el);
        } catch {
          // Master shape errors are non-fatal.
        }
      }
    }

    if (slide.showMasterSp) {
      const layoutCtx: RenderContext = {
        ...ctx,
        slide: { ...ctx.slide, rels: ctx.layout.rels },
        partPath: ctx.layoutPath,
        skipPlaceholderChildren: true,
      };
      const layoutShapes = getTemplateShapes(
        ctx.layout.spTree,
        ctx.layout.rels,
        ctx.layoutPath,
        presentation.diagramDrawings,
      );
      for (const node of layoutShapes) {
        try {
          const el = renderNode(node, layoutCtx);
          container.appendChild(el);
        } catch {
          // Layout shape errors are non-fatal.
        }
      }
    }

    for (const node of slide.nodes) {
      try {
        const el = renderNode(node, ctx);
        // Master and layout shapes are not editable, so only slide nodes are stamped.
        el.setAttribute(PPTX_ATTRS.nodeId, node.id);
        container.appendChild(el);
      } catch (error) {
        options?.onNodeError?.(node.id, error);
        const placeholder = createErrorPlaceholder(node);
        placeholder.setAttribute(PPTX_ATTRS.nodeId, node.id);
        container.appendChild(placeholder);
      }
    }
  } finally {
    restoreMeasurementMount();
  }

  let disposed = false;
  const mediaUrlCache = ctx.mediaUrlCache;
  const ready = Promise.allSettled(asyncTasks).then(() => undefined);

  const dispose = (): void => {
    if (disposed) return;
    disposed = true;

    // Detach observers before tearing down what they watch.
    for (const cleanup of cleanups) cleanup();
    cleanups.length = 0;

    if (chartInstances) {
      for (const chart of chartInstances) {
        if (!chart.isDisposed() && container.contains(chart.getDom())) {
          chart.dispose();
          chartInstances.delete(chart);
        }
      }
    }

    // A shared cache belongs to the caller.
    if (!isSharedCache) {
      for (const url of mediaUrlCache.values()) {
        URL.revokeObjectURL(url);
      }
      mediaUrlCache.clear();
    }
  };

  return {
    element: container,
    ready,
    dispose,
    [Symbol.dispose](): void {
      dispose();
    },
  };
}
