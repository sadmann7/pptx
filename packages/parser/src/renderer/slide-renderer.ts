/**
 * Slide renderer — orchestrates rendering of a complete slide with all its nodes.
 */

import type { ECharts } from "echarts";

import type { PdfjsConfig } from "../media/pdf-renderer";
import { BaseNodeData } from "../model/nodes/base-node";
import { ChartNodeData } from "../model/nodes/chart-node";
import { GroupNodeData } from "../model/nodes/group-node";
import { isPlaceholderNode, parseRenderableChild } from "../model/nodes/node-parser";
import { PicNodeData } from "../model/nodes/picture-node";
import { ShapeNodeData } from "../model/nodes/shape-node";
import { TableNodeData } from "../model/nodes/table-node";
import { materializeSlideNodes, PresentationData } from "../model/presentation";
import { SlideData } from "../model/slide";
import type { RelEntry } from "../ooxml/rel-parser";
import { SafeXmlNode } from "../ooxml/xml-parser";
import { renderBackground } from "./background-renderer";
import { renderChart } from "./chart-renderer";
import { renderGroup } from "./group-renderer";
import { renderImage } from "./image-renderer";
import { createRenderContext, RenderContext } from "./render-context";
import { renderShape } from "./shape-renderer";
import { renderTable } from "./table-renderer";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface SlideRendererOptions {
  /** Called when a single node fails to render. */
  onNodeError?: (nodeId: string, error: unknown) => void;
  /**
   * Navigation callback for shape-level hyperlink actions (action buttons, etc.).
   * Called with target slide index (0-based) for slide jumps,
   * or with a URL string for external links.
   */
  onNavigate?: (target: { slideIndex?: number; url?: string }) => void;
  /** Shared media URL cache for blob URL reuse across slides. */
  mediaUrlCache?: Map<string, string>;
  /** Optional pdfjs URLs for EMF-embedded PDF fallback rendering. */
  pdfjs?: PdfjsConfig;
  /** Shared set of live ECharts instances for explicit disposal. */
  chartInstances?: Set<ECharts>;
  /**
   * Render empty slide placeholders with a dashed outline and the
   * layout/master prompt text, like PowerPoint's editing view.
   * Intended for edit mode; default `false`.
   */
  placeholderPrompts?: boolean;
}

/**
 * Per-slide resource handle returned by `renderSlide()`.
 * Allows the caller to dispose of slide-specific resources (chart instances,
 * blob URLs in standalone mode) without tearing down the whole viewer.
 */
export interface SlideHandle {
  /** The rendered slide DOM element. */
  readonly element: HTMLElement;
  /** Resolves when asynchronous slide resources (for example EMF-PDF fallbacks) finish. */
  readonly ready: Promise<void>;
  /** Dispose slide-specific resources (charts inside this slide, blob URLs if standalone). */
  dispose(): void;
  /** Support `using` declarations (TC39 Explicit Resource Management). */
  [Symbol.dispose](): void;
}

// ---------------------------------------------------------------------------
// Node Dispatch
// ---------------------------------------------------------------------------

/**
 * Dispatch a typed node to its appropriate renderer.
 * This function is also passed into GroupRenderer for recursive child rendering.
 */
function renderNode(node: BaseNodeData, ctx: RenderContext): HTMLElement {
  switch (node.nodeType) {
    case "shape":
      return renderShape(node as ShapeNodeData, ctx);
    case "picture":
      return renderImage(node as PicNodeData, ctx);
    case "table":
      return renderTable(node as TableNodeData, ctx);
    case "group":
      return renderGroup(node as GroupNodeData, ctx, renderNode);
    case "chart":
      return renderChart(node as ChartNodeData, ctx);
    default: {
      // Unknown node type — render as empty positioned div
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

// ---------------------------------------------------------------------------
// Error Placeholder
// ---------------------------------------------------------------------------

/**
 * Create a visual error placeholder at the node's position.
 */
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

// ---------------------------------------------------------------------------
// Master/Layout Shape Parsing
// ---------------------------------------------------------------------------

interface TemplateShapeCacheEntry {
  nodes: BaseNodeData[];
  rels?: Map<string, RelEntry>;
  partPath?: string;
  diagramDrawings?: Map<string, string>;
}

const templateShapeCache = new WeakMap<SafeXmlNode, TemplateShapeCacheEntry>();

/**
 * Parse and collect renderable shapes from a master or layout spTree.
 * Only includes NON-placeholder shapes (decorative elements, logos, footers).
 * Placeholder shapes are never rendered from master/layout — they only serve
 * as position/size inheritance templates.
 */
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
    // Skip ALL placeholder shapes — they're templates, not renderable content
    if (isPlaceholderNode(child)) continue;

    try {
      const node = parseRenderableChild(child, parseContext);
      // Skip empty/invisible nodes (0x0 size and no text)
      if (node && (node.size.w > 0 || node.size.h > 0)) {
        nodes.push(node);
      }
    } catch {
      // Skip unparseable template shapes silently
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

/**
 * Persistent, hidden, layout-contained host used to connect slide containers
 * to the document during render so text-autofit measurement works.
 *
 * A single host is created lazily and reused by every `renderSlide()` call.
 * The previous approach appended each slide container directly to
 * `document.body` (with per-call style mutation) and removed it afterwards —
 * two full-document layout invalidations per slide. With a persistent
 * `contain: strict` host, appending/removing slide subtrees only invalidates
 * layout inside the host, never the surrounding page. This is what makes
 * rendering thumbnails during scroll feasible.
 */
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

// ---------------------------------------------------------------------------
// Main Slide Render Function
// ---------------------------------------------------------------------------

/**
 * Options for `renderThumbnail()`.
 */
export interface ThumbnailRendererOptions {
  /** Shared media URL cache for blob URL reuse across slides. */
  mediaUrlCache?: Map<string, string>;
}

/**
 * Render a slide optimised for thumbnail display.
 *
 * Uses the same rendering pipeline as `renderSlide()` — real theme colours,
 * actual shapes, images, and text — but skips operations that are
 * imperceptible at thumbnail scale and expensive on the main thread:
 *
 * - **No DOM measurement host**: text autofit measurement is skipped; text
 *   renders without scale correction. At 100–200 px thumbnail width the
 *   difference is invisible.
 * - **No ECharts**: chart nodes render as a lightweight grey placeholder
 *   instead of initialising a full canvas-based chart engine.
 * - **No EMF/PDF async tasks**: EMF-embedded images render as empty boxes
 *   rather than triggering async PDF decode.
 *
 * The returned `SlideHandle` has an immediately resolved `ready` promise
 * (no async work is started). Dispose is still required to revoke blob
 * URLs in standalone mode.
 */
export function renderThumbnail(
  presentation: PresentationData,
  slide: SlideData,
  options?: ThumbnailRendererOptions,
): SlideHandle {
  materializeSlideNodes(presentation, slide);

  const isSharedCache = !!options?.mediaUrlCache;

  const ctx = createRenderContext(presentation, slide, options?.mediaUrlCache);
  // Signal to all renderers that they are in thumbnail mode.
  ctx.thumbnail = true;
  // No asyncTasks array — async work is never started in thumbnail mode.

  const container = document.createElement("div");
  container.style.position = "relative";
  container.style.width = `${presentation.width}px`;
  container.style.height = `${presentation.height}px`;
  container.style.overflow = "hidden";
  container.style.backgroundColor = "#FFFFFF";

  // No measurement host: skip the temporarilyConnectForMeasurement call
  // that renderSlide uses for text autofit. This avoids two full-document
  // layout invalidations per slide.

  try {
    try {
      renderBackground(ctx, container);
    } catch {
      // Non-fatal
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
          // Non-fatal
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
          // Non-fatal
        }
      }
    }

    for (const node of slide.nodes) {
      try {
        container.appendChild(renderNode(node, ctx));
      } catch {
        // Non-fatal — skip failed nodes silently in thumbnail mode
      }
    }
  } catch {
    // Outer guard for any unexpected renderer failure
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

/**
 * Render a complete slide into an HTML element.
 *
 * Rendering order:
 * 1. Background (slide → layout → master inheritance)
 * 2. Master non-placeholder shapes (behind everything)
 * 3. Layout non-placeholder shapes
 * 4. Slide shapes (on top)
 */
export function renderSlide(
  presentation: PresentationData,
  slide: SlideData,
  options?: SlideRendererOptions,
): SlideHandle {
  materializeSlideNodes(presentation, slide);

  const isSharedCache = !!options?.mediaUrlCache;
  const chartInstances = options?.chartInstances ?? new Set<ECharts>();
  const asyncTasks: Promise<void>[] = [];

  // Create render context (resolves slide -> layout -> master -> theme chain)
  const ctx = createRenderContext(
    presentation,
    slide,
    options?.mediaUrlCache,
    chartInstances,
    options?.pdfjs,
  );
  ctx.asyncTasks = asyncTasks;
  if (options?.onNavigate) {
    ctx.onNavigate = options.onNavigate;
  }
  if (options?.placeholderPrompts) {
    ctx.placeholderPrompts = true;
  }

  // Create slide container
  const container = document.createElement("div");
  container.style.position = "relative";
  container.style.width = `${presentation.width}px`;
  container.style.height = `${presentation.height}px`;
  container.style.overflow = "hidden";
  container.style.backgroundColor = "#FFFFFF";

  ctx.measurementRoot = container;
  const restoreMeasurementMount = temporarilyConnectForMeasurement(container);

  try {
    // Render background
    try {
      renderBackground(ctx, container);
    } catch (e) {
      options?.onNodeError?.("__background__", e);
    }

    // --- Render master template shapes (behind layout and slide) ---
    // Respect showMasterSp flags:
    //  - layout.showMasterSp === false  → skip master shapes
    //  - slide.showMasterSp === false   → skip both master AND layout shapes
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
          // Master shape errors are non-fatal
        }
      }
    }

    // --- Render layout template shapes ---
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
          // Layout shape errors are non-fatal
        }
      }
    }

    // --- Render slide shapes (on top) ---
    for (const node of slide.nodes) {
      try {
        const el = renderNode(node, ctx);
        // Only slide-level (editable) nodes are stamped; master/layout
        // template shapes are not part of the slide's editable content.
        el.setAttribute("data-pptx-node-id", node.id);
        container.appendChild(el);
      } catch (e) {
        options?.onNodeError?.(node.id, e);
        const placeholder = createErrorPlaceholder(node);
        placeholder.setAttribute("data-pptx-node-id", node.id);
        container.appendChild(placeholder);
      }
    }
  } finally {
    restoreMeasurementMount();
  }

  // Build SlideHandle
  let disposed = false;
  const mediaUrlCache = ctx.mediaUrlCache;
  const ready = Promise.allSettled(asyncTasks).then(() => undefined);

  const dispose = (): void => {
    if (disposed) return;
    disposed = true;

    // Dispose chart instances whose DOM is inside this slide container
    if (chartInstances) {
      for (const chart of chartInstances) {
        if (!chart.isDisposed() && container.contains(chart.getDom())) {
          chart.dispose();
          chartInstances.delete(chart);
        }
      }
    }

    // Revoke blob URLs only in standalone mode (caller doesn't own a shared cache)
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
