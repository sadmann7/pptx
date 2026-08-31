/**
 * Parses a slide XML into a structured SlideData
 * with typed node objects for each shape on the slide.
 */

import { parseOoxmlBool } from "../ooxml/boolean";
import { RelEntry } from "../ooxml/rel";
import { parseXml, SafeXmlNode } from "../ooxml/xml";
import { parseRenderableChild, type RenderableNode } from "./nodes/parser";

export { parseOleFrameAsPicture } from "./nodes/parser";

export type SlideNode = RenderableNode;

export interface SlideData {
  /**
   * Stable unique identifier for this slide.
   * Derived from the slide's file path inside the PPTX ZIP
   * (e.g. `"ppt/slides/slide3.xml"`). Stable across reorders,
   * insertions, and deletions. Safe to use as a React key or
   * navigation target.
   */
  id: string;
  index: number;
  /** True when p:sld@show is false/0; hidden slides stay addressable but are skipped by PDF exports. */
  hidden?: boolean;
  nodes: SlideNode[];
  background?: SafeXmlNode;
  layoutIndex: string;
  rels: Map<string, RelEntry>;
  /** When false, shapes from the layout and master should NOT be rendered on this slide. */
  showMasterSp: boolean;
  /** @internal Raw slide XML used when slide node parsing is deferred. */
  sourceXml?: string;
  /** @internal Whether `nodes` has been parsed from `sourceXml`. */
  nodesMaterialized?: boolean;
  /** @internal Whether layout/master placeholder inheritance has been applied. */
  placeholderInheritanceResolved?: boolean;
}

function parseDefaultTrueBoolAttr(value: string | undefined): boolean {
  return parseOoxmlBool(value, true);
}

/**
 * Finds the layout relationship target from a slide's rels map.
 * The relationship type URI for slide layouts ends with "slideLayout".
 */
function findLayoutRel(rels: Map<string, RelEntry>): string {
  for (const [, entry] of rels) {
    if (entry.type.includes("slideLayout")) {
      return entry.target;
    }
  }
  return "";
}

/**
 * Parses a slide XML root (`p:sld`) into SlideData.
 *
 * @param root      Parsed XML root of the slide
 * @param index     Zero-based slide index
 * @param rels      Relationship entries for this slide
 * @param slidePath Full path to the slide file (e.g. "ppt/slides/slide1.xml")
 */
export function parseSlide(
  root: SafeXmlNode,
  index: number,
  rels: Map<string, RelEntry>,
  slidePath: string = "",
  diagramDrawings?: Map<string, string>,
): SlideData {
  const cSld = root.child("cSld");

  // --- Background ---
  const bg = cSld.child("bg");
  const background = bg.exists() ? bg : undefined;

  // --- Parse shape tree children ---
  const spTree = cSld.child("spTree");
  const nodes: SlideNode[] = [];

  for (const child of spTree.allChildren()) {
    const node = parseRenderableChild(child, {
      rels,
      partPath: slidePath,
      diagramDrawings,
    });
    if (node) {
      nodes.push(node);
    }
  }

  // --- Layout relationship ---
  const layoutIndex = findLayoutRel(rels);

  // --- showMasterSp: if false, layout/master shapes should not be rendered on this slide ---
  const showMasterSp = parseDefaultTrueBoolAttr(root.attr("showMasterSp"));
  const hidden = !parseDefaultTrueBoolAttr(root.attr("show"));

  return {
    id: slidePath,
    index,
    hidden,
    nodes,
    background,
    layoutIndex,
    rels,
    showMasterSp,
    nodesMaterialized: true,
  };
}

export function createLazySlide(
  sourceXml: string,
  index: number,
  rels: Map<string, RelEntry>,
  slidePath: string = "",
): SlideData {
  return {
    id: slidePath,
    index,
    nodes: [],
    layoutIndex: findLayoutRel(rels),
    rels,
    showMasterSp: true,
    sourceXml,
    nodesMaterialized: false,
  };
}

/**
 * Parses a lazy slide's deferred XML into nodes.
 * Returns the parsed slide XML root (for package registration), or undefined
 * when the slide was already materialized.
 */
export function materializeSlideData(
  slide: SlideData,
  diagramDrawings?: Map<string, string>,
): SafeXmlNode | undefined {
  if (slide.nodesMaterialized) return undefined;
  if (!slide.sourceXml) {
    slide.nodesMaterialized = true;
    return undefined;
  }

  const resolvedLayoutIndex = slide.layoutIndex;
  const root = parseXml(slide.sourceXml);
  const parsed = parseSlide(root, slide.index, slide.rels, slide.id, diagramDrawings);

  slide.hidden = parsed.hidden;
  slide.nodes = parsed.nodes;
  slide.background = parsed.background;
  slide.layoutIndex = resolvedLayoutIndex || parsed.layoutIndex;
  slide.showMasterSp = parsed.showMasterSp;
  slide.nodesMaterialized = true;
  slide.sourceXml = undefined;
  return root;
}
