import { parseOoxmlBool } from "../ooxml/boolean";
import { RelEntry } from "../ooxml/rel";
import { parseXml, SafeXmlNode } from "../ooxml/xml";
import { parseRenderableChild, type RenderableNode } from "./nodes/parser";

export { parseOleFrameAsPicture } from "./nodes/parser";

export type SlideNode = RenderableNode;

export interface SlideData {
  /**
   * A stable id from the slide's zip path, like `"ppt/slides/slide3.xml"`,
   * that survives reorders, insertions, and deletions.
   */
  id: string;
  index: number;
  /** Whether the slide is hidden, which keeps it addressable but out of PDF exports. */
  hidden?: boolean;
  nodes: SlideNode[];
  background?: SafeXmlNode;
  layoutIndex: string;
  rels: Map<string, RelEntry>;
  /** Whether the layout and master shapes render on this slide. */
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

function findLayoutRel(rels: Map<string, RelEntry>): string {
  for (const [, entry] of rels) {
    if (entry.type.includes("slideLayout")) {
      return entry.target;
    }
  }
  return "";
}

export function parseSlide(
  root: SafeXmlNode,
  index: number,
  rels: Map<string, RelEntry>,
  slidePath: string = "",
  diagramDrawings?: Map<string, string>,
): SlideData {
  const cSld = root.child("cSld");

  const bg = cSld.child("bg");
  const background = bg.exists() ? bg : undefined;

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

  const layoutIndex = findLayoutRel(rels);

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

// Returns undefined when the slide was already materialized.
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
  const parsed = parseSlide(
    root,
    slide.index,
    slide.rels,
    slide.id,
    diagramDrawings,
  );

  slide.hidden = parsed.hidden;
  slide.nodes = parsed.nodes;
  slide.background = parsed.background;
  slide.layoutIndex = resolvedLayoutIndex || parsed.layoutIndex;
  slide.showMasterSp = parsed.showMasterSp;
  slide.nodesMaterialized = true;
  slide.sourceXml = undefined;
  return root;
}
