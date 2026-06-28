import { parseXml, SafeXmlNode } from '../parser/XmlParser';
import type { RelEntry } from '../parser/RelParser';
import { parseRenderableChild } from './RenderableChild';
import type { RenderableNode } from './RenderableChild';
import { parseOoxmlBool } from '../parser/booleans';

export type SlideNode = RenderableNode;

export interface SlideData {
  index: number;
  hidden?: boolean;
  nodes: SlideNode[];
  background?: SafeXmlNode;
  layoutIndex: string;
  rels: Map<string, RelEntry>;
  slidePath: string;
  showMasterSp: boolean;
  sourceXml?: string;
  nodesMaterialized?: boolean;
}

function findLayoutRel(rels: Map<string, RelEntry>): string {
  for (const [, entry] of rels) { if (entry.type.includes('slideLayout')) return entry.target; }
  return '';
}

export function parseSlide(root: SafeXmlNode, index: number, rels: Map<string, RelEntry>, slidePath: string = '', diagramDrawings?: Map<string, string>): SlideData {
  const cSld = root.child('cSld');
  const bg = cSld.child('bg');
  const background = bg.exists() ? bg : undefined;
  const spTree = cSld.child('spTree');
  const nodes: SlideNode[] = [];
  for (const child of spTree.allChildren()) {
    const node = parseRenderableChild(child, { rels, partPath: slidePath, diagramDrawings });
    if (node) nodes.push(node);
  }
  const layoutIndex = findLayoutRel(rels);
  const showMasterSp = parseOoxmlBool(root.attr('showMasterSp'), true);
  const hidden = !parseOoxmlBool(root.attr('show'), true);
  return { index, hidden, nodes, background, layoutIndex, rels, slidePath, showMasterSp, nodesMaterialized: true };
}

export function createLazySlide(sourceXml: string, index: number, rels: Map<string, RelEntry>, slidePath: string = ''): SlideData {
  return { index, nodes: [], layoutIndex: findLayoutRel(rels), rels, slidePath, showMasterSp: true, sourceXml, nodesMaterialized: false };
}

export function materializeSlideData(slide: SlideData, diagramDrawings?: Map<string, string>): void {
  if (slide.nodesMaterialized) return;
  if (!slide.sourceXml) { slide.nodesMaterialized = true; return; }
  const parsed = parseSlide(parseXml(slide.sourceXml), slide.index, slide.rels, slide.slidePath, diagramDrawings);
  slide.hidden = parsed.hidden;
  slide.nodes = parsed.nodes;
  slide.background = parsed.background;
  slide.layoutIndex = slide.layoutIndex || parsed.layoutIndex;
  slide.showMasterSp = parsed.showMasterSp;
  slide.nodesMaterialized = true;
  slide.sourceXml = undefined;
}
