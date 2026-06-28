import { SafeXmlNode } from '../parser/XmlParser';
import type { RelEntry } from '../parser/RelParser';
import { parseShapeNode } from './nodes/ShapeNode';
import type { ShapeNodeData } from './nodes/ShapeNode';
import { parsePicNode } from './nodes/PicNode';
import type { PicNodeData } from './nodes/PicNode';
import { parseTableNode } from './nodes/TableNode';
import type { TableNodeData } from './nodes/TableNode';
import { parseGroupNode } from './nodes/GroupNode';
import type { GroupNodeData } from './nodes/GroupNode';
import { parseChartNode } from './nodes/ChartNode';
import type { ChartNodeData } from './nodes/ChartNode';

export type RenderableNode = ShapeNodeData | PicNodeData | TableNodeData | GroupNodeData | ChartNodeData;

interface ParseContext {
  rels: Map<string, RelEntry>;
  partPath?: string;
  diagramDrawings?: Map<string, string>;
  skipPlaceholders?: boolean;
}

const PLACEHOLDER_WRAPPERS = ['nvSpPr', 'nvPicPr', 'nvGrpSpPr', 'nvGraphicFramePr', 'nvCxnSpPr'];

export function isPlaceholderNode(node: SafeXmlNode): boolean {
  for (const wrapper of PLACEHOLDER_WRAPPERS) {
    const nv = node.child(wrapper);
    if (nv.exists() && nv.child('nvPr').child('ph').exists()) return true;
  }
  return false;
}

function isTableFrame(node: SafeXmlNode): boolean {
  return node.child('graphic').child('graphicData').child('tbl').exists();
}

function isChartFrame(node: SafeXmlNode): boolean {
  const uri = node.child('graphic').child('graphicData').attr('uri') || '';
  return uri.includes('chart');
}

export function parseRenderableChild(
  childXml: SafeXmlNode,
  ctx: ParseContext,
): RenderableNode | undefined {
  if (ctx.skipPlaceholders && isPlaceholderNode(childXml)) return undefined;

  switch (childXml.localName) {
    case 'sp':
    case 'cxnSp':
      return parseShapeNode(childXml);
    case 'pic':
      return parsePicNode(childXml);
    case 'grpSp':
      return parseGroupNode(childXml);
    case 'graphicFrame':
      if (isTableFrame(childXml)) return parseTableNode(childXml);
      if (isChartFrame(childXml)) return parseChartNode(childXml, ctx.rels, ctx.partPath ?? '');
      return undefined;
    default:
      return undefined;
  }
}
