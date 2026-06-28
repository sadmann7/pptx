import { SafeXmlNode } from "../../parser/XmlParser";
import { resolveRelTarget } from "../../parser/RelParser";
import type { RelEntry } from "../../parser/RelParser";
import { parseBaseProps } from "./BaseNode";
import type { BaseNodeData } from "./BaseNode";

export interface ChartNodeData extends BaseNodeData {
  nodeType: "chart";
  chartPath: string;
}

export function parseChartNode(
  graphicFrame: SafeXmlNode,
  slideRels: Map<string, RelEntry>,
  slidePath: string,
): ChartNodeData | undefined {
  const base = parseBaseProps(graphicFrame);
  const graphicData = graphicFrame.child("graphic").child("graphicData");

  let chartRId: string | undefined;
  for (const child of graphicData.allChildren()) {
    if (child.localName === "chart") {
      chartRId = child.attr("r:id") || child.attr("id");
      break;
    }
  }
  if (!chartRId) return undefined;

  const rel = slideRels.get(chartRId);
  if (!rel) return undefined;

  const slideDir = slidePath.substring(0, slidePath.lastIndexOf("/"));
  const chartPath = resolveRelTarget(slideDir, rel.target);

  return { ...base, nodeType: "chart" as const, chartPath };
}
