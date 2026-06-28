import { SafeXmlNode } from "../../parser/XmlParser";
import { parseBaseProps } from "./BaseNode";
import type { BaseNodeData } from "./BaseNode";
import { parseTextBody } from "./ShapeNode";
import type { TextBody } from "./ShapeNode";
import { emuToPx } from "../../parser/units";
import { parseOoxmlBool } from "../../parser/booleans";

export interface TableCell {
  gridSpan: number;
  rowSpan: number;
  hMerge: boolean;
  vMerge: boolean;
  textBody?: TextBody;
  properties?: SafeXmlNode;
}
export interface TableRow {
  height: number;
  cells: TableCell[];
}

export interface TableNodeData extends BaseNodeData {
  nodeType: "table";
  columns: number[];
  rows: TableRow[];
  properties?: SafeXmlNode;
  tableStyleId?: string;
}

export function parseTableNode(frameNode: SafeXmlNode): TableNodeData {
  const base = parseBaseProps(frameNode);
  const tbl = frameNode.child("graphic").child("graphicData").child("tbl");
  const tblGrid = tbl.child("tblGrid");
  const columns: number[] = [];
  for (const gridCol of tblGrid.children("gridCol"))
    columns.push(emuToPx(gridCol.numAttr("w") ?? 0));

  const rows: TableRow[] = [];
  for (const trNode of tbl.children("tr")) {
    const height = emuToPx(trNode.numAttr("h") ?? 0);
    const cells: TableCell[] = [];
    for (const tcNode of trNode.children("tc")) {
      cells.push({
        gridSpan: tcNode.numAttr("gridSpan") ?? 1,
        rowSpan: tcNode.numAttr("rowSpan") ?? 1,
        hMerge: parseOoxmlBool(tcNode.attr("hMerge")),
        vMerge: parseOoxmlBool(tcNode.attr("vMerge")),
        textBody: parseTextBody(tcNode.child("txBody")),
        properties: tcNode.child("tcPr").exists() ? tcNode.child("tcPr") : undefined,
      });
    }
    rows.push({ height, cells });
  }

  const tblPr = tbl.child("tblPr");
  let tableStyleId: string | undefined;
  const tsId = tblPr.child("tableStyleId");
  if (tsId.exists()) tableStyleId = tsId.text() || tsId.attr("val");
  if (!tableStyleId) {
    const ts = tblPr.child("tblStyle");
    if (ts.exists()) tableStyleId = ts.attr("val") ?? (ts.text() || undefined);
  }
  if (!tableStyleId) tableStyleId = tblPr.attr("tblStyle");

  return {
    ...base,
    nodeType: "table",
    columns,
    rows,
    properties: tblPr.exists() ? tblPr : undefined,
    tableStyleId,
  };
}
