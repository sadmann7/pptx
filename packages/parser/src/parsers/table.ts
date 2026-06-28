import type { TableCell, TableRow, TableShape } from "../types";
import { parseFill, parseStroke } from "./fill";
import { parseTextBody } from "./text";
import { attr, attrNum, get, toArray } from "../xml";
import { emuToPoints } from "../emu";

export function parseTable(
  graphicFrameNode: Record<string, unknown>,
  id: string,
  name: string,
  position: { x: number; y: number },
  size: { width: number; height: number },
): TableShape {
  const tbl = get(graphicFrameNode, "a:graphic", "a:graphicData", "a:tbl") as
    | Record<string, unknown>
    | undefined;

  if (!tbl) {
    return { type: "table", id, name, position, size, columnWidths: [], rows: [] };
  }

  // Column widths
  const tblGrid = get(tbl, "a:tblGrid");
  const gridColNodes = toArray(get(tblGrid, "a:gridCol") as unknown[]);
  const columnWidths = gridColNodes.map((col) => {
    const w = attrNum(col, "w");
    return w !== undefined ? emuToPoints(w) : 0;
  });

  // Rows
  const trNodes = toArray(tbl["a:tr"] as unknown[]);
  const rows: TableRow[] = trNodes.map((tr) => {
    const trN = tr as Record<string, unknown>;
    const h = attrNum(trN, "h");
    const height = h !== undefined ? emuToPoints(h) : undefined;

    const tcNodes = toArray(trN["a:tc"] as unknown[]);
    const cells: TableCell[] = tcNodes.map((tc) => {
      const tcN = tc as Record<string, unknown>;
      const rowSpan = attrNum(tcN, "rowSpan") ?? 1;
      const colSpan = attrNum(tcN, "gridSpan") ?? 1;
      const vMerge = attr(tcN, "vMerge") === "1" || attr(tcN, "vMerge") === "true";
      const hMerge = attr(tcN, "hMerge") === "1" || attr(tcN, "hMerge") === "true";

      const txBody = get(tcN, "a:txBody");
      const { paragraphs } = parseTextBody(txBody);

      const tcPr = get(tcN, "a:tcPr");
      const fill = parseFill(tcPr);

      const lnL = get(tcPr, "a:lnL");
      const lnR = get(tcPr, "a:lnR");
      const lnT = get(tcPr, "a:lnT");
      const lnB = get(tcPr, "a:lnB");

      return {
        rowSpan,
        colSpan,
        paragraphs,
        fill,
        ...(lnL ? { strokeLeft: parseStroke(lnL) } : {}),
        ...(lnR ? { strokeRight: parseStroke(lnR) } : {}),
        ...(lnT ? { strokeTop: parseStroke(lnT) } : {}),
        ...(lnB ? { strokeBottom: parseStroke(lnB) } : {}),
        merged: vMerge || hMerge,
      };
    });

    return { height, cells };
  });

  return { type: "table", id, name, position, size, columnWidths, rows };
}
