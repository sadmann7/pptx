import React from "react";
import type { Stroke, TableCell, TableShape, ThemeColors } from "@pptx/parser";
import { fillToCSS, strokeToSVGAttrs } from "../render/color";
import { elementStyle } from "../render/transform";
import { ParagraphElement } from "./shared/ParagraphElement";

interface TableElementProps {
  element: TableShape;
  theme: ThemeColors;
}

export function TableElement({ element, theme }: TableElementProps) {
  const outer: React.CSSProperties = {
    ...elementStyle(element),
    // Tables always render above overlapping non-table shapes in PowerPoint,
    // regardless of spTree z-order. zIndex:1 replicates that: transparent cells
    // let background shapes show through while table text stays on top.
    zIndex: 1,
    // Don't clip: border-collapse places outer cell borders at the exact edge
    // of the table box, overflow:hidden would cut the right and bottom borders.
    overflow: "visible",
  };

  return (
    <div style={outer} data-element-type="table" data-element-id={element.id}>
      <table
        style={{
          width: "100%",
          height: "100%",
          borderCollapse: "collapse",
          tableLayout: "fixed",
        }}
      >
        <colgroup>
          {element.columnWidths.map((w, i) => (
            <col key={i} style={{ width: `${w}pt` }} />
          ))}
        </colgroup>
        <tbody>
          {element.rows.map((row, ri) => (
            <tr key={ri} style={row.height ? { height: `${row.height}pt` } : {}}>
              {row.cells.map((cell, ci) => (
                <TableCellElement key={ci} cell={cell} theme={theme} />
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/** Convert a Stroke (or undefined) to a CSS border value string, e.g. "0.8pt solid #D1D5DB". */
function borderValue(stroke: Stroke | undefined, theme: ThemeColors): string {
  if (!stroke) return "none";
  const s = strokeToSVGAttrs(stroke, theme);
  if (!s.stroke || s.stroke === "none") return "none";
  const width = s.strokeWidth || "0.5pt";
  return `${width} solid ${s.stroke}`;
}

function TableCellElement({ cell, theme }: { cell: TableCell; theme: ThemeColors }) {
  if (cell.merged) return null;

  // Use per-side borders when available; fall back to a thin gray line so
  // the table always has some grid even if the PPTX has no explicit borders.
  const fallback = "0.5pt solid #e5e7eb";
  const bL = cell.strokeLeft ? borderValue(cell.strokeLeft, theme) : fallback;
  const bR = cell.strokeRight ? borderValue(cell.strokeRight, theme) : fallback;
  const bT = cell.strokeTop ? borderValue(cell.strokeTop, theme) : fallback;
  const bB = cell.strokeBottom ? borderValue(cell.strokeBottom, theme) : fallback;

  const style: React.CSSProperties = {
    verticalAlign: "middle",
    padding: "4pt 6pt",
    borderLeft: bL,
    borderRight: bR,
    borderTop: bT,
    borderBottom: bB,
    background: fillToCSS(cell.fill, theme),
    overflow: "hidden",
  };

  const tdProps: React.TdHTMLAttributes<HTMLTableCellElement> = {
    style,
    ...(cell.rowSpan > 1 ? { rowSpan: cell.rowSpan } : {}),
    ...(cell.colSpan > 1 ? { colSpan: cell.colSpan } : {}),
  };

  return (
    <td {...tdProps}>
      {cell.paragraphs.map((p, i) => (
        <ParagraphElement key={i} paragraph={p} theme={theme} />
      ))}
    </td>
  );
}
