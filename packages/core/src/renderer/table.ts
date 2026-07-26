/**
 * Converts TableNodeData into positioned HTML table elements.
 *
 * Table style behavior follows:
 * - OOXML ECMA-376 §21.1.3.15 tblPr: firstRow, firstCol, bandRow, bandCol, lastRow, lastCol
 *   are attributes; when not specified they default to off (no styling).
 * - references/pptxjs (gen-table.ts, get-table-row-style.ts, get-table-cell-params.ts):
 *   reads tblPr attrs only (e.g. firstCol === "1"), applies style parts when attr is "1",
 *   and uses tcTxStyle from each part for cell text color/font (a:tcTxStyle under firstRow, firstCol, etc.).
 */

import { TableCell, TableNodeData } from "../model/nodes/table";
import { parseOoxmlBool } from "../ooxml/boolean";
import { emuToPx } from "../ooxml/unit";
import { SafeXmlNode } from "../ooxml/xml";
import { hexToRgb } from "../utils/color";
import { RenderContext } from "./context";
import { resolveThemeFontStack } from "./font";
import { resolveColor, resolveFill, resolveLineStyle, resolveThemeFillReference } from "./style";
import { getPredefinedTableStyle, NO_STYLE_TABLE_GRID } from "./table-style";
import { renderTextBody } from "./text";

function applyCssFillBackground(el: HTMLElement, fillCss: string): void {
  clearCssFillBackground(el);

  if (fillCss.includes("gradient") && fillCss.includes(" 0 0 / ")) {
    const bgMatch = fillCss.match(/,\s*(#[0-9a-fA-F]{3,8}|rgba?\([^)]+\)|[a-zA-Z]+)\s*$/);
    if (bgMatch && bgMatch.index !== undefined) {
      const imageLayers = fillCss.slice(0, bgMatch.index).replace(/\s+0 0\s*\/\s*8px 8px/g, "");
      el.style.backgroundImage = imageLayers;
      el.style.backgroundSize = "8px 8px";
      el.style.backgroundRepeat = "repeat";
      el.style.backgroundColor = bgMatch[1];
      return;
    }
  }

  if (
    fillCss.includes("gradient") ||
    fillCss.startsWith("url(") ||
    fillCss.includes("repeating-")
  ) {
    el.style.background = fillCss;
  } else {
    el.style.backgroundColor = fillCss;
  }
}

function clearCssFillBackground(el: HTMLElement): void {
  el.style.background = "";
  el.style.backgroundColor = "";
  el.style.backgroundImage = "";
  el.style.backgroundRepeat = "";
  el.style.backgroundSize = "";
}

// ---------------------------------------------------------------------------
// Table Style Lookup
// ---------------------------------------------------------------------------

/**
 * Find a table style node by its ID from presentation.tableStyles.
 * tableStyles XML structure: <a:tblStyleLst> <a:tblStyle styleId="{UUID}" ...>
 *
 * A table without an <a:tableStyleId> falls back to the built-in "No Style,
 * Table Grid": PowerPoint paints a plain 1pt tx1 grid for such tables. The
 * `def` attribute on <a:tblStyleLst> is not used here: it only seeds the style
 * of newly inserted tables, it is not a render-time fallback.
 */
function findTableStyle(
  tableStyleId: string | undefined,
  ctx: RenderContext,
): SafeXmlNode | undefined {
  if (!tableStyleId) return getPredefinedTableStyle(NO_STYLE_TABLE_GRID);
  if (!ctx.presentation.tableStyles) return undefined;
  const tblStyleLst = ctx.presentation.tableStyles;
  for (const style of tblStyleLst.children("tblStyle")) {
    if (style.attr("styleId") === tableStyleId) {
      return style;
    }
  }
  // Also check from root if tableStyles IS the tblStyleLst
  for (const style of tblStyleLst.children()) {
    if (style.localName === "tblStyle" && style.attr("styleId") === tableStyleId) {
      return style;
    }
  }
  // Fallback: check predefined (built-in) Office table styles not embedded in the PPTX
  return getPredefinedTableStyle(tableStyleId);
}

/**
 * Get the appropriate style section from a table style for a given cell position.
 * Priority: specific section > wholeTbl (fallback).
 */
function getStyleSections(
  tblStyle: SafeXmlNode,
  rowIdx: number,
  colIdx: number,
  totalRows: number,
  totalCols: number,
  tblPr: SafeXmlNode | undefined,
): SafeXmlNode[] {
  const sections: SafeXmlNode[] = [];

  // Style parts enabled only when tblPr has attribute "1" (or true); per spec default is off.
  // pptxjs uses attrs only (firstCol === "1"); we also accept child elements for compatibility.
  const flag = (attrName: string, childName: string): boolean => {
    if (!tblPr) return false;
    const attr = tblPr.attr(attrName);
    if (attr !== undefined) return parseOoxmlBool(attr);
    const ch = tblPr.child(childName);
    if (ch.exists()) {
      return parseOoxmlBool(ch.attr("val"), true);
    }
    return false;
  };
  const bandRow = flag("bandRow", "bandRow");
  const bandCol = flag("bandCol", "bandCol");
  const isFirstRow = flag("firstRow", "firstRow");
  const isLastRow = flag("lastRow", "lastRow");
  const isFirstCol = flag("firstCol", "firstCol");
  const isLastCol = flag("lastCol", "lastCol");

  // wholeTbl is the base (lowest priority)
  const wholeTbl = tblStyle.child("wholeTbl");
  if (wholeTbl.exists()) sections.push(wholeTbl);

  // Banding (applied on top of wholeTbl)
  if (bandRow) {
    const effectiveRow = isFirstRow ? rowIdx - 1 : rowIdx;
    if (effectiveRow >= 0 && effectiveRow % 2 === 1) {
      const band = tblStyle.child("band2H");
      if (band.exists()) sections.push(band);
    } else if (effectiveRow >= 0 && effectiveRow % 2 === 0) {
      const band = tblStyle.child("band1H");
      if (band.exists()) sections.push(band);
    }
  }

  if (bandCol) {
    if (colIdx % 2 === 1) {
      const band = tblStyle.child("band2V");
      if (band.exists()) sections.push(band);
    } else {
      const band = tblStyle.child("band1V");
      if (band.exists()) sections.push(band);
    }
  }

  // Special rows/cols (highest priority, override banding)
  if (isFirstRow && rowIdx === 0) {
    const s = tblStyle.child("firstRow");
    if (s.exists()) sections.push(s);
  }
  if (isLastRow && rowIdx === totalRows - 1) {
    const s = tblStyle.child("lastRow");
    if (s.exists()) sections.push(s);
  }
  if (isFirstCol && colIdx === 0) {
    const s = tblStyle.child("firstCol");
    if (s.exists()) sections.push(s);
  }
  if (isLastCol && colIdx === totalCols - 1) {
    const s = tblStyle.child("lastCol");
    if (s.exists()) sections.push(s);
  }

  return sections;
}

/** Resolved text properties from table style tcTxStyle. */
interface TableStyleTextProps {
  color?: string;
  bold?: boolean;
  italic?: boolean;
  fontFamily?: string | string[];
}

/**
 * Get the effective text properties from table style sections (last section with tcTxStyle wins).
 * tcTxStyle supports: b (bold), i (italic), and color children (schemeClr, solidFill, etc.).
 * When a style part (e.g. firstCol, firstRow) is applied, we use that part's tcTxStyle for cell
 * text styling so text stays readable on styled fill.
 */
function getEffectiveTableStyleTextProps(
  sections: SafeXmlNode[],
  ctx: RenderContext,
): TableStyleTextProps | undefined {
  for (let i = sections.length - 1; i >= 0; i--) {
    const tcTxStyle = sections[i].child("tcTxStyle");
    if (!tcTxStyle.exists()) continue;

    const props: TableStyleTextProps = {};

    // Bold: b="on" or b="off" (OOXML CT_TableStyleTextStyle)
    const b = tcTxStyle.attr("b");
    if (b !== undefined) props.bold = parseOoxmlBool(b);

    // Italic: i="on" or i="off"
    const italic = tcTxStyle.attr("i");
    if (italic !== undefined) props.italic = parseOoxmlBool(italic);

    // Color: child elements (schemeClr, solidFill, srgbClr, etc.)
    for (const child of tcTxStyle.allChildren()) {
      const tag = child.localName;
      if (
        tag === "schemeClr" ||
        tag === "solidFill" ||
        tag === "srgbClr" ||
        tag === "scrgbClr" ||
        tag === "prstClr" ||
        tag === "sysClr"
      ) {
        const { color, alpha } = resolveColor(child, ctx);
        const hex = color.startsWith("#") ? color : `#${color}`;
        if (alpha < 1) {
          const { r, g, b: bl } = hexToRgb(hex);
          props.color = `rgba(${r},${g},${bl},${alpha.toFixed(3)})`;
        } else {
          props.color = hex;
        }
        break;
      }
    }

    // Font family: <font><latin>/<ea>/<cs> typeface or <fontRef idx="major|minor">
    const font = tcTxStyle.child("font");
    if (font.exists()) {
      const latin = font.child("latin").attr("typeface");
      const ea = font.child("ea").attr("typeface");
      const cs = font.child("cs").attr("typeface");
      const fontStack = resolveThemeFontStack([latin, ea, cs], ctx);
      if (fontStack.length > 0) props.fontFamily = fontStack;
    }
    if (!props.fontFamily) {
      const fontRef = tcTxStyle.child("fontRef");
      if (fontRef.exists()) {
        const idx = fontRef.attr("idx");
        if (idx === "major") {
          const fontStack = resolveThemeFontStack(["+mj-lt", "+mj-ea", "+mj-cs"], ctx);
          if (fontStack.length > 0) props.fontFamily = fontStack;
        } else if (idx === "minor") {
          const fontStack = resolveThemeFontStack(["+mn-lt", "+mn-ea", "+mn-cs"], ctx);
          if (fontStack.length > 0) props.fontFamily = fontStack;
        }
      }
    }

    return props;
  }
  return undefined;
}

/**
 * Apply fill from a table style tcStyle node.
 * Structure: <a:tcStyle> <a:fill> <a:solidFill>... or <a:fillRef>...
 */
function applyStyleFill(td: HTMLElement, tcStyle: SafeXmlNode, ctx: RenderContext): boolean {
  const fill = tcStyle.child("fill");
  if (!fill.exists()) return false;

  // noFill
  const noFill = fill.child("noFill");
  if (noFill.exists()) {
    clearCssFillBackground(td);
    td.style.background = "transparent";
    return true;
  }

  // solidFill
  const solidFill = fill.child("solidFill");
  if (solidFill.exists()) {
    clearCssFillBackground(td);
    const { color, alpha } = resolveColor(solidFill, ctx);
    const hex = color.startsWith("#") ? color : `#${color}`;
    if (alpha < 1) {
      const { r, g, b } = hexToRgb(hex);
      td.style.backgroundColor = `rgba(${r},${g},${b},${alpha.toFixed(3)})`;
    } else {
      td.style.backgroundColor = hex;
    }
    return true;
  }

  // gradFill / pattFill
  const directFillCss = resolveFill(fill, ctx);
  if (directFillCss) {
    applyCssFillBackground(td, directFillCss);
    return true;
  }

  // fillRef (theme fill reference)
  const fillRef = fill.child("fillRef");
  if (fillRef.exists()) {
    const { fillCss } = resolveThemeFillReference(fillRef, ctx);
    applyCssFillBackground(td, fillCss);
    return true;
  }

  return false;
}

/** The four CSS border shorthands of a cell, as resolved from OOXML. */
type CellBorders = Partial<Record<"top" | "bottom" | "left" | "right", string>>;

/**
 * Collect borders from a table style tcStyle node.
 * Structure: <a:tcStyle> <a:tcBdr> <a:top>/<a:bottom>/<a:left>/<a:right> <a:ln>...
 */
function collectStyleBorders(
  borders: CellBorders,
  tcStyle: SafeXmlNode,
  ctx: RenderContext,
  rowIdx?: number,
  colIdx?: number,
  totalRows?: number,
  totalCols?: number,
): void {
  const tcBdr = tcStyle.child("tcBdr");
  if (!tcBdr.exists()) return;

  const borderMap: Array<[string, "top" | "bottom" | "left" | "right"]> = [
    ["top", "top"],
    ["bottom", "bottom"],
    ["left", "left"],
    ["right", "right"],
  ];

  // Map insideH/insideV to individual cell borders:
  // insideH → bottom for non-last rows, top for non-first rows
  // insideV → right for non-last cols, left for non-first cols
  const insideH = tcBdr.child("insideH");
  if (insideH.exists() && rowIdx !== undefined && totalRows !== undefined) {
    if (rowIdx < totalRows - 1) {
      borderMap.push(["insideH", "bottom"]);
    }
    if (rowIdx > 0) {
      borderMap.push(["insideH", "top"]);
    }
  }
  const insideV = tcBdr.child("insideV");
  if (insideV.exists() && colIdx !== undefined && totalCols !== undefined) {
    if (colIdx < totalCols - 1) {
      borderMap.push(["insideV", "right"]);
    }
    if (colIdx > 0) {
      borderMap.push(["insideV", "left"]);
    }
  }

  for (const [xmlName, side] of borderMap) {
    const sideNode = tcBdr.child(xmlName);
    if (!sideNode.exists()) continue;

    // Direct <a:ln> element
    const ln = sideNode.child("ln");
    if (ln.exists()) {
      const noFill = ln.child("noFill");
      if (noFill.exists()) continue;

      const style = resolveLineStyle(ln, ctx);
      if (style.width > 0 && style.color !== "transparent") {
        borders[side] = `${Math.max(style.width, 0.5)}px ${style.dash} ${style.color}`;
      }
      continue;
    }

    // <a:lnRef>: reference to theme line style (common in table styles)
    const lnRef = sideNode.child("lnRef");
    if (lnRef.exists()) {
      const idx = lnRef.numAttr("idx") ?? 0;
      if (idx === 0) continue; // idx 0 = no line

      // Resolve color from the lnRef's child color element
      const { color, alpha } = resolveColor(lnRef, ctx);
      const hex = color.startsWith("#") ? color : `#${color}`;

      // Get width from theme line style
      let width = 1; // default 1px
      if (ctx.theme.lineStyles && ctx.theme.lineStyles.length >= idx) {
        const themeLn = ctx.theme.lineStyles[idx - 1];
        const themeW = themeLn.numAttr("w") ?? 12700; // default 1pt
        width = emuToPx(themeW);
      }

      const cssColor =
        alpha < 1
          ? `rgba(${hexToRgb(hex).r},${hexToRgb(hex).g},${hexToRgb(hex).b},${alpha.toFixed(3)})`
          : hex;
      if (width > 0) {
        borders[side] = `${Math.max(width, 0.5)}px solid ${cssColor}`;
      }
    }
  }
}

/**
 * Apply table-level background from tblStyle > tblBg.
 * tblBg can contain fillRef (theme fill reference) or solidFill.
 */
function applyTableBackground(table: HTMLElement, tblStyle: SafeXmlNode, ctx: RenderContext): void {
  const tblBg = tblStyle.child("tblBg");
  if (!tblBg.exists()) return;

  // fillRef: references a theme fill style with a color override
  const fillRef = tblBg.child("fillRef");
  if (fillRef.exists()) {
    const { fillCss } = resolveThemeFillReference(fillRef, ctx);
    applyCssFillBackground(table, fillCss);
    return;
  }

  // solidFill
  const solidFill = tblBg.child("solidFill");
  if (solidFill.exists()) {
    clearCssFillBackground(table);
    const { color, alpha } = resolveColor(solidFill, ctx);
    const hex = color.startsWith("#") ? color : `#${color}`;
    if (alpha < 1) {
      const { r, g, b } = hexToRgb(hex);
      table.style.backgroundColor = `rgba(${r},${g},${b},${alpha.toFixed(3)})`;
    } else {
      table.style.backgroundColor = hex;
    }
    return;
  }

  const directFillCss = resolveFill(tblBg, ctx);
  if (directFillCss) {
    applyCssFillBackground(table, directFillCss);
  }
}

/**
 * Track whether every cell of a row resolved to the same opaque colour.
 *
 * Neighbouring cell backgrounds meet on a fractional device pixel once the
 * slide is scaled, and the anti-aliased seam exposes whatever is behind the
 * table, which on a dark row reads as a light hairline next to the border.
 * Painting the shared colour on the row fills that seam; row backgrounds sit
 * below every cell background and border, so nothing else changes.
 *
 * Returns the shared colour, or null once the row is disqualified (a cell with
 * no fill, a translucent fill, or a gradient/image fill, or differing colours).
 */
function mergeRowFill(current: string | null | undefined, td: HTMLElement): string | null {
  if (current === null) return null;

  const color = td.style.backgroundColor;
  if (td.style.backgroundImage || !color || color === "transparent" || color.startsWith("rgba(")) {
    return null;
  }
  if (current === undefined) return color;
  return current === color ? current : null;
}

/** One edge of the table's outline, collected while the rows are built. */
type OuterEdge = {
  tds: HTMLElement[];
  value: string | undefined;
  uniform: boolean;
  span: number;
};

function createOuterEdge(): OuterEdge {
  return { tds: [], value: undefined, uniform: true, span: 0 };
}

function trackOuterEdge(
  edge: OuterEdge,
  td: HTMLElement,
  value: string | undefined,
  span: number,
): void {
  if (edge.tds.length === 0) edge.value = value;
  else if (edge.value !== value) edge.uniform = false;
  edge.tds.push(td);
  edge.span += span;
}

/**
 * Move an outline edge from its cells onto the table element.
 *
 * Every cell paints its borders as its own rectangles, so two neighbouring
 * segments of the same outline meet on a fractional device pixel once the slide
 * is scaled, and their anti-aliased ends do not add up to full coverage. On an
 * interior grid line the shortfall exposes the neighbouring cell fill and stays
 * invisible; on the outline it exposes the slide background, so every grid line
 * reaching the outline leaves a light notch that reads as the line poking
 * through. One border on the table paints the edge as a single rectangle.
 *
 * Only possible when every cell along the edge resolved to the same border and
 * together they cover it completely (`expectedSpan`).
 */
function hoistOuterEdge(
  table: HTMLElement,
  edge: OuterEdge,
  side: "top" | "bottom" | "left" | "right",
  expectedSpan: number,
): void {
  if (!edge.uniform || !edge.value || edge.span !== expectedSpan) return;
  table.style.setProperty(`border-${side}`, edge.value);
  for (const td of edge.tds) {
    td.style.removeProperty(`border-${side}`);
  }
}

function tableFlipTransform(node: TableNodeData): string {
  const transforms: string[] = [];
  if (node.flipH) transforms.push("scaleX(-1)");
  if (node.flipV) transforms.push("scaleY(-1)");
  return transforms.join(" ");
}

// ---------------------------------------------------------------------------
// Table Rendering
// ---------------------------------------------------------------------------

/**
 * Render a table node into an absolutely-positioned HTML element.
 */
export function renderTable(node: TableNodeData, ctx: RenderContext): HTMLElement {
  // A table's true size is defined by its grid (sum of gridCol widths and
  // tr heights), not by the graphicFrame extent. PowerPoint ignores the
  // frame extent when laying out tables, and producers (notably Google
  // Slides exports) often leave a stale dummy value there (3000000x3000000),
  // which would crush the table into a small square.
  const gridWidth = node.columns.reduce((sum, w) => sum + w, 0);
  const gridHeight = node.rows.reduce((sum, r) => sum + r.height, 0);
  const frameW = gridWidth > 0 ? gridWidth : node.size.w;
  const frameH = gridHeight > 0 ? gridHeight : node.size.h;

  const wrapper = document.createElement("div");
  wrapper.style.position = "absolute";
  wrapper.style.left = `${node.position.x}px`;
  wrapper.style.top = `${node.position.y}px`;
  wrapper.style.width = `${frameW}px`;
  wrapper.style.height = `${frameH}px`;
  // Row heights are minimums in OOXML: PowerPoint grows rows to fit their
  // text, letting the table extend beyond the declared height. Clipping here
  // would cut off cell content mid-row.
  wrapper.style.overflow = "visible";

  // Apply transforms
  const transforms: string[] = [];
  if (node.rotation !== 0) {
    transforms.push(`rotate(${node.rotation}deg)`);
  }
  if (node.flipH) {
    transforms.push("scaleX(-1)");
  }
  if (node.flipV) {
    transforms.push("scaleY(-1)");
  }
  if (transforms.length > 0) {
    wrapper.style.transform = transforms.join(" ");
  }

  // Resolve table style
  const tblStyle = findTableStyle(node.tableStyleId, ctx);
  const tblPr = node.properties;
  const totalRows = node.rows.length;
  const totalCols = node.columns.length;

  // Create table element
  const table = document.createElement("table");
  // Separate borders (with zero spacing, so the geometry matches the collapsed
  // model) keep every cell's background covering its whole border box. Under
  // `collapse` the shared border strip is composited against whatever sits
  // behind the table instead, so at fractional zoom the anti-aliased line
  // picks up the slide background and reads as a lighter or darker line
  // depending on the fills it separates. Each shared edge is therefore painted
  // exactly once, by the cell that owns it (see the row loop below).
  table.style.borderCollapse = "separate";
  table.style.borderSpacing = "0";
  // The outline may end up on the table itself (see hoistOuterEdge); keep it
  // inside the frame instead of growing the table past the grid it describes.
  table.style.boxSizing = "border-box";
  table.style.width = "100%";
  table.style.height = "100%";
  table.style.tableLayout = "fixed";

  // Apply table background from table style (tblBg)
  if (tblStyle) {
    applyTableBackground(table, tblStyle, ctx);
  }

  // Column widths
  const totalWidth = node.columns.reduce((sum, w) => sum + w, 0);
  if (totalWidth > 0 && node.columns.length > 0) {
    const colgroup = document.createElement("colgroup");
    for (const colW of node.columns) {
      const col = document.createElement("col");
      col.style.width = `${(colW / totalWidth) * 100}%`;
      colgroup.appendChild(col);
    }
    table.appendChild(colgroup);
  }

  // Compute total row height so we can express each row as a proportion
  const totalRowHeight = node.rows.reduce((sum, r) => sum + r.height, 0);

  // Render rows.
  //
  // Shared edges are painted by the cell above/left of them: `bottomEdges`
  // holds, per grid column, the bottom border of the cell that last covered it,
  // and `leftEdge` the right border of the previous cell in the row. A cell
  // drops its own top/left border when the neighbour already paints that edge,
  // which mirrors the precedence of the collapsed-border model.
  const tbody = document.createElement("tbody");
  const bottomEdges: (string | undefined)[] = [];
  const outline = {
    top: createOuterEdge(),
    bottom: createOuterEdge(),
    left: createOuterEdge(),
    right: createOuterEdge(),
  };
  let colIdx = 0;
  for (let rowIdx = 0; rowIdx < node.rows.length; rowIdx++) {
    const row = node.rows[rowIdx];
    const tr = document.createElement("tr");
    if (row.height > 0 && totalRowHeight > 0) {
      // Use percentage heights so rows stay proportional within the
      // table's constrained height instead of expanding beyond it.
      tr.style.height = `${(row.height / totalRowHeight) * 100}%`;
    }

    colIdx = 0;
    let leftEdge: string | undefined;
    // Uniform opaque row fill, used as a backdrop for the sub-pixel seams
    // between neighbouring cell backgrounds (see mergeRowFill).
    let rowFill: string | null | undefined;
    let filledCols = 0;
    for (const cell of row.cells) {
      // Skip merged cells
      if (cell.hMerge || cell.vMerge) {
        // Horizontal merge continuation cells are already accounted for by the
        // origin cell's gridSpan. Vertical continuations still occupy their
        // own grid column.
        if (cell.vMerge && !cell.hMerge) {
          colIdx += cell.gridSpan;
        }
        continue;
      }

      const td = document.createElement("td");
      td.style.overflow = "hidden";

      // Spanning
      if (cell.gridSpan > 1) {
        td.colSpan = cell.gridSpan;
      }
      if (cell.rowSpan > 1) {
        td.rowSpan = cell.rowSpan;
      }

      // Apply table style first (as base), then direct tcPr overrides
      const borders: CellBorders = {};
      let sections: SafeXmlNode[] = [];
      if (tblStyle) {
        sections = getStyleSections(tblStyle, rowIdx, colIdx, totalRows, totalCols, tblPr);
        // Apply sections in order (later sections override earlier ones)
        for (const section of sections) {
          const tcStyle = section.child("tcStyle");
          if (tcStyle.exists()) {
            applyStyleFill(td, tcStyle, ctx);
            collectStyleBorders(borders, tcStyle, ctx, rowIdx, colIdx, totalRows, totalCols);
          }
        }
      }

      // Apply direct cell properties (override table style)
      applyCellProperties(td, cell, ctx, borders);

      if (leftEdge !== undefined) delete borders.left;
      if (bottomEdges[colIdx] !== undefined) delete borders.top;
      if (borders.top) td.style.borderTop = borders.top;
      if (borders.bottom) td.style.borderBottom = borders.bottom;
      if (borders.left) td.style.borderLeft = borders.left;
      if (borders.right) td.style.borderRight = borders.right;

      if (rowIdx === 0) trackOuterEdge(outline.top, td, borders.top, cell.gridSpan);
      if (rowIdx + cell.rowSpan === totalRows) {
        trackOuterEdge(outline.bottom, td, borders.bottom, cell.gridSpan);
      }
      if (colIdx === 0) trackOuterEdge(outline.left, td, borders.left, cell.rowSpan);
      if (colIdx + cell.gridSpan === totalCols) {
        trackOuterEdge(outline.right, td, borders.right, cell.rowSpan);
      }

      leftEdge = borders.right;
      for (let c = colIdx; c < colIdx + cell.gridSpan; c++) {
        bottomEdges[c] = borders.bottom;
      }
      rowFill = mergeRowFill(rowFill, td);
      filledCols += cell.gridSpan;

      // Resolve table style text properties (color, bold, italic from tcTxStyle)
      const textProps =
        sections.length > 0 ? getEffectiveTableStyleTextProps(sections, ctx) : undefined;

      // Render text inside cell
      if (cell.textBody) {
        const textTarget = tableFlipTransform(node) ? document.createElement("div") : td;
        const counterFlip = tableFlipTransform(node);
        if (counterFlip && textTarget !== td) {
          textTarget.style.width = "100%";
          textTarget.style.height = "100%";
          textTarget.style.transform = counterFlip;
          textTarget.style.transformOrigin = "center center";
        }
        const opts = {
          defaultLineHeight: "1",
          trimOuterParagraphSpacing: true,
          ...(textProps
            ? {
                cellTextColor: textProps.color,
                cellTextBold: textProps.bold,
                cellTextItalic: textProps.italic,
                cellTextFontFamily: textProps.fontFamily,
              }
            : {}),
        };
        renderTextBody(cell.textBody, undefined, ctx, textTarget, opts);
        if (textTarget !== td) {
          td.appendChild(textTarget);
        }
      }

      tr.appendChild(td);
      colIdx += cell.gridSpan;
    }

    // Only rows whose own cells span the full grid can safely carry a backdrop:
    // a band left to a vertically merged cell would be tinted through it.
    if (rowFill && filledCols === totalCols) {
      tr.style.backgroundColor = rowFill;
    }

    tbody.appendChild(tr);
  }

  hoistOuterEdge(table, outline.top, "top", totalCols);
  hoistOuterEdge(table, outline.bottom, "bottom", totalCols);
  hoistOuterEdge(table, outline.left, "left", totalRows);
  hoistOuterEdge(table, outline.right, "right", totalRows);

  table.appendChild(tbody);
  wrapper.appendChild(table);
  return wrapper;
}

// ---------------------------------------------------------------------------
// Cell Property Application
// ---------------------------------------------------------------------------

/**
 * Apply table cell properties (tcPr) to a <td> element, collecting the cell's
 * borders into `borders` (they are painted later, once per shared edge).
 */
function applyCellProperties(
  td: HTMLElement,
  cell: TableCell,
  ctx: RenderContext,
  borders: CellBorders,
): void {
  const tcPr = cell.properties;

  if (tcPr?.attr("horzOverflow") === "overflow") {
    td.style.overflow = "visible";
  }

  // Fill (overrides table style fill)
  if (tcPr) {
    const noFill = tcPr.child("noFill");
    if (noFill.exists()) {
      clearCssFillBackground(td);
      td.style.background = "transparent";
    } else if (tcPr.child("solidFill").exists()) {
      const solidFill = tcPr.child("solidFill");
      clearCssFillBackground(td);
      const { color, alpha } = resolveColor(solidFill, ctx);
      const hex = color.startsWith("#") ? color : `#${color}`;
      if (alpha < 1) {
        const { r, g, b } = hexToRgb(hex);
        td.style.backgroundColor = `rgba(${r},${g},${b},${alpha.toFixed(3)})`;
      } else {
        td.style.backgroundColor = hex;
      }
    } else {
      const directFillCss = resolveFill(tcPr, ctx);
      if (directFillCss) {
        applyCssFillBackground(td, directFillCss);
      }
    }

    // Borders (override table style borders)
    collectBorder(borders, tcPr, "lnT", "top", ctx);
    collectBorder(borders, tcPr, "lnB", "bottom", ctx);
    collectBorder(borders, tcPr, "lnL", "left", ctx);
    collectBorder(borders, tcPr, "lnR", "right", ctx);
  }

  // Margins / Padding
  const marL = tcPr?.numAttr("marL");
  const marR = tcPr?.numAttr("marR");
  const marT = tcPr?.numAttr("marT");
  const marB = tcPr?.numAttr("marB");

  // Default margin is 91440 EMU (0.1 inch) = ~9.6px
  const defaultMargin = 91440;
  td.style.paddingLeft = `${emuToPx(marL ?? defaultMargin)}px`;
  td.style.paddingRight = `${emuToPx(marR ?? defaultMargin)}px`;
  td.style.paddingTop = `${emuToPx(marT ?? 45720)}px`;
  td.style.paddingBottom = `${emuToPx(marB ?? 45720)}px`;

  // Vertical alignment
  const anchor = tcPr?.attr("anchor");
  const alignMap: Record<string, string> = {
    t: "top",
    ctr: "middle",
    b: "bottom",
  };
  td.style.verticalAlign = alignMap[anchor || "t"] || "top";
}

/**
 * Collect a single cell border from a tcPr line node.
 */
function collectBorder(
  borders: CellBorders,
  tcPr: SafeXmlNode,
  lineName: string,
  side: "top" | "bottom" | "left" | "right",
  ctx: RenderContext,
): void {
  const ln = tcPr.child(lineName);
  if (!ln.exists()) return;

  // noFill: the cell explicitly suppresses any border the table style set.
  const noFill = ln.child("noFill");
  if (noFill.exists()) {
    delete borders[side];
    return;
  }

  const style = resolveLineStyle(ln, ctx);
  if (style.width > 0 && style.color !== "transparent") {
    borders[side] = `${Math.max(style.width, 0.5)}px ${style.dash} ${style.color}`;
  }
}
