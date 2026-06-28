import React from 'react'
import type { TableCell, TableShape, ThemeColors } from '@pptx/parser'
import { fillToCSS, strokeToSVGAttrs } from '../render/color'
import { elementStyle } from '../render/transform'
import { ParagraphElement } from './shared/ParagraphElement'

interface TableElementProps {
  element: TableShape
  theme: ThemeColors
}

export function TableElement({ element, theme }: TableElementProps) {
  const outer: React.CSSProperties = {
    ...elementStyle(element),
    overflow: 'hidden',
  }

  return (
    <div style={outer} data-element-type="table" data-element-id={element.id}>
      <table
        style={{
          width: '100%',
          height: '100%',
          borderCollapse: 'collapse',
          tableLayout: 'fixed',
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
  )
}

function TableCellElement({
  cell,
  theme,
}: {
  cell: TableCell
  theme: ThemeColors
}) {
  if (cell.merged) return null

  const strokeAttrs = strokeToSVGAttrs(cell.stroke, theme)
  const borderColor = strokeAttrs.stroke !== 'none' ? strokeAttrs.stroke : '#d1d5db'
  const borderWidth = strokeAttrs.strokeWidth !== '0' ? strokeAttrs.strokeWidth : '0.5pt'

  const style: React.CSSProperties = {
    verticalAlign: 'middle',
    padding: '4pt 6pt',
    border: `${borderWidth} solid ${borderColor}`,
    background: fillToCSS(cell.fill, theme),
    overflow: 'hidden',
  }

  const tdProps: React.TdHTMLAttributes<HTMLTableCellElement> = {
    style,
    ...(cell.rowSpan > 1 ? { rowSpan: cell.rowSpan } : {}),
    ...(cell.colSpan > 1 ? { colSpan: cell.colSpan } : {}),
  }

  return (
    <td {...tdProps}>
      {cell.paragraphs.map((p, i) => (
        <ParagraphElement key={i} paragraph={p} theme={theme} />
      ))}
    </td>
  )
}
