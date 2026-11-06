---
"@diceui/pptx-core": minor
"@diceui/pptx": minor
---

Add table editing: `addTable` inserts a table, the text operations take a `cell` to edit one table cell, `insertTableRow`, `deleteTableRow`, `insertTableColumn`, and `deleteTableColumn` change its rows and columns, `setTableGrid` sizes them, `mergeTableCells` and `splitTableCell` merge and split cells, `setTableCellStyle` sets cell fill, borders, and vertical alignment, and `setTableOptions` toggles the header row, banded rows, and other table style parts. In the editor, click into a cell to edit it, Tab and Shift+Tab move between cells and select their text, arrow keys at the edge of a cell's text move to the next one, Tab on the last cell adds a row, dragging any of its borders resizes a row or column, and dragging across cells or Shift+clicking selects a range to format, merge, clear, or copy. Pasting tab-separated text fills cells. `useSelection` adds `insertTable`, `cell`, `cellRange`, `changeTable`, and `setTableCellStyle`, and `editText` takes a cell.

Undo and redo shortcuts also work while editing text in apps without `Presentation.Root`.
