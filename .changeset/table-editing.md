---
"@diceui/pptx-core": minor
"@diceui/pptx": minor
---

Add table editing: `addTable` inserts a table, the text operations take a `cell` to edit one table cell, `insertTableRow`, `deleteTableRow`, `insertTableColumn`, and `deleteTableColumn` change its rows and columns, `setTableGrid` sizes them, and `mergeTableCells` and `splitTableCell` merge and split cells. In the editor, click or double-click a cell to edit it, Tab and Shift+Tab move between cells, Tab on the last cell adds a row, dragging any of its borders resizes a row or column, and dragging across cells or Shift+clicking selects a range to format or merge. `useSelection` adds `insertTable`, `cell`, `cellRange`, and `changeTable`, and `editText` takes a cell.
