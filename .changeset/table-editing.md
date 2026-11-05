---
"@diceui/pptx-core": minor
"@diceui/pptx": minor
---

Add table editing: `addTable` inserts a table, the text operations take a `cell` to edit one table cell, `insertTableRow`, `deleteTableRow`, `insertTableColumn`, and `deleteTableColumn` change its rows and columns, and `setTableGrid` sizes them. In the editor, click or double-click a cell to edit it, Tab and Shift+Tab move between cells, Tab on the last cell adds a row, and dragging any of its borders resizes a row or column. `useSelection` adds `insertTable`, `cell`, and `changeTable`, and `editText` takes a cell.
