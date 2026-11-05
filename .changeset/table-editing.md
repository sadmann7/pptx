---
"@diceui/pptx-core": minor
"@diceui/pptx": minor
---

Add table editing: `addTable` inserts a table, the text operations take a `cell` to edit one table cell, and `insertTableRow`, `deleteTableRow`, `insertTableColumn`, and `deleteTableColumn` change its rows and columns. In the editor, click or double-click a cell to edit it, Tab and Shift+Tab move between cells, and Tab on the last cell adds a row. `useSelection` adds `insertTable`, `cell`, and `changeTable`, and `editText` takes a cell.
