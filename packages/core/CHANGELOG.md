# @diceui/pptx-core

## 0.4.0

### Minor Changes

- 2a84577: Add table editing: `addTable` inserts a table, the text operations take a `cell` to edit one table cell, `insertTableRow`, `deleteTableRow`, `insertTableColumn`, and `deleteTableColumn` change its rows and columns, `setTableGrid` sizes them, `mergeTableCells` and `splitTableCell` merge and split cells, `setTableCellStyle` sets cell fill, borders, and vertical alignment, and `setTableOptions` toggles the header row, banded rows, and other table style parts. In the editor, click into a cell to edit it, Tab and Shift+Tab move between cells and select their text, arrow keys at the edge of a cell's text move to the next one, Tab on the last cell adds a row, dragging any of its borders resizes a row or column, and dragging across cells or Shift+clicking selects a range to format, merge, clear, or copy. Pasting tab-separated text fills cells. `useSelection` adds `insertTable`, `cell`, `cellRange`, `changeTable`, and `setTableCellStyle`, and `editText` takes a cell.

  Undo and redo shortcuts also work while editing text in apps without `Presentation.Root`.

## 0.3.0

### Minor Changes

- 043a412: Add edit operations for building a slide editor.

  - `setTextStyle`, `setParagraphStyle`, `setOutline`, `addSlide`, `setSlideBackground`, `addTextBox`, `addShape`, `addImage`, `duplicateNode`, `reorderNode`, and `pasteNodes` with `copyNodes()`. Edits that create nodes report them in `EditResult.createdNodeIds`.
  - `getPresentationFonts()` and `getPresetShapePath()` for font and shape pickers. Layouts now carry their `name`.
  - `setTextStyle` rejects font sizes outside the 1–4000pt OOXML allows.
  - Fix themes nested in another folder (such as `ppt/slideMasters/theme/`) not loading, which rendered theme colors and new slides' backgrounds black.

## 0.2.1

### Patch Changes

- 9de30ee: Load embedded fonts whatever extension their part uses. Only `ppt/fonts/*.fntdata` parts were read, which is what PowerPoint writes, so decks from producers that pick another extension (Walnut Exporter writes `.dat`) had every embedded font reported missing and rendered with a fallback typeface. Any part under `ppt/fonts/` is now read.

## 0.2.0

### Minor Changes

- f07e526: Add an `onError` option to `loadEmbeddedFonts` so an embedded font that cannot be used is reported instead of silently falling back. Each report names the part, the typefaces left without it, whether it went missing, failed to decode, or was rejected by the `FontFace` API, and the decoder's message.

## 0.1.4

### Patch Changes

- 86adcc8: Support embedded fonts with CFF outlines, and only use the font worker when a part actually needs MTX decompression. The MTX decoder previously required the TrueType `glyf` and `loca` tables, so decks carrying an OpenType/CFF font (`OTTO` sfnt version) failed to decode and fell back to a substitute font; such fonts now pass the `CFF ` table through and skip TrueType glyph reconstruction. Uncompressed EOT parts, which most non-PowerPoint producers emit, are cheap enough to unwrap inline and now avoid the worker and its lazily loaded chunk entirely. When no worker is available, the main-thread fallback yields on a 5ms time budget instead of once per font part, which cuts decode time substantially for decks with many embedded fonts while keeping frames responsive.

## 0.1.3

### Patch Changes

- f82d8ae: Fix a build failure in consuming apps (`Module not found: Can't resolve './worker.ts'`) when importing the package from npm. The embedded-font worker was spawned from a relative source path that only resolved inside the repo, so the published build pointed at a file it never shipped. The worker is now bundled into the build and spawned from a blob URL, the same way the PDF renderer already worked.

## 0.1.2

### Patch Changes

- 1a2c368: Fix text disappearing from a slide's rendered DOM when a deferred autofit measurement pass ran while the slide was detached (e.g. a thumbnail scrolled out of view). The pass now waits for the slide to reattach instead of relocating it to measure, which could delete the shape.

## 0.1.1

### Patch Changes

- 1f6d781: Link the documentation site from the package README and update the package description.

## 0.1.0

### Minor Changes

- Initial release. PPTX engine for parsing, rendering, editing, and saving PowerPoint presentations in the browser.
