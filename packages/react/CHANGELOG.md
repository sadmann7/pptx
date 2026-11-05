# @diceui/pptx

## 0.3.0

### Minor Changes

- 043a412: Add editing toolbar support and more PowerPoint-style editing to `Presentation.Selection`.

  - `useSelection` reads the selection and its text formatting, formats text and paragraphs, and selects, duplicates, deletes, reorders, and inserts text boxes and shapes, including draw tools (`startDrawing`) that place them where the slide is clicked or dragged. Pressing UI marked `data-pptx-keep-selection` keeps the selection. `anchor` positions a floating toolbar over the selection, and `mode` only turns to `"move"` once a press drags.
  - Snapping with guides (`snapping`, `snapThreshold`, Alt to place freely), rotation, resizing rotated shapes, copy and paste, arrange and duplicate shortcuts, and Tab between shapes with screen reader announcements.
  - The overlay is composable from `SelectionBox`, `ResizeHandle`, `RotateHandle`, `SnapGuide`, and `Marquee` parts, styled with `className`, `style`, and CSS variables.
  - A click now selects a text box or placeholder; click again, double-click, or type to edit. An empty text box is removed when editing moves elsewhere, and the default selection color is `#1a73e8`.
  - Resize handles sit outside shapes too small to grab between them, and a flat line shows only its two ends.
  - The thumbnail list keeps at most 50 rendered miniatures, so memory stays flat on long decks.
  - Re-exports `copyNodes`, `getPresentationFonts`, and `getPresetShapePath` with the new operation types.
  - Undoing an added or duplicated slide returns to the slide that was active before it.
  - Arrow keys in the thumbnail list follow page order after a slide is added or moved.
  - Edits, undo, and redo run one at a time; `undo()` now returns a promise, like `redo()`. Line breaks survive text editing, and Ctrl/Cmd+Z in a text box undoes formatting applied to it.

### Patch Changes

- Updated dependencies [043a412]
  - @diceui/pptx-core@0.3.0

## 0.2.2

### Patch Changes

- 5dee95f: Stop reading and writing refs during render, for React Compiler compatibility.

## 0.2.1

### Patch Changes

- Updated dependencies [9de30ee]
  - @diceui/pptx-core@0.2.1

## 0.2.0

### Minor Changes

- 3f8ac02: Add `initialFocus` to `ThumbnailList`. Loading a deck on mount no longer captures the first tab stop.

## 0.1.5

### Patch Changes

- Updated dependencies [f07e526]
  - @diceui/pptx-core@0.2.0

## 0.1.4

### Patch Changes

- Updated dependencies [86adcc8]
  - @diceui/pptx-core@0.1.4

## 0.1.3

### Patch Changes

- Updated dependencies [f82d8ae]
  - @diceui/pptx-core@0.1.3

## 0.1.2

### Patch Changes

- Updated dependencies [1a2c368]
  - @diceui/pptx-core@0.1.2

## 0.1.1

### Patch Changes

- 1f6d781: Link the documentation site from the package README and update the package description.
- Updated dependencies [1f6d781]
  - @diceui/pptx-core@0.1.1

## 0.1.0

### Minor Changes

- Initial release. Composable primitives for viewing and editing PowerPoint presentations in the browser.
