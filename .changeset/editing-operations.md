---
"@diceui/pptx-core": minor
---

Add edit operations for building a slide editor.

- `setTextStyle`, `setParagraphStyle`, `setOutline`, `addSlide`, `setSlideBackground`, `addTextBox`, `addShape`, `addImage`, `duplicateNode`, `reorderNode`, and `pasteNodes` with `copyNodes()`. Edits that create nodes report them in `EditResult.createdNodeIds`.
- `getPresentationFonts()` and `getPresetShapePath()` for font and shape pickers. Layouts now carry their `name`.
- `setTextStyle` rejects font sizes outside the 1–4000pt OOXML allows.
- Fix themes stored outside `ppt/theme/` not loading, which rendered theme colors and new slides' backgrounds black.
