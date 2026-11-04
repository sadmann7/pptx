---
"@diceui/pptx": minor
---

Add `useSelection` for editor toolbars. It reports the selection and the formatting of the selected text (`textStyle`), and offers `select`, `clear`, `setTextStyle`, and `setParagraphStyle`, which format the selected text or the word at the caret while editing text, and the selected shapes' text otherwise.

Pressing UI marked `data-pptx-keep-selection` (`KEEP_SELECTION_ATTRIBUTE`) no longer clears the selection or ends text editing, so toolbars and their portalled popovers can act on the selection.
