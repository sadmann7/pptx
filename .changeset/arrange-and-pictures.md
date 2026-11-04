---
"@diceui/pptx-core": minor
"@diceui/pptx": minor
---

Add the `duplicateNode`, `reorderNode` (front, back, forward, backward), and `addImage` (PNG, JPEG, GIF; sized from the file and centered by default) edit operations. Edits that create nodes report them in `EditResult.createdNodeIds`. `Presentation.Selection` duplicates the selection with Ctrl/Cmd+D and arranges it with Ctrl/Cmd+] and [, with Shift for front and back.
