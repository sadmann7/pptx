---
"@diceui/pptx": minor
---

Rotate shapes by dragging a handle above the selection, around the shape's center; hold Shift to step by 15°. The handle is a new part, `Presentation.RotateHandle`, placed inside `Presentation.SelectionBox` and included in the default look. Its distance from the outline is `--presentation-rotate-handle-offset`. `Presentation.Selection` reports the new `"rotate"` mode while dragging.
