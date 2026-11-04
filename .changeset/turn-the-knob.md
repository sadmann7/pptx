---
"@diceui/pptx": minor
---

Rotate shapes by dragging a handle above the selection, around the shape's center; hold Shift to step by 15°. The handle is a new part, `Presentation.RotateHandle`, placed inside `Presentation.SelectionBox` and included in the default look. Its distance from the outline is `--presentation-rotate-handle-offset`. `Presentation.Selection` reports the new `"rotate"` mode while dragging. The rotate handle shows a rotate cursor, which `--presentation-rotate-cursor` replaces, and the resize and rotate cursors now stay put for the whole drag.
