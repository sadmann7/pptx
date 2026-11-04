---
"@diceui/pptx": minor
---

Snap shapes to the slide's edges and center and to other shapes while moving or resizing, with guides showing what lined up. Hold Alt (Option) to place freely. `Presentation.Selection` takes `snapping` to turn it off and `snapThreshold` (screen pixels, default 6) to change how close is close enough.

The overlay is now built from parts: `Presentation.SelectionBox`, `Presentation.ResizeHandle`, `Presentation.SnapGuide`, and `Presentation.Marquee`. Without children, `Presentation.Selection` renders them with the default look as before; compose them inside it to restyle or leave some out. Each part renders once per instance (outline per selected shape, handle per direction, guide per snapped line), accepts `render`, and takes `className` and `style` as values or functions of its state. The default look also reads `--presentation-snap-guide` and `--presentation-snap-guide-width`.
