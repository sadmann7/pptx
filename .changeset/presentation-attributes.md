---
"@diceui/pptx": minor
"@diceui/pptx-core": patch
---

Rename the attributes the React primitives write to `data-presentation-*`, matching the `--presentation-*` CSS variables. Toolbars and their popovers now mark themselves with `data-presentation-controls` (`CONTROLS_ATTRIBUTE`, formerly `KEEP_SELECTION_ATTRIBUTE`), and parts carry `data-presentation-selection-box`, `data-presentation-resize-handle`, `data-presentation-rotate-handle`, `data-presentation-snap-guide`, and `data-presentation-marquee`. The viewer's slide wrappers carry `data-pptx-slide-index` instead of `data-slide-index`.
