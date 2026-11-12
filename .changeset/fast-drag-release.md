---
"@diceui/pptx": patch
---

A move, resize, rotate, draw, or marquee drag now ends where the pointer is released, even when it's released before the last moves render. A cell range dragged while the table re-renders is kept, and Delete pressed on the selected cells right after a re-render clears them.
