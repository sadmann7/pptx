---
"@diceui/pptx-core": patch
"@diceui/pptx": patch
---

Load themes that a master keeps outside `ppt/theme/` (e.g. `ppt/slideMasters/theme/`), which left new slides and theme colors black. The default overlay draws its selection frame 2px wide and centered on the shape's edge, over any outline, in Google Slides' blue (`#1a73e8`), and uses a crosshair over the rotate handle.
