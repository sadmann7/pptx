---
"@diceui/pptx-core": patch
---

Fix text disappearing from a slide's rendered DOM when a deferred autofit measurement pass ran while the slide was detached (e.g. a thumbnail scrolled out of view). The pass now waits for the slide to reattach instead of relocating it to measure, which could delete the shape.
