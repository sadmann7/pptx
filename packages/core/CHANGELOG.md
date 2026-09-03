# @diceui/pptx-core

## 0.1.2

### Patch Changes

- 1a2c368: Fix text disappearing from a slide's rendered DOM when a deferred autofit measurement pass ran while the slide was detached (e.g. a thumbnail scrolled out of view). The pass now waits for the slide to reattach instead of relocating it to measure, which could delete the shape.

## 0.1.1

### Patch Changes

- 1f6d781: Link the documentation site from the package README and update the package description.
