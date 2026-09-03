# @diceui/pptx-core

## 0.1.3

### Patch Changes

- f82d8ae: Fix a build failure in consuming apps (`Module not found: Can't resolve './worker.ts'`) when importing the package from npm. The embedded-font worker was spawned from a relative source path that only resolved inside the repo, so the published build pointed at a file it never shipped. The worker is now bundled into the build and spawned from a blob URL, the same way the PDF renderer already worked.

## 0.1.2

### Patch Changes

- 1a2c368: Fix text disappearing from a slide's rendered DOM when a deferred autofit measurement pass ran while the slide was detached (e.g. a thumbnail scrolled out of view). The pass now waits for the slide to reattach instead of relocating it to measure, which could delete the shape.

## 0.1.1

### Patch Changes

- 1f6d781: Link the documentation site from the package README and update the package description.

## 0.1.0

### Minor Changes

- Initial release. PPTX engine for parsing, rendering, editing, and saving PowerPoint presentations in the browser.
