# pptx

Composable primitives for rendering and editing PowerPoint presentations. Supports slide thumbnails, inline editing, drag-to-move, resize, and undo/redo.

`@diceui/pptx` gives you headless building blocks (`Root`, `Viewport`, `Slide`, `ThumbnailList`, `Selection`, ...) that you assemble and style yourself. Under the hood, `@diceui/pptx-core` parses OOXML, renders slides to DOM, and provides a typed edit API with full undo/redo, all without a framework dependency.

## Packages

| Package                                | Description                                                | Bundle               |
| -------------------------------------- | ---------------------------------------------------------- | -------------------- |
| [`@diceui/pptx-core`](./packages/core) | Framework-agnostic PPTX engine (parse, render, edit, save) | 729 KB / 161 KB gzip |
| [`@diceui/pptx`](./packages/react)     | React component library built on top of the core           | 85 KB / 22 KB gzip   |

## Quick start

```bash
# React
npm install @diceui/pptx

# Core only (no React dependency)
npm install @diceui/pptx-core
```

### Viewer

```tsx
import * as Presentation from "@diceui/pptx";

export function Viewer({ file }: { file: File }) {
  return (
    <Presentation.Root file={file}>
      <Presentation.ThumbnailList />
      <Presentation.Viewport>
        <Presentation.Slide />
      </Presentation.Viewport>
    </Presentation.Root>
  );
}
```

### Viewer with editing

```tsx
import * as Presentation from "@diceui/pptx";

export function Editor({ file }: { file: File }) {
  return (
    <Presentation.Root file={file}>
      <Presentation.ThumbnailList />
      <Presentation.Viewport>
        <Presentation.Slide>
          <Presentation.Selection />
        </Presentation.Slide>
      </Presentation.Viewport>
    </Presentation.Root>
  );
}
```

### Core only (no React)

```ts
import { buildPresentation, readPptx, renderSlide } from "@diceui/pptx-core";

const files = await readPptx(arrayBuffer);
const presentation = buildPresentation(files);
const slide = presentation.slides[0];
const handle = renderSlide(presentation, slide);

document.body.appendChild(handle.element);

// When done
handle.dispose();
```

## React API

### `Presentation.Root`

The context provider. Accepts a `File`, `ArrayBuffer`, or URL string as `file`.

### `Presentation.Viewport`

Scrollable canvas area. Renders the active slide scaled to fit.

### `Presentation.Slide`

Renders the active slide. Must be inside `Presentation.Viewport`.

### `Presentation.Selection`

Editing overlay. Enables drag-to-move, resize (with Shift for aspect-ratio lock), inline text editing, multi-select, marquee selection, and keyboard shortcuts. Must be placed inside `Presentation.Slide` (as its child) so it can overlay the slide surface.

**Interaction model:**

| Action                       | Behavior                       |
| ---------------------------- | ------------------------------ |
| Click text box / placeholder | Select + enter text mode       |
| Click regular shape          | Select                         |
| Double-click regular shape   | Enter text mode                |
| Type while shape selected    | Enter text mode                |
| Drag shape                   | Move                           |
| Drag border of text box      | Move while keeping text mode   |
| Drag resize handle           | Resize                         |
| Shift + drag corner handle   | Resize preserving aspect ratio |
| Ctrl/Cmd+A                   | Select all                     |
| Shift/Ctrl+click             | Toggle shape in selection      |
| Drag empty canvas            | Marquee select                 |
| Delete / Backspace           | Delete selected shape(s)       |
| Arrow keys                   | Nudge (1 px; Shift = 10 px)    |
| Ctrl/Cmd+Z                   | Undo                           |
| Ctrl/Cmd+Y / Ctrl+Shift+Z    | Redo                           |
| Escape                       | Deselect / exit text mode      |

### `Presentation.ThumbnailList`

Slide strip. Renders a preview per slide, filling each one as it approaches the viewport so a long deck does not render all at once. Compose your own item with `Presentation.ThumbnailItem`, `Presentation.ThumbnailItemNumber`, and `Presentation.ThumbnailItemPreview`, or render it childless for the default.

### `Presentation.Error` / `Presentation.Loading`

Slot components for custom loading and error states.

### `Presentation.Provider`

Optional. `Root` creates its own store, so this is only needed when you want to own the store and drive it from outside the tree.

```tsx
const store = useCreatePresentationStore();

<Presentation.Provider store={store}>
  <Presentation.Root file={file}>{/* ... */}</Presentation.Root>
</Presentation.Provider>;
```

### Hooks

```ts
const { presentation, status } = usePresentation();
const { slide } = useSlide();
const { zoom, setZoom } = useZoom();
const store = useCreatePresentationStore();
```

## Edit operations (core)

All mutations go through `store.edit(operation)` and support undo/redo via `store.undo()` / `store.redo()`.

| Operation          | Description                                      |
| ------------------ | ------------------------------------------------ |
| `setTextRun`       | Replace a single text run's content              |
| `setTextBody`      | Replace all paragraphs and runs in a shape       |
| `setNodeTransform` | Move / resize a shape                            |
| `setSolidFill`     | Change a shape's fill color                      |
| `deleteNode`       | Delete a shape from a slide                      |
| `moveSlide`        | Reorder slides                                   |
| `duplicateSlide`   | Duplicate a slide                                |
| `deleteSlide`      | Delete a slide                                   |
| `batch`            | Group multiple operations into one undoable step |

## Development

```bash
pnpm install
pnpm dev          # starts apps/docs on http://localhost:3000
pnpm build        # build all packages
pnpm test         # run all tests
pnpm typecheck    # TypeScript type check
pnpm check        # lint + typecheck + format (oxlint + tsc + oxfmt)
```

## Project structure

```text
pptx/
├── apps/
│   ├── docs/              # Next.js docs + interactive playground
│   └── video/             # Remotion video renderer
└── packages/
    ├── core/              # @diceui/pptx-core (OOXML parse, render, edit, save)
    ├── react/             # @diceui/pptx (React primitives)
    ├── ui/                # @pptx/ui (shared shadcn/ui components)
    ├── e2e/               # @diceui/pptx-e2e (visual regression tests)
    └── config/            # Shared TypeScript / build config
```

## OOXML support

See [`packages/core/OOXML-SUPPORT.md`](./packages/core/OOXML-SUPPORT.md) for a detailed feature matrix against ECMA-376.

## Credits

- **[pptx-renderer](https://github.com/aiden0z/pptx-renderer)** (Apache-2.0): the parser was originally derived from this work and has been substantially modified, extended, and refactored.
- **[mtx-decompressor](https://github.com/ChristopherVR/mtx-decompressor)** (MPL-2.0): the embedded font decoder (`packages/core/src/fonts/mtx/`) is derived from this library and optimized for our use case.
- **[LibreOffice](https://www.libreoffice.org/)** (MPL-2.0): the predefined table style data (`packages/core/src/renderer/table-style.ts`) is derived from `predefined-table-styles.cxx`.

## License

Apache-2.0. See [LICENSE](./LICENSE).
