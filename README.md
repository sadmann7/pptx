# @diceui/pptx

A TypeScript monorepo for parsing, rendering, and editing PowerPoint (`.pptx`) files in the browser.

## Packages

| Package                                    | Description                                                  | Bundle               |
| ------------------------------------------ | ------------------------------------------------------------ | -------------------- |
| [`@diceui/pptx-parser`](./packages/parser) | Core OOXML parser and DOM renderer — no framework dependency | 729 KB / 161 KB gzip |
| [`@diceui/pptx`](./packages/react)         | React component library built on top of the parser           | 85 KB / 22 KB gzip   |

## Quick start

```bash
# React
npm install @diceui/pptx

# Parser only (no React dependency)
npm install @diceui/pptx-parser
```

### Viewer

```tsx
import * as Presentation from "@diceui/pptx";

export function Viewer({ file }: { file: File }) {
  return (
    <Presentation.Root file={file}>
      <Presentation.Sidebar />
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
      <Presentation.Sidebar />
      <Presentation.Viewport>
        <Presentation.Slide />
        <Presentation.Selection
          onUndo={(_, error) => error && toast.error("Nothing to undo")}
          onRedo={(_, error) => error && toast.error("Nothing to redo")}
          onNodeTransform={(_, error) => error && toast.error("Could not move shape")}
          onNodeDelete={(_, error) => error && toast.error("Could not delete shape")}
          onTextChange={(_, error) => error && toast.error("Could not save text")}
        />
      </Presentation.Viewport>
    </Presentation.Root>
  );
}
```

### Parser only (no React)

```ts
import { parsePptx, renderSlide } from "@diceui/pptx-parser";

const presentation = await parsePptx(arrayBuffer);
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

Editing overlay. Enables drag-to-move, resize (with Shift for aspect-ratio lock), inline text editing, multi-select, marquee selection, and keyboard shortcuts. Must be placed as a sibling to `Presentation.Slide` inside `Presentation.Viewport`.

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

Slide strip. Supports virtualized rendering for large decks.

### `Presentation.Toolbar` / `Presentation.Controls`

Pre-built toolbar and slide navigation controls.

### `Presentation.Error` / `Presentation.Loading`

Slot components for custom loading and error states.

### Hooks

```ts
const { presentation, status } = usePresentation();
const { slide } = useSlide();
const { zoom, setZoom } = useZoom();
const store = useCreatePresentationStore({ file });
```

## Edit operations (parser)

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
pnpm check-types  # TypeScript type check
pnpm check        # lint + format (oxlint)
```

## Project structure

```text
pptx/
├── apps/
│   └── docs/              # Next.js docs + interactive playground
└── packages/
    ├── parser/            # @diceui/pptx-parser — OOXML parser, renderer, edit ops
    ├── react/             # @diceui/pptx — React primitives
    └── config/            # Shared TypeScript / build config
```

## OOXML support

See [`packages/parser/OOXML-SUPPORT.md`](./packages/parser/OOXML-SUPPORT.md) for a detailed feature matrix against ECMA-376.

## License

Apache-2.0 — see [LICENSE](./LICENSE).
