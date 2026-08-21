# @diceui/pptx

Composable React primitives for rendering and editing PowerPoint presentations. Headless building blocks you assemble and style yourself: slide viewport, thumbnails, inline text editing, drag-to-move, resize, and undo/redo.

Built on [`@diceui/pptx-core`](https://www.npmjs.com/package/@diceui/pptx-core), a framework-agnostic OOXML engine.

## Install

```bash
npm install @diceui/pptx
```

## Viewer

```tsx
import * as Presentation from "@diceui/pptx";

export function Viewer({ file }: { file: File }) {
  return (
    <Presentation.Root file={file}>
      <Presentation.Viewport>
        <Presentation.Slide />
      </Presentation.Viewport>
    </Presentation.Root>
  );
}
```

## Editor

Add `Selection` inside `Slide` to enable drag-to-move, resize, inline text editing, multi-select, marquee selection, and keyboard shortcuts (undo/redo, nudge, delete):

```tsx
import * as Presentation from "@diceui/pptx";

export function Editor({ file }: { file: File }) {
  return (
    <Presentation.Root file={file}>
      <Presentation.Viewport>
        <Presentation.Slide>
          <Presentation.Selection />
        </Presentation.Slide>
      </Presentation.Viewport>
    </Presentation.Root>
  );
}
```

## Components

| Component                    | Description                                         |
| ---------------------------- | --------------------------------------------------- |
| `Presentation.Root`          | Context provider; accepts `File`, `Blob`, or buffer |
| `Presentation.Viewport`      | Scrollable canvas; scales the active slide to fit   |
| `Presentation.Slide`         | Renders the active slide                            |
| `Presentation.Selection`     | Editing overlay (move, resize, text, undo/redo)     |
| `Presentation.ThumbnailList` | Scrollable list of slide thumbnails                 |
| `Presentation.Loading`       | Slot for custom loading states                      |
| `Presentation.Error`         | Slot for custom error states                        |

All components support a `render` prop to replace the underlying element.

## Hooks

```ts
const { presentation, status } = usePresentation();
const { slide, index, next, prev } = useSlide();
const { zoom, setZoom } = useZoom();

// Controlled mode: create the store yourself and drive it manually
const store = useCreatePresentationStore();
await store.load(buffer);
await store.edit({ type: "setSolidFill", slideId, nodeId, color: "FF0000" });
const bytes = await store.save();
```

## Theming

The selection overlay color is customizable via a CSS variable:

```css
:root {
  --presentation-selection: #7c3aed;
}
```

## License

Apache-2.0. See [NOTICE](https://github.com/sadmann7/pptx/blob/main/NOTICE) for third-party attributions.
