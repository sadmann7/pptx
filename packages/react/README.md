# @diceui/pptx

Composable primitives for viewing and editing PowerPoint presentations in the browser.

## Documentation

Visit [pptx.diceui.com](https://pptx.diceui.com) to view the full documentation.

## Install

```bash
npm install @diceui/pptx
```

## Viewer

```tsx
import { Presentation } from "@diceui/pptx";

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

Load with `readOnly={false}` and add `Selection` inside `Slide` to enable drag-to-move, resize, rotation, snapping, inline text editing, multi-select, marquee selection, copy and paste, and keyboard shortcuts (nudge, delete, duplicate, arrange, and opt-in undo/redo):

```tsx
import { Presentation } from "@diceui/pptx";

export function Editor({ file }: { file: File }) {
  return (
    <Presentation.Root file={file} readOnly={false}>
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

| Component                    | Description                                                      |
| ---------------------------- | ---------------------------------------------------------------- |
| `Presentation.Provider`      | Makes a store available to descendants without rendering DOM     |
| `Presentation.Root`          | Loads a file and creates a store or inherits one from `Provider` |
| `Presentation.Viewport`      | Scrollable canvas that scales the active slide to fit            |
| `Presentation.Slide`         | Renders the active slide                                         |
| `Presentation.Selection`     | Editing overlay (move, resize, rotate, text, undo/redo)          |
| `Presentation.SelectionBox`  | Outline per selected shape, holding the handles                  |
| `Presentation.ResizeHandle`  | Resize handles of a selected shape                               |
| `Presentation.RotateHandle`  | Rotation handle of a selected shape                              |
| `Presentation.SnapGuide`     | Guides shown while a move or resize snaps                        |
| `Presentation.Marquee`       | Rubber band for marquee selection and draw tools                 |
| `Presentation.ThumbnailList` | Scrollable list of slide thumbnails                              |
| `Presentation.Loading`       | Slot for custom loading states                                   |
| `Presentation.Error`         | Slot for custom error states                                     |

All components support a `render` prop to replace the underlying element.

## Hooks

These read from the nearest `Root` or `Provider`:

```ts
const { presentation, status } = usePresentation();
const { slide, index, next, prev } = useSlide();
const { zoom, setZoom } = useZoom();
const { canUndo, canRedo, isDirty, undo, redo } = useHistory();
const { textStyle, setTextStyle, insertShape, startDrawing } = useSelection();
```

`useSelection` drives a formatting toolbar outside the slide. Mark the toolbar with `data-presentation-keep-selection` so pressing it keeps the selection.

### Uncontrolled

`Root` creates the store. Reach it from a descendant when you need `load`, `edit`, `undo`, or `save`:

```ts
const store = usePresentationStore();
const slideIndex = useSlideIndex(store, slideId);
const revision = useSlideRevision(store, slideId);
```

### Controlled

Create the store manually and pass it to `Provider`:

```ts
const store = useCreatePresentationStore();
await store.load(buffer);
await store.edit({ type: "setSolidFill", slideId, nodeId, color: "FF0000" });
const bytes = await store.save();
```

## Theming

The selection overlay reads its colors from CSS variables:

| Variable                              | Default     |
| ------------------------------------- | ----------- |
| `--presentation-selection`            | `#1a73e8`   |
| `--presentation-snap-guide`           | `#f43f5e`   |
| `--presentation-snap-guide-width`     | `1px`       |
| `--presentation-rotate-handle-offset` | `20px`      |
| `--presentation-rotate-cursor`        | `crosshair` |

```css
:root {
  --presentation-selection: #7c3aed;
}
```

## License

Apache-2.0. See [NOTICE](https://github.com/sadmann7/pptx/blob/main/NOTICE) for third-party attributions.
