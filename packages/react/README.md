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

| Component                    | Description                                                      |
| ---------------------------- | ---------------------------------------------------------------- |
| `Presentation.Provider`      | Makes a store available to descendants without rendering DOM     |
| `Presentation.Root`          | Loads a file and creates a store or inherits one from `Provider` |
| `Presentation.Viewport`      | Scrollable canvas that scales the active slide to fit            |
| `Presentation.Slide`         | Renders the active slide                                         |
| `Presentation.Selection`     | Editing overlay (move, resize, text, undo/redo)                  |
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
```

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

The selection overlay color is customizable via a CSS variable:

```css
:root {
  --presentation-selection: #7c3aed;
}
```

## License

Apache-2.0. See [NOTICE](https://github.com/sadmann7/pptx/blob/main/NOTICE) for third-party attributions.
