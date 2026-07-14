# @diceui/pptx-core

Framework-agnostic PPTX engine: parse, render, edit, and save PowerPoint presentations in the browser. No framework dependency.

Powers [`@diceui/pptx`](https://www.npmjs.com/package/@diceui/pptx), the React component library.

## Install

```bash
npm install @diceui/pptx-core
```

## Render a slide

```ts
import { buildPresentation, readPptx, renderSlide } from "@diceui/pptx-core";

const files = await readPptx(arrayBuffer);
const presentation = buildPresentation(files);
const handle = renderSlide(presentation, presentation.slides[0]);

document.body.appendChild(handle.element);

// When done
handle.dispose();
```

## Full viewer

`PptxViewer` provides navigation, zoom, thumbnails, and text search out of the box:

```ts
import { PptxViewer } from "@diceui/pptx-core";

const viewer = new PptxViewer(container);
await viewer.load(arrayBuffer);
viewer.goTo(2);
```

## Edit and save

Load with `keepSourcePackage: true` to retain the source archive, then apply typed edit operations and write back a `.pptx` that round-trips everything the parser doesn't model (animations, comments, vendor extensions):

```ts
import { applyEdit, buildPresentation, readPptx, writePptx } from "@diceui/pptx-core";

const presentation = buildPresentation(await readPptx(buffer, { keepSourcePackage: true }));

const result = await applyEdit(presentation, {
  type: "setTextRun",
  slideId: presentation.slides[0].id,
  nodeId: "2",
  paragraphIndex: 0,
  runIndex: 0,
  text: "Hello",
});

// result.undo() restores the previous state
const bytes = await writePptx(presentation);
```

Supported operations: `setTextRun`, `setTextBody`, `setNodeTransform`, `setSolidFill`, `deleteNode`, `moveSlide`, `duplicateSlide`, `deleteSlide`, and `batch` (multiple operations as one undoable step).

## Embedded fonts

Font decoding (including MTX-compressed embedded fonts) lives in a separate entry point so it stays out of your bundle unless used:

```ts
import { decompressMtx } from "@diceui/pptx-core/fonts";
```

## OOXML support

See [OOXML-SUPPORT.md](https://github.com/sadmann7/pptx/blob/main/packages/core/OOXML-SUPPORT.md) for a feature matrix against ECMA-376.

## License

Apache-2.0. Portions derived from third-party projects; see [NOTICE](./NOTICE). Files under `LICENSES/` carry their respective licenses (MPL-2.0 for the MTX font decoder and predefined table styles).
