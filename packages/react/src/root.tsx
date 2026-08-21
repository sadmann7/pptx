import * as React from "react";

import type { SlideData } from "@diceui/pptx-core";

import { Context, useStoreEvent } from "./context";
import { useLatestRef } from "./hook";
import type { RenderProp } from "./render";
import { renderElement } from "./render";
import type {
  EditEvent,
  HistoryChangeEvent,
  PreviewInput,
  SlideChangeEvent,
  StatusChangeEvent,
  Store,
} from "./store";
import { createStore } from "./store";

export interface RootState {
  /**
   * Current file being presented, or `null` if none is loaded.
   */
  file: PreviewInput | null | undefined;
}

export interface RootProps extends Omit<React.ComponentProps<"div">, "onLoad" | "onError"> {
  /**
   * Replace the root `<div>` with a custom element or render function.
   *
   * ```tsx
   * // Element: props are merged in
   * render={<section />}
   *
   * // Function: full control
   * render={(props, state) => <section {...props} data-file={!!state.file} />}
   * ```
   */
  render?: RenderProp<RootState>;

  /**
   * The file to parse and display. Accepts `File`, `Blob`, `ArrayBuffer`,
   * or `Uint8Array`.
   *
   * - Set to `null` to explicitly reset the viewer.
   * - Omit (`undefined`) to leave the store untouched, e.g. when the store
   *   is owned elsewhere and loaded manually via `store.load(file)`.
   *
   * Works the same with an internal store or one inherited from
   * `Presentation.Provider`: whoever sets `file` triggers the load.
   *
   * ```tsx
   * <Presentation.Root file={file} />
   * ```
   */
  file?: PreviewInput | null | undefined;

  /**
   * 0-based index of the slide to navigate to after a successful parse.
   * Also accepts a resolver called with the parsed slides, useful when the
   * target depends on content (e.g. last slide, or a specific id).
   *
   * @default 0
   *
   * ```tsx
   * // Static
   * <Presentation.Root file={file} defaultSlideIndex={2} />
   *
   * // Dynamic: last slide
   * <Presentation.Root file={file} defaultSlideIndex={(slides) => slides.length - 1} />
   * ```
   */
  defaultSlideIndex?: number | ((slides: SlideData[]) => number);

  /**
   * Zoom level to open at, where `1` equals 100%.
   *
   * Only meaningful without auto-fitting: a `Viewport autoFit` fits on mount
   * and overrides it.
   *
   * ```tsx
   * <Presentation.Root file={file} defaultZoom={0.5} />
   * ```
   */
  defaultZoom?: number;

  /**
   * When `false`, the source package is retained so the presentation can be
   * edited via `store.edit()` and saved back to a .pptx via `store.save()`.
   *
   * Follows the same convention as `<input readOnly>`: omit the prop (or
   * pass `true`) for a read-only viewer; pass `false` to enable editing.
   *
   * @default true
   */
  readOnly?: boolean;

  /**
   * Event handler called once the presentation has been parsed successfully.
   *
   * ```tsx
   * onLoad={(store) => {
   *   console.log("slides:", store.getState().presentation?.slides.length);
   * }}
   * ```
   */
  onLoad?: (store: Store) => void;

  /**
   * Event handler called when parsing fails.
   *
   * ```tsx
   * onError={(error) => {
   *   console.error("Failed to load presentation:", error.message);
   * }}
   * ```
   */
  onError?: (error: Error) => void;

  /**
   * Event handler called whenever the active slide changes, whether from
   * navigation, a completed load, an edit that moved the active slide, or a
   * reset. Inspect `reason` to tell them apart.
   *
   * ```tsx
   * onSlideChange={({ slideId, index, reason }) => {
   *   if (reason === "navigate") analytics.track("slide_viewed", { index });
   * }}
   * ```
   */
  onSlideChange?: (event: SlideChangeEvent) => void;

  /**
   * Event handler called whenever the load status changes, e.g. `"idle"` to
   * `"loading"` or `"loading"` to `"ready"`. Covers loads started by the
   * `file` prop and by `store.load()` alike.
   */
  onStatusChange?: (event: StatusChangeEvent) => void;

  /**
   * Event handler called after an edit is applied, undone, or redone.
   * Inspect `source` to tell them apart.
   *
   * ```tsx
   * onEdit={({ operation, source }) => {
   *   if (source === "edit") queueAutosave(operation);
   * }}
   * ```
   */
  onEdit?: (event: EditEvent) => void;

  /**
   * Event handler called when undo/redo availability or the unsaved-changes
   * flag moves. Use it to drive toolbar state without polling the store.
   *
   * ```tsx
   * onHistoryChange={({ canUndo, canRedo, isDirty }) => {
   *   setToolbar({ canUndo, canRedo });
   *   setHasUnsavedChanges(isDirty);
   * }}
   * ```
   */
  onHistoryChange?: (event: HistoryChangeEvent) => void;
}

export function Root({
  file,
  defaultSlideIndex,
  defaultZoom,
  readOnly,
  render,
  onLoad,
  onError,
  onSlideChange,
  onStatusChange,
  onEdit,
  onHistoryChange,
  ...rootProps
}: RootProps) {
  const contextStore = React.useContext(Context);

  const internalStoreRef = React.useRef<Store | null>(null);
  if (contextStore === null && internalStoreRef.current === null) {
    internalStoreRef.current = createStore();
  }

  const store = contextStore ?? internalStoreRef.current;
  if (store === null) {
    // Unreachable: one of the two branches above always assigns a store.
    throw new Error("`Presentation.Root` failed to resolve a store");
  }

  const onLoadRef = useLatestRef(onLoad);
  const onErrorRef = useLatestRef(onError);
  const defaultSlideIndexRef = useLatestRef(defaultSlideIndex);
  const defaultZoomRef = useLatestRef(defaultZoom);

  useStoreEvent(store, "slideChange", onSlideChange);
  useStoreEvent(store, "statusChange", onStatusChange);
  useStoreEvent(store, "edit", onEdit);
  useStoreEvent(store, "historyChange", onHistoryChange);

  React.useEffect(() => {
    // `undefined` means the file API is not in use: Root leaves the store
    // alone so an owner (e.g. via `Presentation.Provider`) can drive it.
    if (file === undefined) return;

    // `null` is an explicit request to clear the viewer.
    if (file === null) {
      store.reset();
      return;
    }

    store
      .load(file, {
        defaultSlideIndex: defaultSlideIndexRef.current,
        defaultZoom: defaultZoomRef.current,
        readOnly,
      })
      .then(() => onLoadRef.current?.(store))
      .catch((err: unknown) => {
        // AbortError means a newer load() superseded this one so it's not a real failure.
        if (err instanceof DOMException && err.name === "AbortError") return;
        onErrorRef.current?.(err instanceof Error ? err : new Error(String(err)));
      });
  }, [store, file, readOnly, defaultSlideIndexRef, defaultZoomRef, onLoadRef, onErrorRef]);

  return (
    <Context.Provider value={store}>
      {renderElement(
        "div",
        { render },
        {
          state: { file },
          props: [rootProps],
        },
      )}
    </Context.Provider>
  );
}

export namespace Root {
  export type State = RootState;
  export type Props = RootProps;
}
