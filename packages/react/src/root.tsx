import * as React from "react";

import type { SlideData } from "@diceui/pptx-core";

import { Context, RootContext, useStoreEvent } from "./context";
import { useLatestRef } from "./hooks";
import type { PrimitiveProps } from "./render";
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
  /** Current file being presented, or `null` if none is loaded. */
  file: PreviewInput | null | undefined;
}

export interface RootProps extends Omit<
  PrimitiveProps<"div", RootState>,
  "onLoad" | "onError"
> {
  /**
   * The file to parse and display, as a `File`, `Blob`, `ArrayBuffer` or `Uint8Array`.
   * Whoever sets it triggers the load, with an internal or a `Presentation.Provider` store.
   * - `null` resets the viewer.
   * - `undefined` leaves the store alone, e.g. when it's loaded with `store.load(file)`.
   *
   * ```tsx
   * <Presentation.Root file={file} />
   * ```
   */
  file?: PreviewInput | null | undefined;

  /**
   * 0-based index of the slide to open at, or a resolver called with the parsed slides.
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
   * Zoom level to open at, where `1` is 100%. A `Viewport` with `autoFit` overrides it.
   *
   * ```tsx
   * <Presentation.Root file={file} defaultZoom={0.5} />
   * ```
   */
  defaultZoom?: number;

  /**
   * When `false`, the presentation can be edited with `store.edit()` and saved with `store.save()`.
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
   * Event handler called whenever the active slide changes; `reason` tells navigation, load, edit
   * and reset apart.
   *
   * ```tsx
   * onSlideChange={({ slideId, index, reason }) => {
   *   if (reason === "navigate") analytics.track("slide_viewed", { index });
   * }}
   * ```
   */
  onSlideChange?: (event: SlideChangeEvent) => void;

  /**
   * Event handler called whenever the load status changes, from the `file` prop or `store.load()`.
   */
  onStatusChange?: (event: StatusChangeEvent) => void;

  /**
   * Event handler called after an edit is applied, undone or redone, as `source` tells.
   *
   * ```tsx
   * onEdit={({ operation, source }) => {
   *   if (source === "edit") queueAutosave(operation);
   * }}
   * ```
   */
  onEdit?: (event: EditEvent) => void;

  /**
   * Event handler called when undo/redo availability or the unsaved-changes flag changes.
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

export const Root = React.forwardRef<HTMLDivElement, RootProps>(function Root(
  {
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
  },
  forwardedRef,
) {
  const contextStore = React.useContext(Context);

  const internalStoreRef = React.useRef<Store | null>(null);
  if (contextStore === null && internalStoreRef.current === null) {
    internalStoreRef.current = createStore();
  }

  const store = contextStore ?? internalStoreRef.current;
  if (store === null) {
    throw new Error("`Presentation.Root` failed to resolve a store.");
  }

  const rootRef = React.useRef<HTMLElement | null>(null);

  const onLoadRef = useLatestRef(onLoad);
  const onErrorRef = useLatestRef(onError);
  const defaultSlideIndexRef = useLatestRef(defaultSlideIndex);
  const defaultZoomRef = useLatestRef(defaultZoom);

  useStoreEvent(store, "slideChange", onSlideChange);
  useStoreEvent(store, "statusChange", onStatusChange);
  useStoreEvent(store, "edit", onEdit);
  useStoreEvent(store, "historyChange", onHistoryChange);

  React.useEffect(() => {
    if (store !== internalStoreRef.current) return;
    return () => store.reset();
  }, [store]);

  React.useEffect(() => {
    if (file === undefined) return;
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
        if (err instanceof DOMException && err.name === "AbortError") return;
        onErrorRef.current?.(
          err instanceof Error ? err : new Error(String(err)),
        );
      });
  }, [
    store,
    file,
    readOnly,
    defaultSlideIndexRef,
    defaultZoomRef,
    onLoadRef,
    onErrorRef,
  ]);

  return (
    <Context.Provider value={store}>
      <RootContext.Provider value={rootRef}>
        {renderElement(
          "div",
          { render },
          {
            state: { file },
            ref: [rootRef, forwardedRef],
            props: [rootProps],
          },
        )}
      </RootContext.Provider>
    </Context.Provider>
  );
});

export namespace Root {
  export type State = RootState;
  export type Props = RootProps;
}
