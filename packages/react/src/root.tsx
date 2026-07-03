import * as React from "react";

import type { SlideData } from "@diceui/pptx-parser";

import { Context } from "./context";
import type { RenderProp } from "./render";
import { renderElement } from "./render";
import type { PresentationStore, PreviewInput } from "./store";
import { createPresentationStore } from "./store";

export interface RootState {
  /**
   *  Current file being presented, or `null` if none is loaded.
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
   * Event handler called once the presentation has been parsed successfully.
   *
   * ```tsx
   * onLoad={(store) => {
   *   console.log("slides:", store.getState().presentation?.slides.length);
   * }}
   * ```
   */
  onLoad?: (store: PresentationStore) => void;

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
}

export function Root({
  file,
  defaultSlideIndex,
  children,
  className,
  style,
  render,
  onLoad,
  onError,
  ...rootProps
}: RootProps) {
  const contextStore = React.useContext(Context);

  const internalStoreRef = React.useRef<PresentationStore | null>(null);
  if (contextStore === null && internalStoreRef.current === null) {
    internalStoreRef.current = createPresentationStore();
  }

  const store = contextStore ?? internalStoreRef.current;
  if (store === null) {
    // Unreachable: one of the two branches above always assigns a store.
    throw new Error("`Presentation.Root` failed to resolve a store");
  }

  const onLoadRef = React.useRef(onLoad);
  onLoadRef.current = onLoad;

  const onErrorRef = React.useRef(onError);
  onErrorRef.current = onError;

  const defaultSlideIndexRef = React.useRef(defaultSlideIndex);
  defaultSlideIndexRef.current = defaultSlideIndex;

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
      .load(file, { defaultSlideIndex: defaultSlideIndexRef.current })
      .then(() => onLoadRef.current?.(store))
      .catch((err: unknown) => {
        // AbortError means a newer load() superseded this one so it's not a real failure.
        if (err instanceof DOMException && err.name === "AbortError") return;
        onErrorRef.current?.(err instanceof Error ? err : new Error(String(err)));
      });
  }, [store, file]);

  return (
    <Context.Provider value={store}>
      {renderElement(
        "div",
        { render, className, style },
        {
          state: { file },
          props: [{ children }, rootProps],
        },
      )}
    </Context.Provider>
  );
}

export namespace Root {
  export type State = RootState;
  export type Props = RootProps;
}
