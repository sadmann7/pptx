import * as React from "react";

import type { SlideData } from "@diceui/pptx-parser";

import { PresentationContext } from "./context";
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

interface RootImplProps extends Omit<React.ComponentProps<"div">, "onLoad" | "onError"> {
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
}

interface ControlledRootProps extends RootImplProps {
  /**
   * Controlled mode: a store created via `useCreatePresentationStore`.
   * You drive the store directly. `file`, `defaultSlideIndex`, `onLoad`,
   * and `onError` are not available in this mode.
   *
   * ```tsx
   * const store = useCreatePresentationStore();
   *
   * async function open(file: File) {
   *   const presentation = await store.load(file, { defaultSlideIndex: 2 });
   *   console.log("slides:", presentation.slides.length);
   * }
   *
   * <Presentation.Root store={store}>...</Presentation.Root>
   * ```
   */
  store: PresentationStore;
  file?: never;
  defaultSlideIndex?: never;
  onLoad?: never;
  onError?: never;
}

interface UncontrolledRootProps extends RootImplProps {
  store?: never;

  /**
   * The file to parse and display. Accepts `File`, `Blob`, `ArrayBuffer`,
   * or `Uint8Array`. Set to `null` or `undefined` to reset the viewer.
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

export type RootProps = ControlledRootProps | UncontrolledRootProps;

export function Root({
  store: externalStore,
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
  const internalStoreRef = React.useRef<PresentationStore | null>(null);
  if (internalStoreRef.current === null) {
    internalStoreRef.current = createPresentationStore();
  }

  const store = externalStore ?? internalStoreRef.current;

  const onLoadRef = React.useRef(onLoad);
  onLoadRef.current = onLoad;

  const onErrorRef = React.useRef(onError);
  onErrorRef.current = onError;

  const defaultSlideIndexRef = React.useRef(defaultSlideIndex);
  defaultSlideIndexRef.current = defaultSlideIndex;

  React.useEffect(() => {
    if (externalStore) return;

    if (!file) {
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
  }, [store, externalStore, file]);

  return (
    <PresentationContext.Provider value={store}>
      {renderElement(
        "div",
        { render, className, style },
        {
          state: { file },
          props: [{ children }, rootProps],
        },
      )}
    </PresentationContext.Provider>
  );
}

export namespace Root {
  export type State = RootState;
  export type Props = RootProps;
}
