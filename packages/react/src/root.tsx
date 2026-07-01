import * as React from "react";

import { PresentationContext } from "./context";
import type { RenderProp } from "./render";
import { renderElement } from "./render";
import type { PresentationStore, PreviewInput } from "./store";
import { createPresentationStore } from "./store";

export interface RootState {
  /** Current file being presented, or `null` if none is loaded. */
  file: PreviewInput | null | undefined;
}

export interface RootProps extends Omit<React.ComponentProps<"div">, "onLoad" | "onError"> {
  /**
   * **Controlled mode** — a store created via `useCreatePresentationStore`.
   * When provided, `file`, `initialSlideIndex`, `onLoad`, and `onError` are
   * ignored; you drive the store directly (call `store.load()`, etc.).
   */
  store?: PresentationStore;
  /**
   * **Uncontrolled mode** — the file to parse and display.
   * Ignored when `store` is provided.
   */
  file?: PreviewInput | null | undefined;
  /**
   * 0-based index of the slide to navigate to after a successful parse.
   * Defaults to `0`. Ignored when `store` is provided.
   */
  defaultSlideIndex?: number;
  /** Called once the presentation has been parsed. Ignored in controlled mode. */
  onLoad?: (store: PresentationStore) => void;
  /** Called when parsing fails. Ignored in controlled mode. */
  onError?: (error: Error) => void;
  /**
   * Replace the root wrapper element.
   * - ReactElement: cloned with composed props
   * - Function: `(props, state) => ReactElement`
   */
  render?: RenderProp<RootState>;
}

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
  // Uncontrolled: stable internal store, created once via lazy-ref pattern
  const internalStoreRef = React.useRef<PresentationStore | null>(null);
  if (internalStoreRef.current === null) {
    internalStoreRef.current = createPresentationStore();
  }

  const store = externalStore ?? internalStoreRef.current;

  const onLoadRef = React.useRef(onLoad);
  onLoadRef.current = onLoad;

  const onErrorRef = React.useRef(onError);
  onErrorRef.current = onError;

  // Keep a ref so the effect dep array stays stable while still reading the
  // latest value at the moment the load resolves.
  const defaultSlideIndexRef = React.useRef(defaultSlideIndex);
  defaultSlideIndexRef.current = defaultSlideIndex;

  React.useEffect(() => {
    // Controlled mode — caller owns the store lifecycle entirely.
    if (externalStore) return;

    if (!file) {
      store.reset();
      return;
    }

    store
      .load(file, { defaultSlideIndex: defaultSlideIndexRef.current })
      .then(() => onLoadRef.current?.(store))
      .catch((err: unknown) =>
        onErrorRef.current?.(err instanceof Error ? err : new Error(String(err))),
      );
  }, [store, externalStore, file]);

  return (
    <PresentationContext.Provider value={store}>
      {renderElement(
        "div",
        { render, className, style },
        {
          state: { file },
          props: { ...rootProps, children },
        },
      )}
    </PresentationContext.Provider>
  );
}

export namespace Root {
  export type State = RootState;
  export type Props = RootProps;
}
