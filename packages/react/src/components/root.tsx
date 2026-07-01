import * as React from "react";
import { createPresentationStore } from "../store";
import type { PresentationStore, PreviewInput } from "../store";
import { PresentationContext } from "../context";
import { renderElement } from "../utils/render";
import type { RenderProp } from "../utils/render";

export interface RootState {
  /** Current file being presented, or `null` if none is loaded. */
  file: PreviewInput | null | undefined;
}

export interface RootProps extends Omit<React.ComponentProps<"div">, "onLoad" | "onError"> {
  file: PreviewInput | null | undefined;
  onLoad?: (store: PresentationStore) => void;
  onError?: (error: Error) => void;
  /**
   * Replace the root wrapper element.
   * - ReactElement: cloned with composed props
   * - Function: `(props, state) => ReactElement`
   */
  render?: RenderProp<RootState>;
}

export namespace Root {
  export type State = RootState;
  export type Props = RootProps;
}

export function Root({
  file,
  children,
  className,
  style,
  render,
  onLoad,
  onError,
  ...elementProps
}: RootProps) {
  const store = React.useMemo(() => createPresentationStore(), []);
  const onLoadRef = React.useRef(onLoad);
  onLoadRef.current = onLoad;
  const onErrorRef = React.useRef(onError);
  onErrorRef.current = onError;

  React.useEffect(() => {
    if (!file) {
      store.reset();
      return;
    }
    store
      .load(file)
      .then(() => onLoadRef.current?.(store))
      .catch((err: unknown) =>
        onErrorRef.current?.(err instanceof Error ? err : new Error(String(err))),
      );
  }, [store, file]);

  return (
    <PresentationContext.Provider value={store}>
      {renderElement(
        "div",
        { render, className, style },
        {
          state: { file },
          props: { ...elementProps, children },
        },
      )}
    </PresentationContext.Provider>
  );
}
