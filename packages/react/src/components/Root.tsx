import React from "react";
import { PresentationStore } from "../store";
import type { PreviewInput } from "../store";
import { PresentationContext } from "../context";

export interface RootProps {
  file: PreviewInput | null | undefined;
  children: React.ReactNode;
  onLoad?: (store: PresentationStore) => void;
  onError?: (error: Error) => void;
}

export function Root({ file, children, onLoad, onError }: RootProps) {
  const store = React.useMemo(() => new PresentationStore(), []);
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

  return <PresentationContext.Provider value={store}>{children}</PresentationContext.Provider>;
}
