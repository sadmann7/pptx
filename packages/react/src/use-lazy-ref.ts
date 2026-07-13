import * as React from "react";

const UNINITIALIZED: unique symbol = Symbol("useLazyRef.uninitialized");

function useLazyRef<T>(fn: () => T): React.RefObject<T> {
  const ref = React.useRef<T | typeof UNINITIALIZED>(UNINITIALIZED);

  if (ref.current === UNINITIALIZED) {
    ref.current = fn();
  }

  return ref as React.RefObject<T>;
}

export { useLazyRef };
