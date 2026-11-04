import * as React from "react";

function useLazyRef<T>(fn: () => T): React.RefObject<T> {
  // A state initializer runs once, without reading a ref during render.
  const [ref] = React.useState(() => ({ current: fn() }));
  return ref;
}

function useLatestRef<T>(value: T): React.RefObject<T> {
  const ref = React.useRef<T>(value);
  React.useLayoutEffect(() => {
    ref.current = value;
  });
  return ref as React.RefObject<T>;
}

export { useLatestRef, useLazyRef };
