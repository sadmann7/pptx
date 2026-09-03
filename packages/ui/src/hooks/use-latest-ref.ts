import * as React from "react";

function useLatestRef<T>(value: T): React.RefObject<T> {
  const ref = React.useRef<T>(value);
  React.useLayoutEffect(() => {
    ref.current = value;
  });
  return ref as React.RefObject<T>;
}

export { useLatestRef };
