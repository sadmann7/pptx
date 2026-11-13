import * as React from "react";

import { DEFAULT_STORE_STATE } from "./constant";
import { useStoreContext, useStoreSelector } from "./context";
import type { PrimitiveProps } from "./render";
import { renderElement } from "./render";

type GlobalError = globalThis.Error;

export interface ErrorState {
  /** The error thrown during parsing. */
  error: GlobalError;
}

export interface ErrorProps extends Omit<
  PrimitiveProps<"div", ErrorState>,
  "children"
> {
  /**
   * Rendered when the presentation is in the `"error"` state.
   * Pass a function to receive the `Error` instance.
   *
   * ```tsx
   * <Presentation.Error>
   *   {(err) => <span>{err.message}</span>}
   * </Presentation.Error>
   * ```
   */
  children?: React.ReactNode | ((error: GlobalError) => React.ReactNode);
}

export const Error = React.forwardRef<HTMLDivElement, ErrorProps>(
  function Error({ children, render, ...errorProps }, forwardedRef) {
    const store = useStoreContext("Presentation.Error");
    const status = useStoreSelector(
      store,
      (s) => s.status,
      DEFAULT_STORE_STATE.status,
    );
    const error = useStoreSelector(store, (s) => s.error, null);

    if (status !== "error" || !error) return null;

    const resolvedChildren =
      typeof children === "function" ? children(error) : children;

    return renderElement(
      "div",
      { render },
      {
        state: { error },
        ref: forwardedRef,
        props: [{ role: "alert", children: resolvedChildren }, errorProps],
      },
    );
  },
);

export namespace Error {
  export type State = ErrorState;
  export type Props = ErrorProps;
}
