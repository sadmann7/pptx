import * as React from "react";

import { DEFAULT_STORE_STATE } from "./constant";
import { useStoreContext, useStoreSelector } from "./context";
import type { PrimitiveProps } from "./render";
import { renderElement } from "./render";

export interface LoadingState {
  /** 0–100 parse progress reported by the store. */
  progress: number;
}

export interface LoadingProps extends Omit<
  PrimitiveProps<"div", LoadingState>,
  "children"
> {
  /**
   * Rendered while the presentation is loading.
   * Pass a function to receive the current `progress` (0–100).
   *
   * ```tsx
   * <Presentation.Loading>
   *   {(progress) => <span>Loading {progress}%</span>}
   * </Presentation.Loading>
   * ```
   */
  children?: React.ReactNode | ((progress: number) => React.ReactNode);
}

export const Loading = React.forwardRef<HTMLDivElement, LoadingProps>(
  function Loading({ children, render, ...loadingProps }, forwardedRef) {
    const store = useStoreContext("Presentation.Loading");
    const status = useStoreSelector(
      store,
      (s) => s.status,
      DEFAULT_STORE_STATE.status,
    );
    const progress = useStoreSelector(
      store,
      (s) => s.progress,
      DEFAULT_STORE_STATE.progress,
    );

    if (status !== "loading") return null;

    const resolvedChildren =
      typeof children === "function" ? children(progress) : children;

    return renderElement(
      "div",
      { render },
      {
        state: { progress },
        ref: forwardedRef,
        props: [
          { role: "status", "aria-live": "polite", children: resolvedChildren },
          loadingProps,
        ],
      },
    );
  },
);

export namespace Loading {
  export type State = LoadingState;
  export type Props = LoadingProps;
}
