import * as React from "react";

import { usePresentation } from "./context";
import type { RenderProp } from "./render";
import { renderElement } from "./render";

export interface LoadingState {
  /** 0–100 parse progress reported by the store. */
  progress: number;
}

export interface LoadingProps extends Omit<React.ComponentProps<"div">, "children"> {
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

  /**
   * Replace the wrapper element.
   * - ReactElement: cloned with composed props
   * - Function: `(props, state) => ReactElement`
   */
  render?: RenderProp<LoadingState>;
}

/**
 * Renders its children only while the presentation is in the `"loading"` state.
 * Place anywhere inside `<Presentation.Root>`.
 */
export const Loading = React.forwardRef<HTMLDivElement, LoadingProps>(function Loading(
  { children, className, style, render, ...loadingProps },
  forwardedRef,
) {
  const { status, progress } = usePresentation();

  if (status !== "loading") return null;

  const resolvedChildren = typeof children === "function" ? children(progress) : children;

  return renderElement(
    "div",
    { render, className, style },
    {
      state: { progress },
      ref: forwardedRef,
      props: { ...loadingProps, children: resolvedChildren },
    },
  );
});

export namespace Loading {
  export type State = LoadingState;
  export type Props = LoadingProps;
}
