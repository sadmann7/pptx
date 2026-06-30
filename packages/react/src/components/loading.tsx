import * as React from "react";
import { usePresentation } from "../context";
import { renderElement } from "../utils/render";
import type { RenderProp } from "../utils/render";

export interface LoadingState {
  /** 0–100 parse progress reported by the store. */
  progress: number;
}

export interface LoadingProps extends Omit<React.ComponentProps<"div">, "children"> {
  /**
   * Rendered while the presentation is loading.
   * Pass a function to receive the current `progress` (0–100).
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
 *
 * @example
 * <Presentation.Loading>
 *   {(progress) => <span>Loading {progress}%…</span>}
 * </Presentation.Loading>
 */
export const Loading = React.forwardRef<HTMLDivElement, LoadingProps>(function Loading(
  { children, className, style, render, ...elementProps },
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
      props: { ...elementProps, children: resolvedChildren },
    },
  );
});
