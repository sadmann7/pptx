import * as React from "react";
import { usePresentation } from "../context";
import { renderElement } from "../utils/render";
import type { RenderProp } from "../utils/render";

export interface ErrorState {
  /** The error thrown during parsing. */
  error: Error;
}

export interface PresentationErrorProps extends Omit<React.ComponentProps<"div">, "children"> {
  /**
   * Rendered when the presentation is in the `"error"` state.
   * Pass a function to receive the `Error` instance.
   */
  children?: React.ReactNode | ((error: Error) => React.ReactNode);
  /**
   * Replace the wrapper element.
   * - ReactElement: cloned with composed props
   * - Function: `(props, state) => ReactElement`
   */
  render?: RenderProp<ErrorState>;
}

/**
 * Renders its children only when the presentation failed to parse.
 * Place anywhere inside `<Presentation.Root>`.
 *
 * @example
 * <Presentation.Error>
 *   {(err) => <span className="text-destructive">{err.message}</span>}
 * </Presentation.Error>
 */
export const PresentationError = React.forwardRef<HTMLDivElement, PresentationErrorProps>(
  function PresentationError(
    { children, className, style, render, ...elementProps },
    forwardedRef,
  ) {
    const { status, error } = usePresentation();

    if (status !== "error" || !error) return null;

    const resolvedChildren = typeof children === "function" ? children(error) : children;

    return renderElement(
      "div",
      { render, className, style },
      {
        state: { error },
        ref: forwardedRef,
        props: { ...elementProps, children: resolvedChildren },
      },
    );
  },
);
