import * as React from "react";

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/**
 * A render callback that receives composed props and component state,
 * returning the element to render — base-ui style.
 */
export type ComponentRenderFn<P, S> = (props: P, state: S) => React.ReactElement;

/**
 * Render prop accepted by every component.
 * - `ReactElement`: cloned with composed props (className/style merged, events chained)
 * - `Function`: called with `(composedProps, state)` — full control
 */
export type RenderProp<S = Record<string, never>> =
  | React.ReactElement
  | ComponentRenderFn<React.HTMLAttributes<any>, S>;

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

type AnyProps = Record<string, unknown>;

/**
 * Merges two prop objects:
 * - `className`: joined with a space
 * - `style`: shallowly merged (b wins per-key)
 * - `on*` handlers: chained (a runs first)
 * - Everything else: b overwrites a
 */
export function mergeProps(a: AnyProps, b: AnyProps): AnyProps {
  const result: AnyProps = { ...a };
  for (const key of Object.keys(b)) {
    const av = a[key];
    const bv = b[key];
    if (key === "className") {
      const joined = [av, bv].filter(Boolean).join(" ");
      result[key] = joined || undefined;
    } else if (
      key === "style" &&
      av != null &&
      bv != null &&
      typeof av === "object" &&
      typeof bv === "object"
    ) {
      result[key] = { ...(av as object), ...(bv as object) };
    } else if (key.startsWith("on") && typeof av === "function" && typeof bv === "function") {
      const fa = av as (...args: unknown[]) => void;
      const fb = bv as (...args: unknown[]) => void;
      result[key] = (...args: unknown[]) => {
        fa(...args);
        fb(...args);
      };
    } else {
      result[key] = bv;
    }
  }
  return result;
}

// ---------------------------------------------------------------------------
// mergeRefs
// ---------------------------------------------------------------------------

/**
 * Combines multiple refs into a single callback ref.
 */
export function mergeRefs<T>(...refs: (React.Ref<T> | null | undefined)[]): React.RefCallback<T> {
  return (value) => {
    for (const ref of refs) {
      if (typeof ref === "function") {
        ref(value);
      } else if (ref != null) {
        (ref as React.RefObject<T | null>).current = value;
      }
    }
  };
}

// ---------------------------------------------------------------------------
// useRenderElement
// ---------------------------------------------------------------------------

/**
 * Props the component receives from outside that drive element customisation.
 * Pass the full component props object — only `render`, `className`, and `style`
 * are consumed; everything else is ignored here.
 */
export interface UseRenderElementComponentProps<S> {
  render?: RenderProp<S>;
  className?: string;
  style?: React.CSSProperties;
}

/**
 * Internal parameters that control how the element is built.
 */
export interface UseRenderElementParams<S, E extends Element> {
  /** Component state forwarded to the render callback. */
  state: S;
  /**
   * Ref(s) to attach to the rendered element.
   * Pass an array when multiple refs need to be merged (e.g. `[internalRef, forwardedRef]`).
   */
  ref?: React.Ref<E> | (React.Ref<E> | null | undefined)[];
  /**
   * Internal default props for the element.
   * May be a single object or an ordered array of objects — merged left-to-right.
   * User-supplied `className` and `style` (from `componentProps`) are composed on top.
   */
  props?: AnyProps | (AnyProps | undefined)[];
}

/**
 * Renders a Base-UI-style element.
 *
 * - No `render`: renders `defaultTag` with composed props
 * - `render` is a function: calls `render(composedProps, state)`
 * - `render` is a ReactElement: clones it, merging composed props into its own
 *
 * `className` and `style` from `componentProps` are always merged on top of any
 * matching keys in `params.props` (user wins per-key for style, user appended for className).
 */
export function useRenderElement<S, E extends Element = Element>(
  defaultTag: keyof React.JSX.IntrinsicElements,
  componentProps: UseRenderElementComponentProps<S>,
  params: UseRenderElementParams<S, E>,
): React.ReactElement {
  const { render, className: userClassName, style: userStyle } = componentProps;
  const { state, ref, props } = params;

  // Flatten props array into a single object
  const internalProps: AnyProps = Array.isArray(props)
    ? props.reduce<AnyProps>((acc, p) => (p ? mergeProps(acc, p) : acc), {})
    : (props ?? {});

  // className: internal first, user appended
  const internalClassName = internalProps.className as string | undefined;
  const mergedClassName = [internalClassName, userClassName].filter(Boolean).join(" ") || undefined;

  // style: internal defaults, user wins per-key
  const internalStyle = internalProps.style as React.CSSProperties | undefined;
  const mergedStyle = internalStyle || userStyle ? { ...internalStyle, ...userStyle } : undefined;

  const finalProps: AnyProps = {
    ...internalProps,
    ...(mergedClassName !== undefined && { className: mergedClassName }),
    ...(mergedStyle !== undefined && { style: mergedStyle }),
  };

  // Attach merged ref
  if (ref !== undefined) {
    finalProps.ref = Array.isArray(ref)
      ? mergeRefs(...(ref as (React.Ref<E> | null | undefined)[]))
      : ref;
  }

  // Evaluate render prop
  if (render !== undefined) {
    if (typeof render === "function") {
      return render(finalProps as React.HTMLAttributes<any>, state);
    }
    const merged = mergeProps(finalProps, render.props as AnyProps);
    merged.ref = finalProps.ref;
    return React.cloneElement(render, merged);
  }

  const Tag = defaultTag as React.ElementType;
  return <Tag {...finalProps} />;
}
