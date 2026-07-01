/**
 * Render utilities — pattern adapted from Base UI.
 *
 * The render-prop model, prop-merging strategy, and `useRenderElement` API are
 * directly inspired by the Base UI component library by MUI.
 *
 * @see {@link https://github.com/mui/base-ui Base UI source}
 * @see {@link https://base-ui.com/react/utils/use-render useRender hook docs}
 * @license MIT — Base UI © MUI
 */

import * as React from "react";

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/**
 * A render callback that receives composed props and component state,
 * returning the element to render.
 *
 * Mirrors {@link https://github.com/mui/base-ui/blob/master/packages/react/src/internals/types.ts `ComponentRenderFn`}
 * from Base UI.
 */
export type ComponentRenderFn<P, S> = (props: P, state: S) => React.ReactElement;

/**
 * Render prop accepted by every component.
 *
 * - `ReactElement` — cloned with composed props (`className`/`style` merged, event handlers chained)
 * - `Function` — called with `(composedProps, state)` for full control
 *
 * Mirrors the `render` prop pattern from Base UI.
 * @see {@link https://base-ui.com/react/utils/use-render#render Base UI render prop}
 */
export type RenderProp<S = Record<string, never>> =
  | React.ReactElement
  | ComponentRenderFn<React.HTMLAttributes<any>, S>;

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

type AnyProps = Record<string, unknown>;

/**
 * Merges two prop objects with smart composition:
 *
 * - `className` — joined with a space (both preserved)
 * - `style` — shallowly merged, `b` wins per-key
 * - `on*` event handlers — chained (`a` runs first, then `b`)
 * - Everything else — `b` overwrites `a`
 *
 * Adapted from
 * {@link https://github.com/mui/base-ui/blob/master/packages/react/src/merge-props/mergeProps.ts `mergeProps`}
 * in Base UI.
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
 *
 * Mirrors
 * {@link https://github.com/mui/base-ui/blob/master/packages/utils/src/useMergedRefs.ts `useMergedRefs`}
 * from Base UI (simplified — no hook, no SSR guard needed here).
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
// renderElement
// ---------------------------------------------------------------------------

/**
 * The subset of component props that drive element customisation.
 * Pass the full component props object — only `render`, `className`, and `style`
 * are consumed here; everything else is ignored.
 *
 * Mirrors
 * {@link https://github.com/mui/base-ui/blob/master/packages/react/src/internals/useRenderElement.ts `UseRenderElementComponentProps`}
 * from Base UI.
 */
export interface RenderElementComponentProps<S> {
  render?: RenderProp<S>;
  className?: string;
  style?: React.CSSProperties;
}

/**
 * Internal parameters that control how the element is built.
 *
 * Mirrors
 * {@link https://github.com/mui/base-ui/blob/master/packages/react/src/internals/useRenderElement.ts `UseRenderElementParameters`}
 * from Base UI.
 */
export interface RenderElementParams<S, E extends Element> {
  /** Component state forwarded as the second argument to a function `render` prop. */
  state: S;
  /**
   * Ref(s) to attach to the rendered element.
   * Pass an array to merge multiple refs (e.g. `[internalRef, forwardedRef]`).
   */
  ref?: React.Ref<E> | (React.Ref<E> | null | undefined)[];
  /**
   * Internal default props for the element.
   * May be a single object or an ordered array — merged left-to-right.
   * User-supplied `className` and `style` from `componentProps` are always
   * composed on top (user wins per-key for `style`; appended for `className`).
   */
  props?: AnyProps | (AnyProps | undefined)[];
}

/**
 * Renders a component element using the Base UI render-prop pattern.
 *
 * - **No `render` prop** → renders `defaultTag` with the composed props.
 * - **`render` is a function** → calls `render(composedProps, state)`.
 * - **`render` is a `ReactElement`** → clones it, merging composed props
 *   (`className`/`style`/events composed; everything else external wins).
 *
 * `className` and `style` from `componentProps` are merged on top of any
 * matching keys in `params.props` so that internal defaults never override
 * user-supplied values.
 *
 * Adapted from
 * {@link https://github.com/mui/base-ui/blob/master/packages/react/src/internals/useRenderElement.ts `useRenderElement`}
 * in Base UI.
 *
 * @see {@link https://base-ui.com/react/utils/use-render Base UI useRender docs}
 */
export function renderElement<S, E extends Element = Element>(
  defaultTag: keyof React.JSX.IntrinsicElements,
  componentProps: RenderElementComponentProps<S>,
  params: RenderElementParams<S, E>,
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
