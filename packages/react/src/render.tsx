import * as React from "react";

type AnyProps = Record<string, unknown>;

export type ComponentRenderFn<P, S> = (props: P, state: S) => React.ReactElement;

export type RenderProp<S = Record<string, never>, P = React.HTMLAttributes<any>> =
  | React.ReactElement
  | ComponentRenderFn<P, S>;

export type PrimitiveProps<
  Tag extends React.ElementType,
  State = Record<string, never>,
  RenderFunctionProps = React.ComponentProps<Tag>,
> = React.ComponentProps<Tag> & {
  /**
   * Replaces the rendered element with a different element, or composes it with another component.
   *
   * - **ReactElement**: cloned with composed props
   * - **Function**: called with composed props and state, and the default element is replaced with the return value.
   * ```
   */
  render?: RenderProp<State, RenderFunctionProps>;
};

/**
 * Merges two prop objects with smart composition:
 *
 * - `className`: joined with a space (both preserved)
 * - `style`: shallowly merged, `b` wins per-key
 * - `on*` event handlers: chained (`a` runs first, then `b`)
 * - Everything else: `b` overwrites `a`
 *
 * @see https://github.com/mui/base-ui/blob/master/packages/react/src/merge-props/mergeProps.ts
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

/**
 * Merges multiple refs into a single callback ref.
 *
 * @see https://github.com/mui/base-ui/blob/master/packages/utils/src/useMergedRefs.ts
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

export interface RenderElementComponentProps<S> {
  render?: RenderProp<S>;
  className?: string;
  style?: React.CSSProperties;
}

export interface RenderElementParams<S, E extends Element, P extends AnyProps = AnyProps> {
  /** Component state forwarded as the second argument to a function `render` prop. */
  state: S;
  /**
   * Ref(s) to attach to the rendered element.
   * Pass an array to merge multiple refs (e.g. `[internalRef, forwardedRef]`).
   */
  ref?: React.Ref<E> | (React.Ref<E> | null | undefined)[];
  /**
   * Internal default props for the element.
   * May be a single object or an ordered array: merged left-to-right.
   * User-supplied `className` and `style` from `componentProps` are always
   * composed on top (user wins per-key for `style`; appended for `className`).
   */
  props?: P | (P | undefined)[];
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
 * @see https://github.com/mui/base-ui/blob/master/packages/react/src/internals/useRenderElement.tsx
 */
export function renderElement<
  Tag extends keyof React.JSX.IntrinsicElements,
  S,
  E extends Element = Element,
>(
  defaultTag: Tag,
  componentProps: RenderElementComponentProps<S>,
  params: RenderElementParams<S, E, React.JSX.IntrinsicElements[Tag] & AnyProps>,
): React.ReactElement {
  const { render, className: userClassName, style: userStyle } = componentProps;
  const { state, ref, props } = params;

  const internalProps: AnyProps = Array.isArray(props)
    ? props.reduce<AnyProps>((acc, p) => (p ? mergeProps(acc, p) : acc), {})
    : (props ?? {});

  const internalClassName = internalProps.className as string | undefined;
  const mergedClassName = [internalClassName, userClassName].filter(Boolean).join(" ") || undefined;

  const internalStyle = internalProps.style as React.CSSProperties | undefined;
  const mergedStyle = internalStyle || userStyle ? { ...internalStyle, ...userStyle } : undefined;

  const finalProps: AnyProps = {
    ...internalProps,
    ...(mergedClassName !== undefined && { className: mergedClassName }),
    ...(mergedStyle !== undefined && { style: mergedStyle }),
  };

  if (ref !== undefined) {
    finalProps.ref = Array.isArray(ref)
      ? mergeRefs(...(ref as (React.Ref<E> | null | undefined)[]))
      : ref;
  }

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
