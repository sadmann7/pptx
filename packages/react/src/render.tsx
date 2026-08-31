import * as React from "react";

const COMPONENT_IDENTIFIER_PATTERN = /^[A-Z][A-Za-z0-9$]*$/;
const LOWERCASE_CHARACTER_PATTERN = /[a-z]/;

type AnyProps = Record<string, unknown>;
type IntrinsicTag = keyof React.JSX.IntrinsicElements;

export type ComponentRenderFn<P, S> = (props: P, state: S) => React.ReactElement;

export type RenderFunctionProps<Tag extends IntrinsicTag | undefined = undefined> =
  Tag extends IntrinsicTag ? React.JSX.IntrinsicElements[Tag] : React.HTMLAttributes<any>;

export type RenderProp<S = Record<string, never>, P = React.HTMLAttributes<any>> =
  | React.ReactElement
  | ComponentRenderFn<P, S>;

export type PrimitiveProps<
  Tag extends React.ElementType,
  State = Record<string, never>,
  RenderFnProps = React.ComponentPropsWithRef<Tag>,
> = React.ComponentPropsWithRef<Tag> & {
  /**
   * Replaces the rendered element with a different element, or composes it with another component.
   *
   * - **ReactElement**: cloned with composed props
   * - **Function**: called with composed props and state, and the default element is replaced with the return value.
   */
  render?: RenderProp<State, RenderFnProps>;
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

export interface RenderElementComponentProps<S, P = AnyProps> {
  render?: RenderProp<S, P>;
  className?: string;
  style?: React.CSSProperties;
}

export interface RenderElementParams<S, E extends Element, Tag extends IntrinsicTag> {
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
  props?:
    | (RenderFunctionProps<Tag> & AnyProps)
    | Array<(RenderFunctionProps<Tag> & AnyProps) | undefined>;
}

/**
 * Renders a component element using the Base UI render-prop pattern.
 *
 * - **No `render` prop** → renders `defaultTag` with the composed props.
 * - **`render` is a function** → calls `render(composedProps, state)`.
 * - **`render` is a `ReactElement`** → clones it, merging composed props
 *   (`className`/`style`/events composed; everything else external wins).
 *
 * A plain function rather than a hook so it can run after a conditional return.
 *
 * @see https://github.com/mui/base-ui/blob/master/packages/react/src/internals/useRenderElement.tsx
 */
export function renderElement<Tag extends IntrinsicTag, S, E extends Element = Element>(
  defaultTag: Tag,
  componentProps: RenderElementComponentProps<S, RenderFunctionProps<Tag>>,
  params: RenderElementParams<S, E, Tag>,
): React.ReactElement {
  const outProps = composeRenderProps(componentProps, params);
  return evaluateRenderProp(defaultTag, componentProps.render, outProps, params.state);
}

function composeRenderProps<Tag extends IntrinsicTag, S, E extends Element>(
  componentProps: RenderElementComponentProps<S, RenderFunctionProps<Tag>>,
  params: RenderElementParams<S, E, Tag>,
): AnyProps {
  const { render, className, style } = componentProps;
  const { ref, props } = params;

  const outProps = Array.isArray(props)
    ? props.reduce<AnyProps>((acc, p) => (p ? mergeProps(acc, p as AnyProps) : acc), {})
    : { ...(props as AnyProps | undefined) };

  if (className !== undefined) {
    outProps.className = [outProps.className, className].filter(Boolean).join(" ") || undefined;
  }
  if (style !== undefined) {
    outProps.style = { ...(outProps.style as object | undefined), ...style };
  }

  const elementRef = React.isValidElement(render) ? getElementRef(render) : null;
  const paramRefs = Array.isArray(ref) ? ref : [ref];
  if (outProps.ref != null || elementRef != null || paramRefs.some((r) => r != null)) {
    outProps.ref = mergeRefs(
      outProps.ref as React.Ref<E> | undefined,
      elementRef as React.Ref<E> | null,
      ...paramRefs,
    );
  }

  return outProps;
}

function evaluateRenderProp<Tag extends IntrinsicTag, S>(
  defaultTag: Tag,
  render: RenderProp<S, RenderFunctionProps<Tag>> | undefined,
  props: AnyProps,
  state: S,
): React.ReactElement {
  if (typeof render === "function") {
    if (process.env.NODE_ENV !== "production") {
      warnIfRenderLooksLikeComponent(render);
    }
    return render(props as unknown as RenderFunctionProps<Tag>, state);
  }

  if (render != null) {
    if (process.env.NODE_ENV !== "production" && !React.isValidElement(render)) {
      throw new Error("`render` must be a valid React element or a function.");
    }
    const merged = mergeProps(props, render.props as AnyProps);
    merged.ref = props.ref;
    return React.cloneElement(render, merged);
  }

  return renderTag(defaultTag, props);
}

function getElementRef(element: React.ReactElement): React.Ref<unknown> | null {
  return (element.props as { ref?: React.Ref<unknown> }).ref ?? null;
}

function renderTag(tag: string, props: AnyProps): React.ReactElement {
  if (tag === "button") {
    return <button type="button" {...props} />;
  }
  if (tag === "img") {
    return <img alt="" {...props} />;
  }
  return React.createElement(tag, props);
}

function warnIfRenderLooksLikeComponent(renderFn: { name: string }) {
  const { name } = renderFn;
  if (name.length === 0) return;
  if (!COMPONENT_IDENTIFIER_PATTERN.test(name)) return;
  if (!LOWERCASE_CHARACTER_PATTERN.test(name)) return;

  console.warn(
    [
      `The \`render\` prop received a function named \`${name}\` that starts with an uppercase letter.`,
      "This usually means a React component was passed directly as `render={Component}`.",
      "The function is called as a plain callback, which can break the Rules of Hooks.",
      "If this is an intentional render callback, rename it to start with a lowercase letter.",
      "Use `render={<Component />}` or `render={(props) => <Component {...props} />}` instead.",
    ].join("\n"),
  );
}
