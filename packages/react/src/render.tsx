import * as React from "react";

const COMPONENT_IDENTIFIER_PATTERN = /^[A-Z][A-Za-z0-9$]*$/;
const LOWERCASE_CHARACTER_PATTERN = /[a-z]/;

type AnyProps = Record<string, unknown>;
type IntrinsicTag = keyof React.JSX.IntrinsicElements;

export type ComponentRenderFn<P, S> = (
  props: P,
  state: S,
) => React.ReactElement;

export type RenderFunctionProps<
  Tag extends IntrinsicTag | undefined = undefined,
> = Tag extends IntrinsicTag
  ? React.JSX.IntrinsicElements[Tag]
  : React.HTMLAttributes<any>;

export type RenderProp<
  S = Record<string, never>,
  P = React.HTMLAttributes<any>,
> = React.ReactElement | ComponentRenderFn<P, S>;

export type PrimitiveProps<
  Tag extends React.ElementType,
  State = Record<string, never>,
  RenderFnProps = React.ComponentPropsWithRef<Tag>,
> = React.ComponentPropsWithRef<Tag> & {
  /**
   * Replaces the rendered element with a different element, or composes it with another component.
   *
   * - **ReactElement**: cloned with composed props
   * - **Function**: called with composed props and state
   */
  render?: RenderProp<State, RenderFnProps>;
};

// Follows Base UI's `mergeProps`: `b` wins, except class names join, styles merge and handlers chain.
export function mergeProps(a: AnyProps, b: AnyProps): AnyProps {
  return assignProps({ ...a }, b);
}

function assignProps(target: AnyProps, source: AnyProps): AnyProps {
  for (const key of Object.keys(source)) {
    const av = target[key];
    const bv = source[key];
    if (key === "className") {
      target[key] = joinClassNames(av, bv);
    } else if (
      key === "style" &&
      av != null &&
      bv != null &&
      typeof av === "object" &&
      typeof bv === "object"
    ) {
      target[key] = { ...(av as object), ...(bv as object) };
    } else if (
      typeof av === "function" &&
      typeof bv === "function" &&
      key.startsWith("on")
    ) {
      const fa = av as (...args: unknown[]) => void;
      const fb = bv as (...args: unknown[]) => void;
      target[key] = (...args: unknown[]) => {
        fa(...args);
        fb(...args);
      };
    } else {
      target[key] = bv;
    }
  }
  return target;
}

function joinClassNames(a: unknown, b: unknown): unknown {
  return [a, b].filter(Boolean).join(" ") || undefined;
}

interface MergedRefNode {
  merged?: React.RefCallback<unknown>;
  next?: WeakMap<object, MergedRefNode>;
}

const mergedRefs: MergedRefNode = {};

// The same refs always merge into the same callback. React detaches and re-attaches a ref
// whose identity changes, so a fresh callback per render would do that on every commit.
export function mergeRefs<T>(
  ...refs: (React.Ref<T> | null | undefined)[]
): React.RefCallback<T> {
  let node = mergedRefs;
  for (const ref of refs) {
    if (ref == null) continue;
    node.next ??= new WeakMap();
    let child = node.next.get(ref);
    if (!child) {
      child = {};
      node.next.set(ref, child);
    }
    node = child;
  }
  node.merged ??= (value) => {
    for (const ref of refs) {
      if (typeof ref === "function") {
        ref(value as T);
      } else if (ref != null) {
        (ref as React.RefObject<unknown>).current = value;
      }
    }
  };
  return node.merged as React.RefCallback<T>;
}

export type StatefulPrimitiveProps<Tag extends React.ElementType, State> = Omit<
  PrimitiveProps<Tag, State>,
  "className" | "style"
> & {
  /** Class names, or a function returning them from the component's state. */
  className?: string | ((state: State) => string | undefined);

  /** Inline styles, or a function returning them from the component's state. */
  style?:
    | React.CSSProperties
    | ((state: State) => React.CSSProperties | undefined);
};

export interface RenderElementComponentProps<S, P = AnyProps> {
  render?: RenderProp<S, P>;
  className?: string | ((state: S) => string | undefined);
  style?: React.CSSProperties | ((state: S) => React.CSSProperties | undefined);
}

export interface RenderElementParams<
  S,
  E extends Element,
  Tag extends IntrinsicTag,
> {
  state: S;
  ref?: React.Ref<E> | (React.Ref<E> | null | undefined)[];
  // Merged left to right; the user's `className` and `style` always compose on top.
  props?:
    | (RenderFunctionProps<Tag> & AnyProps)
    | Array<(RenderFunctionProps<Tag> & AnyProps) | undefined>;
}

// A plain function rather than a hook so it can run after a conditional return.
export function renderElement<
  Tag extends IntrinsicTag,
  S,
  E extends Element = Element,
>(
  defaultTag: Tag,
  componentProps: RenderElementComponentProps<S, RenderFunctionProps<Tag>>,
  params: RenderElementParams<S, E, Tag>,
): React.ReactElement {
  const outProps = composeRenderProps(componentProps, params);
  return evaluateRenderProp(
    defaultTag,
    componentProps.render,
    outProps,
    params.state,
  );
}

function composeRenderProps<Tag extends IntrinsicTag, S, E extends Element>(
  componentProps: RenderElementComponentProps<S, RenderFunctionProps<Tag>>,
  params: RenderElementParams<S, E, Tag>,
): AnyProps {
  const { render } = componentProps;
  const { ref, props, state } = params;
  const className =
    typeof componentProps.className === "function"
      ? componentProps.className(state)
      : componentProps.className;
  const style =
    typeof componentProps.style === "function"
      ? componentProps.style(state)
      : componentProps.style;

  const outProps: AnyProps = {};
  if (Array.isArray(props)) {
    for (const part of props) {
      if (part) assignProps(outProps, part as AnyProps);
    }
  } else if (props) {
    Object.assign(outProps, props);
  }

  if (className !== undefined) {
    outProps.className = joinClassNames(outProps.className, className);
  }
  if (style !== undefined) {
    outProps.style = { ...(outProps.style as object | undefined), ...style };
  }

  const elementRef = React.isValidElement(render)
    ? getElementRef(render)
    : null;
  const paramRefs = Array.isArray(ref) ? ref : [ref];
  const refs = [
    outProps.ref as React.Ref<E> | undefined,
    elementRef as React.Ref<E> | null,
    ...paramRefs,
  ].filter((r) => r != null);
  // A lone ref passes through untouched, so its identity is the caller's.
  if (refs.length === 1) outProps.ref = refs[0];
  else if (refs.length > 1) outProps.ref = mergeRefs(...refs);

  return outProps;
}

function warnIfRenderLooksLikeComponent(renderFn: { name: string }) {
  const { name } = renderFn;
  if (name.length === 0) return;
  if (!COMPONENT_IDENTIFIER_PATTERN.test(name)) return;
  if (!LOWERCASE_CHARACTER_PATTERN.test(name)) return;

  console.warn(
    [
      `The \`render\` prop received a function named \`${name}\` starting with an uppercase letter.`,
      "You likely passed a React component as `render={Component}`.",
      "This runs it as a plain callback, which may break the Rules of Hooks.",
      "Rename it to start with lowercase or use `render={<Component />}` instead.",
    ].join("\n"),
  );
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
    if (
      process.env.NODE_ENV !== "production" &&
      !React.isValidElement(render)
    ) {
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
