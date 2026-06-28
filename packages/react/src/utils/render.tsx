import * as React from "react";

/**
 * A render function that receives composed props and component state,
 * returning the element to render — base-ui style.
 */
export type ComponentRenderFn<P, S> = (props: P, state: S) => React.ReactElement;

/**
 * Render prop accepted by every component.
 * - ReactElement: cloned with composed props (className/style merged, events chained)
 * - Function: called with (composedProps, state) — full control
 */
export type RenderProp<P = React.HTMLAttributes<HTMLElement>, S = Record<string, never>> =
  | React.ReactElement
  | ComponentRenderFn<P, S>;

// Internal alias for prop bags with unknown shape
type AnyProps = Record<string, unknown>;

/**
 * Merges two prop objects with smart handling:
 * - `className`: joined with a space
 * - `style`: shallowly merged (external wins per-key)
 * - `on*` handlers: chained (internal runs first)
 * - Everything else: external overwrites internal
 */
export function mergeProps(internal: AnyProps, external: AnyProps): AnyProps {
  const result: AnyProps = { ...internal };
  for (const key of Object.keys(external)) {
    const a = internal[key];
    const b = external[key];
    if (key === "className") {
      const joined = [a, b].filter(Boolean).join(" ");
      result[key] = joined || undefined;
    } else if (
      key === "style" &&
      a != null &&
      b != null &&
      typeof a === "object" &&
      typeof b === "object"
    ) {
      result[key] = { ...(a as object), ...(b as object) };
    } else if (key.startsWith("on") && typeof a === "function" && typeof b === "function") {
      const fa = a as (...args: unknown[]) => void;
      const fb = b as (...args: unknown[]) => void;
      result[key] = (...args: unknown[]) => {
        fa(...args);
        fb(...args);
      };
    } else {
      result[key] = b;
    }
  }
  return result;
}

/**
 * Combines multiple refs into a single callback ref.
 * Accepts React.Refs (callback, object, or null/undefined).
 */
export function mergeRefs<T>(...refs: (React.Ref<T> | null | undefined)[]): React.RefCallback<T> {
  return (value) => {
    for (const ref of refs) {
      if (typeof ref === "function") {
        ref(value);
      } else if (ref != null) {
        (ref as React.MutableRefObject<T | null>).current = value;
      }
    }
  };
}

/**
 * Renders an element using the base-ui render prop pattern.
 *
 * - No `render`: renders `defaultTag` with `props`
 * - `render` is a function: calls it with `(props, state)`
 * - `render` is a ReactElement: clones it, merging `props` into its own
 *   (className/style/events composed, everything else external wins)
 */
export function renderElement<S>(
  defaultTag: keyof React.JSX.IntrinsicElements,
  render: RenderProp<AnyProps, S> | undefined,
  props: AnyProps,
  state: S,
): React.ReactElement {
  if (render === undefined) {
    const Tag = defaultTag as React.ElementType;
    return <Tag {...props} />;
  }
  if (typeof render === "function") {
    return render(props, state);
  }
  // ReactElement: internal props are the base; element's own props overlay on top
  const merged = mergeProps(props, render.props as AnyProps);
  return React.cloneElement(render, merged);
}
