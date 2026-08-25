// Types for AutoTypeTable: exposes component-specific props only,
// with inherited HTML element props stripped out.
//
// `Omit<Props, keyof ComponentProps<"div">>` also drops names we override
// (`onLoad`, `onError`, `onSelect`, render-function `children`) because
// those keys exist on the host element. They get picked back.
//
// Usage in MDX:
//   <AutoTypeTable path="types/presentation.ts" name="RootProps" />

import type * as React from "react";

import type { Presentation as PresentationPrimitive } from "@diceui/pptx";

type OverriddenHtmlKeys<T, Html> = {
  [K in keyof T & keyof Html]-?: [T[K]] extends [Html[K]] ? never : K;
}[keyof T & keyof Html];

type IndieProps<T, Html> = Omit<T, keyof Html> & Pick<T, OverriddenHtmlKeys<T, Html>>;

type DivProps = React.ComponentProps<"div">;
type ButtonProps = React.ComponentProps<"button">;
type SpanProps = React.ComponentProps<"span">;

export type ProviderProps = PresentationPrimitive.Provider.Props;

export type RootProps = IndieProps<PresentationPrimitive.Root.Props, DivProps>;

export type ViewportProps = IndieProps<PresentationPrimitive.Viewport.Props, DivProps>;

export type SlideProps = IndieProps<PresentationPrimitive.Slide.Props, DivProps>;

export type SelectionProps = IndieProps<PresentationPrimitive.Selection.Props, DivProps>;

export type ThumbnailListProps = IndieProps<PresentationPrimitive.ThumbnailList.Props, DivProps>;

export type ThumbnailItemProps = IndieProps<PresentationPrimitive.ThumbnailItem.Props, ButtonProps>;

export type ThumbnailItemPreviewProps = IndieProps<
  PresentationPrimitive.ThumbnailItemPreview.Props,
  DivProps
>;

export type ThumbnailItemNumberProps = IndieProps<
  PresentationPrimitive.ThumbnailItemNumber.Props,
  SpanProps
>;

export type LoadingProps = IndieProps<PresentationPrimitive.Loading.Props, DivProps>;

export type ErrorProps = IndieProps<PresentationPrimitive.Error.Props, DivProps>;
