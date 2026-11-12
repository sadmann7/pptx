// Props for AutoTypeTable without inherited HTML element props. Omitting
// `keyof ComponentProps<"div">` also drops overridden names like `onLoad` and
// `children`, so they're picked back.
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

export type SelectionBoxProps = IndieProps<PresentationPrimitive.SelectionBox.Props, DivProps>;

export type ResizeHandleProps = IndieProps<PresentationPrimitive.ResizeHandle.Props, DivProps>;

export type RotateHandleProps = IndieProps<PresentationPrimitive.RotateHandle.Props, DivProps>;

export type SnapGuideProps = IndieProps<PresentationPrimitive.SnapGuide.Props, DivProps>;

export type MarqueeProps = IndieProps<PresentationPrimitive.Marquee.Props, DivProps>;
export type ControlsProps = IndieProps<PresentationPrimitive.Controls.Props, DivProps>;

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
