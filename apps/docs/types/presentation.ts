// Types for AutoTypeTable — exposes component-specific props only,
// with inherited HTML element props stripped out.
//
// Usage in MDX:
//   <AutoTypeTable path="types/presentation.ts" name="RootOwnProps" />

import type * as React from "react";

import type { Presentation as PresentationPrimitive } from "@diceui/pptx";

export type RootProps = Omit<PresentationPrimitive.Root.Props, keyof React.ComponentProps<"div">>;

export type ViewportProps = Omit<
  PresentationPrimitive.Viewport.Props,
  keyof React.ComponentProps<"div">
>;

export type SlideProps = Omit<PresentationPrimitive.Slide.Props, keyof React.ComponentProps<"div">>;

export type SelectionProps = Omit<
  PresentationPrimitive.Selection.Props,
  keyof React.ComponentProps<"div">
>;

export type ThumbnailListProps = Omit<
  PresentationPrimitive.ThumbnailList.Props,
  keyof React.ComponentProps<"div">
>;

export type ThumbnailItemProps = Omit<
  PresentationPrimitive.ThumbnailItem.Props,
  keyof React.ComponentProps<"button">
>;

export type ThumbnailItemPreviewProps = Omit<
  PresentationPrimitive.ThumbnailItemPreview.Props,
  keyof React.ComponentProps<"div">
>;

export type ThumbnailItemNumberProps = Omit<
  PresentationPrimitive.ThumbnailItemNumber.Props,
  keyof React.ComponentProps<"span">
>;

export type LoadingProps = Omit<
  PresentationPrimitive.Loading.Props,
  keyof React.ComponentProps<"div">
>;

export type ErrorProps = Omit<PresentationPrimitive.Error.Props, keyof React.ComponentProps<"div">>;
