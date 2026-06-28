"use client";

import { Presentation as Primitive } from "@diceui/pptx";
import type {
  ThumbnailListProps,
  ThumbnailListRenderState,
  ThumbnailItemProps,
  ThumbnailItemPreviewProps,
  ThumbnailItemNumberProps,
  LoadingProps,
  PresentationErrorProps,
} from "@diceui/pptx";

import { cn } from "@/lib/utils";

function Presentation(props: React.ComponentProps<typeof Primitive.Root>) {
  return <Primitive.Root data-slot="presentation-root" {...props} />;
}

function PresentationViewport(props: React.ComponentProps<typeof Primitive.Viewport>) {
  return <Primitive.Viewport data-slot="presentation-viewport" {...props} />;
}

function PresentationSlide(props: React.ComponentProps<typeof Primitive.Slide>) {
  return <Primitive.Slide data-slot="presentation-slide" {...props} />;
}

function PresentationLoading({ className, children, ...props }: LoadingProps) {
  return (
    <Primitive.Loading
      data-slot="presentation-loading"
      className={cn(
        "absolute inset-0 z-10 flex items-center justify-center bg-background/80 text-sm text-muted-foreground",
        className,
      )}
      {...props}
    >
      {children ?? ((progress: number) => <span>Loading… {progress}%</span>)}
    </Primitive.Loading>
  );
}

function PresentationError({ className, children, ...props }: PresentationErrorProps) {
  return (
    <Primitive.Error
      data-slot="presentation-error"
      className={cn(
        "absolute inset-0 z-10 flex items-center justify-center bg-background/80 text-sm text-destructive",
        className,
      )}
      {...props}
    >
      {children ?? ((err: Error) => <span>{err.message}</span>)}
    </Primitive.Error>
  );
}

function PresentationThumbnailItemPreview({ className, ...props }: ThumbnailItemPreviewProps) {
  return (
    <Primitive.ThumbnailItemPreview
      data-slot="presentation-thumbnail-item-preview"
      className={cn("w-full overflow-hidden rounded-sm", className)}
      {...props}
    />
  );
}

function PresentationThumbnailItemNumber({ className, ...props }: ThumbnailItemNumberProps) {
  return (
    <Primitive.ThumbnailItemNumber
      data-slot="presentation-thumbnail-item-number"
      className={cn(
        "mt-1 block text-center font-mono text-[10px] leading-none text-muted-foreground",
        className,
      )}
      {...props}
    />
  );
}

function PresentationThumbnailItem({ className, children, ...props }: ThumbnailItemProps) {
  return (
    <Primitive.ThumbnailItem
      data-slot="presentation-thumbnail-item"
      className={cn(
        "relative w-full cursor-pointer rounded-md p-1 outline-none",
        "ring-1 ring-transparent transition-all duration-100",
        "hover:bg-accent hover:ring-border",
        "focus-visible:ring-ring",
        "data-active:bg-accent data-active:ring-primary",
        className,
      )}
      {...props}
    >
      {children ?? (
        <>
          <PresentationThumbnailItemPreview />
          <PresentationThumbnailItemNumber />
        </>
      )}
    </Primitive.ThumbnailItem>
  );
}

const defaultThumbnailChildren = ({ slides }: ThumbnailListRenderState) =>
  slides.map((slide) => <PresentationThumbnailItem key={slide.id} slideId={slide.id} />);

function PresentationThumbnailList({ className, style, children, ...props }: ThumbnailListProps) {
  return (
    <Primitive.ThumbnailList
      data-slot="presentation-thumbnail-list"
      className={cn("flex w-40 shrink-0 flex-col gap-1 overflow-y-auto p-2", className)}
      style={style}
      {...props}
    >
      {children ?? defaultThumbnailChildren}
    </Primitive.ThumbnailList>
  );
}

export {
  Presentation,
  PresentationViewport,
  PresentationSlide,
  PresentationLoading,
  PresentationError,
  PresentationThumbnailList,
  PresentationThumbnailItem,
  PresentationThumbnailItemPreview,
  PresentationThumbnailItemNumber,
};
