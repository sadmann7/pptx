"use client";

import type {
  LoadingProps,
  PresentationErrorProps,
  ThumbnailItemNumberProps,
  ThumbnailItemPreviewProps,
  ThumbnailItemProps,
  ThumbnailListProps,
} from "@diceui/pptx";
import { Presentation as Primitive } from "@diceui/pptx";

import { cn } from "@/lib/utils";

function Presentation({ className, ...props }: React.ComponentProps<typeof Primitive.Root>) {
  return (
    <Primitive.Root
      data-slot="presentation-root"
      className={cn("flex flex-col overflow-hidden", className)}
      {...props}
    />
  );
}

function PresentationBody({ className, ...props }: React.ComponentProps<"div">) {
  return (
    <div
      data-slot="presentation-body"
      className={cn("flex flex-1 overflow-hidden", className)}
      {...props}
    />
  );
}

function PresentationContent({ className, ...props }: React.ComponentProps<"div">) {
  return (
    <div
      data-slot="presentation-content"
      className={cn("relative flex flex-1 flex-col overflow-hidden", className)}
      {...props}
    />
  );
}

function PresentationViewport({
  className,
  ...props
}: React.ComponentProps<typeof Primitive.Viewport>) {
  return (
    <Primitive.Viewport
      data-slot="presentation-viewport"
      className={cn("flex-1", className)}
      {...props}
    />
  );
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
      {children ?? ((progress) => <span>Loading… {progress}%</span>)}
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
      {children ?? ((error: Error) => <span>{error.message}</span>)}
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
        "mt-1 block text-center font-mono text-xs leading-none text-muted-foreground tabular-nums",
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
        "relative w-full cursor-pointer rounded-md p-1 outline-none not-first:mt-4",
        "ring-2 ring-transparent ring-offset-2 ring-offset-background transition-all duration-100",
        "hover:bg-accent hover:ring-border",
        "focus-visible:ring-ring",
        "data-active:bg-accent data-active:ring-border",
        "group-focus-within/thumbs:data-active:ring-ring",
        className,
      )}
      {...props}
    >
      {children ?? (
        <>
          <PresentationThumbnailItemPreview />
          <PresentationThumbnailItemNumber className="absolute right-1.5 bottom-1.5" />
        </>
      )}
    </Primitive.ThumbnailItem>
  );
}

function PresentationThumbnailList({ className, style, children, ...props }: ThumbnailListProps) {
  return (
    <Primitive.ThumbnailList
      data-slot="presentation-thumbnail-list"
      className={cn("group/thumbs w-40 shrink-0 overflow-y-auto border-r p-4", className)}
      style={style}
      {...props}
    >
      {children ??
        (({ slides }) =>
          slides.map((slide) => <PresentationThumbnailItem key={slide.id} slideId={slide.id} />))}
    </Primitive.ThumbnailList>
  );
}

export {
  Presentation,
  PresentationBody,
  PresentationContent,
  PresentationError,
  PresentationLoading,
  PresentationSlide,
  PresentationThumbnailItem,
  PresentationThumbnailItemNumber,
  PresentationThumbnailItemPreview,
  PresentationThumbnailList,
  PresentationViewport,
};
