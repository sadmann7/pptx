"use client";

import { Presentation as PresentationPrimitive } from "@diceui/pptx";

import { cn } from "@/lib/utils";

function Presentation({ className, ...props }: PresentationPrimitive.Root.Props) {
  return (
    <PresentationPrimitive.Root
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

function PresentationViewport({ className, ...props }: PresentationPrimitive.Viewport.Props) {
  return (
    <PresentationPrimitive.Viewport
      data-slot="presentation-viewport"
      className={cn("flex-1", className)}
      {...props}
    />
  );
}

function PresentationSlide(props: PresentationPrimitive.Slide.Props) {
  return <PresentationPrimitive.Slide data-slot="presentation-slide" {...props} />;
}

function PresentationLoading({
  className,
  children,
  ...props
}: PresentationPrimitive.Loading.Props) {
  return (
    <PresentationPrimitive.Loading
      data-slot="presentation-loading"
      className={cn(
        "absolute inset-0 z-10 flex items-center justify-center bg-background/80 text-sm text-muted-foreground",
        className,
      )}
      {...props}
    >
      {children ?? ((progress) => <span>Loading… {progress}%</span>)}
    </PresentationPrimitive.Loading>
  );
}

function PresentationError({ className, children, ...props }: PresentationPrimitive.Error.Props) {
  return (
    <PresentationPrimitive.Error
      data-slot="presentation-error"
      className={cn(
        "absolute inset-0 z-10 flex items-center justify-center bg-background/80 text-sm text-destructive",
        className,
      )}
      {...props}
    >
      {children ?? ((error: Error) => <span>{error.message}</span>)}
    </PresentationPrimitive.Error>
  );
}

function PresentationThumbnailItemPreview({
  className,
  ...props
}: PresentationPrimitive.ThumbnailItemPreview.Props) {
  return (
    <PresentationPrimitive.ThumbnailItemPreview
      data-slot="presentation-thumbnail-item-preview"
      className={cn("w-full overflow-hidden rounded-sm", className)}
      {...props}
    />
  );
}

function PresentationThumbnailItemNumber({
  className,
  ...props
}: PresentationPrimitive.ThumbnailItemNumber.Props) {
  return (
    <PresentationPrimitive.ThumbnailItemNumber
      data-slot="presentation-thumbnail-item-number"
      className={cn(
        "mt-1 block text-center font-mono text-xs leading-none text-muted-foreground tabular-nums",
        className,
      )}
      {...props}
    />
  );
}

function PresentationThumbnailItem({
  className,
  children,
  ...props
}: PresentationPrimitive.ThumbnailItem.Props) {
  return (
    <PresentationPrimitive.ThumbnailItem
      data-slot="presentation-thumbnail-item"
      className={cn(
        "relative w-full cursor-pointer rounded-md outline-none not-first:mt-4",
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
    </PresentationPrimitive.ThumbnailItem>
  );
}

function PresentationThumbnailList({
  className,
  style,
  children,
  ...props
}: PresentationPrimitive.ThumbnailList.Props) {
  return (
    <PresentationPrimitive.ThumbnailList
      data-slot="presentation-thumbnail-list"
      className={cn("group/thumbs w-40 shrink-0 overflow-y-auto border-r p-4", className)}
      style={style}
      {...props}
    >
      {children ??
        (({ slides }) =>
          slides.map((slide) => <PresentationThumbnailItem key={slide.id} slideId={slide.id} />))}
    </PresentationPrimitive.ThumbnailList>
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
