"use client";

import { Presentation as PresentationPrimitive } from "@diceui/pptx";

import { cn } from "@/lib/utils";

function PresentationProvider({ ...props }: PresentationPrimitive.Provider.Props) {
  return <PresentationPrimitive.Provider data-slot="presentation-provider" {...props} />;
}

function Presentation({ className, ...props }: PresentationPrimitive.Root.Props) {
  return (
    <PresentationPrimitive.Root
      data-slot="presentation-root"
      className={cn("flex gap-4 overflow-hidden", className)}
      {...props}
    />
  );
}

function PresentationContent({ className, ...props }: React.ComponentProps<"div">) {
  return (
    <div
      data-slot="presentation-content"
      className={cn("relative isolate flex flex-1 flex-col overflow-hidden", className)}
      {...props}
    />
  );
}

function PresentationViewport({ className, ...props }: PresentationPrimitive.Viewport.Props) {
  return (
    <PresentationPrimitive.Viewport
      data-slot="presentation-viewport"
      className={cn("flex flex-1 items-center justify-center overflow-hidden", className)}
      {...props}
    />
  );
}

function PresentationSlide(props: PresentationPrimitive.Slide.Props) {
  return <PresentationPrimitive.Slide data-slot="presentation-slide" {...props} />;
}

function PresentationSelection({ className, ...props }: PresentationPrimitive.Selection.Props) {
  return (
    <PresentationPrimitive.Selection
      data-slot="presentation-selection"
      className={cn("[--pptx-selection:var(--ring)]", className)}
      {...props}
    />
  );
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
        "absolute inset-0 z-10 flex items-center justify-center text-sm text-muted-foreground",
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
        "absolute inset-0 z-10 flex items-center justify-center text-sm text-destructive",
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
      className={cn(
        "w-full overflow-hidden rounded-sm data-pending:animate-pulse data-pending:bg-muted",
        className,
      )}
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
        "block text-center font-mono text-xs leading-none text-muted-foreground tabular-nums",
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
        "relative flex w-full cursor-pointer gap-1.5 rounded-md p-1.5 outline-none hover:bg-accent",
        "ring-2 ring-transparent transition-[color,ring] duration-100 focus:ring-ring data-active:bg-accent",
        className,
      )}
      {...props}
    >
      {children ?? (
        <>
          <PresentationThumbnailItemNumber />
          <PresentationThumbnailItemPreview />
        </>
      )}
    </PresentationPrimitive.ThumbnailItem>
  );
}

function PresentationThumbnailList({
  className,
  children,
  ...props
}: PresentationPrimitive.ThumbnailList.Props) {
  return (
    <PresentationPrimitive.ThumbnailList
      data-slot="presentation-thumbnail-list"
      className={cn(
        "flex w-40 shrink-0 flex-col gap-2 overflow-y-auto border-r p-1 pr-2",
        className,
      )}
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
  PresentationContent,
  PresentationError,
  PresentationLoading,
  PresentationProvider,
  PresentationSelection,
  PresentationSlide,
  PresentationThumbnailItem,
  PresentationThumbnailItemNumber,
  PresentationThumbnailItemPreview,
  PresentationThumbnailList,
  PresentationViewport,
};
