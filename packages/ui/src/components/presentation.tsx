"use client";

import * as React from "react";

import { Presentation as PresentationPrimitive } from "@diceui/pptx";
import { Button } from "@pptx/ui/components/button";
import { useIsMobile } from "@pptx/ui/hooks/use-mobile";
import { cn } from "cn";
import { PanelLeftIcon } from "lucide-react";

function PresentationProvider({ ...props }: PresentationPrimitive.Provider.Props) {
  return <PresentationPrimitive.Provider data-slot="presentation-provider" {...props} />;
}

function Presentation({ className, ...props }: PresentationPrimitive.Root.Props) {
  return (
    <PresentationPrimitive.Root
      data-slot="presentation"
      className={cn("relative flex overflow-hidden", className)}
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

function PresentationViewport({
  className,
  autoFit = true,
  autoFitPadding = 10,
  ...props
}: PresentationPrimitive.Viewport.Props) {
  return (
    <PresentationPrimitive.Viewport
      autoFit={autoFit}
      autoFitPadding={autoFitPadding}
      data-slot="presentation-viewport"
      className={cn(
        "flex flex-1 items-center justify-center overflow-hidden bg-background",
        className,
      )}
      {...props}
    />
  );
}

function PresentationSlide({ className, ...props }: PresentationPrimitive.Slide.Props) {
  return (
    <PresentationPrimitive.Slide
      data-slot="presentation-slide"
      className={cn("border", className)}
      {...props}
    />
  );
}

function PresentationSelection({
  className,
  children,
  ...props
}: PresentationPrimitive.Selection.Props) {
  return (
    <PresentationPrimitive.Selection
      data-slot="presentation-selection"
      // A fixed blue rather than a theme color: slides bring their own colors,
      // light or dark whatever the app's theme is.
      className={cn("[--presentation-selection:oklch(0.546_0.245_262.881)]", className)}
      {...props}
    >
      {children ?? (
        <>
          <PresentationSelectionBox>
            <PresentationResizeHandle />
            <PresentationRotateHandle />
          </PresentationSelectionBox>
          <PresentationSnapGuide />
          <PresentationMarquee />
        </>
      )}
    </PresentationPrimitive.Selection>
  );
}

function PresentationSelectionBox({
  className,
  ...props
}: PresentationPrimitive.SelectionBox.Props) {
  return (
    <PresentationPrimitive.SelectionBox
      data-slot="presentation-selection-box"
      className={withClassName(
        "shadow-[0_0_0_1.5px_var(--presentation-selection)] data-[mode=text]:shadow-[0_0_0_2px_var(--presentation-selection)]",
        className,
      )}
      {...props}
    />
  );
}

function PresentationResizeHandle({
  className,
  ...props
}: PresentationPrimitive.ResizeHandle.Props) {
  return (
    <PresentationPrimitive.ResizeHandle
      data-slot="presentation-resize-handle"
      className={withClassName(
        "size-2.5 rounded-[3px] border-[1.5px] border-(--presentation-selection) bg-white shadow-xs",
        className,
      )}
      {...props}
    />
  );
}

function PresentationRotateHandle({
  className,
  ...props
}: PresentationPrimitive.RotateHandle.Props) {
  return (
    <PresentationPrimitive.RotateHandle
      data-slot="presentation-rotate-handle"
      className={withClassName(
        "size-3 rounded-full border-[1.5px] border-(--presentation-selection) bg-white shadow-xs after:absolute after:top-full after:left-1/2 after:h-[calc(var(--presentation-rotate-handle-offset,20px)-6px)] after:w-px after:-translate-x-1/2 after:bg-(--presentation-selection)",
        className,
      )}
      {...props}
    />
  );
}

function PresentationSnapGuide({ className, ...props }: PresentationPrimitive.SnapGuide.Props) {
  return (
    <PresentationPrimitive.SnapGuide
      data-slot="presentation-snap-guide"
      className={withClassName("bg-(--presentation-snap-guide,#f43f5e)", className)}
      {...props}
    />
  );
}

function PresentationMarquee({ className, ...props }: PresentationPrimitive.Marquee.Props) {
  return (
    <PresentationPrimitive.Marquee
      data-slot="presentation-marquee"
      className={withClassName(
        "border border-(--presentation-selection) bg-[color-mix(in_oklch,var(--presentation-selection)_10%,transparent)]",
        className,
      )}
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
        "relative flex w-full cursor-pointer gap-1.5 rounded-md border p-1.5 outline-none hover:bg-accent",
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
  const isMobile = useIsMobile();
  const id = React.useId();
  const [open, setOpen] = React.useState(false);
  const [wasMobile, setWasMobile] = React.useState(isMobile);
  const listRef = React.useRef<HTMLDivElement | null>(null);
  const triggerRef = React.useRef<HTMLButtonElement | null>(null);

  const isHidden = isMobile && !open;

  if (wasMobile !== isMobile) {
    setWasMobile(isMobile);
    setOpen(false);
  }

  React.useEffect(() => {
    if (!isMobile || !open) return;

    // Pointers that went down outside the sheet; a tap outside closes it.
    const pendingPointerIds = new Set<number>();

    function isInside(target: Node) {
      return !!listRef.current?.contains(target) || !!triggerRef.current?.contains(target);
    }

    function onPointerDown(event: PointerEvent) {
      if (!isInside(event.target as Node)) pendingPointerIds.add(event.pointerId);
    }

    function onPointerUp(event: PointerEvent) {
      if (pendingPointerIds.delete(event.pointerId) && !isInside(event.target as Node)) {
        setOpen(false);
      }
    }

    function onPointerCancel(event: PointerEvent) {
      pendingPointerIds.delete(event.pointerId);
    }

    function onKeyDown(event: KeyboardEvent) {
      if (event.key === "Escape") setOpen(false);
    }

    document.addEventListener("pointerdown", onPointerDown);
    document.addEventListener("pointerup", onPointerUp);
    document.addEventListener("pointercancel", onPointerCancel);
    document.addEventListener("keydown", onKeyDown);
    return () => {
      document.removeEventListener("pointerdown", onPointerDown);
      document.removeEventListener("pointerup", onPointerUp);
      document.removeEventListener("pointercancel", onPointerCancel);
      document.removeEventListener("keydown", onKeyDown);
    };
  }, [isMobile, open]);

  return (
    <>
      {isMobile && (
        <Button
          type="button"
          aria-controls={id}
          aria-expanded={open}
          variant="outline"
          size="icon-sm"
          ref={triggerRef}
          className="absolute top-2 left-2 z-10 bg-background/80 backdrop-blur-sm"
          onClick={() => setOpen((value) => !value)}
        >
          <PanelLeftIcon />
        </Button>
      )}
      <PresentationPrimitive.ThumbnailList
        id={id}
        aria-hidden={isHidden || undefined}
        data-slot="presentation-thumbnail-list"
        data-mobile={isMobile || undefined}
        inert={isHidden}
        ref={listRef}
        render={<aside />}
        className={cn(
          "flex w-40 shrink-0 flex-col gap-2 overflow-y-auto border-r bg-background p-1.5",
          isMobile && "absolute inset-y-0 left-0 z-20 transition-transform",
          isHidden && "-translate-x-full",
          className,
        )}
        {...props}
      >
        {children ??
          (({ slides }) =>
            slides.map((slide) => <PresentationThumbnailItem key={slide.id} slideId={slide.id} />))}
      </PresentationPrimitive.ThumbnailList>
    </>
  );
}

function withClassName<State>(
  base: string,
  className: string | ((state: State) => string | undefined) | undefined,
) {
  return (state: State) => cn(base, typeof className === "function" ? className(state) : className);
}

export {
  Presentation,
  PresentationContent,
  PresentationError,
  PresentationLoading,
  PresentationMarquee,
  PresentationProvider,
  PresentationResizeHandle,
  PresentationRotateHandle,
  PresentationSelection,
  PresentationSelectionBox,
  PresentationSlide,
  PresentationSnapGuide,
  PresentationThumbnailItem,
  PresentationThumbnailItemNumber,
  PresentationThumbnailItemPreview,
  PresentationThumbnailList,
  PresentationViewport,
};
