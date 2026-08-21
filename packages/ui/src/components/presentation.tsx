"use client";

import * as React from "react";

import { Presentation as PresentationPrimitive, useZoom } from "@diceui/pptx";
import { Button } from "@pptx/ui/components/button";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@pptx/ui/components/select";
import { cn } from "@pptx/ui/lib/utils";
import { PanelLeftIcon } from "lucide-react";

const ZOOM_LEVELS = [0.5, 0.75, 1, 1.5, 2];

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
      className={cn("flex flex-1 items-center justify-center overflow-hidden", className)}
      {...props}
    />
  );
}

interface PresentationZoomSelectProps extends Omit<
  React.ComponentProps<typeof SelectTrigger>,
  "value" | "onValueChange"
> {
  /** Explicit levels offered alongside the "Fit" entry, where `1` is 100%. */
  levels?: number[];
}

/**
 * Zoom control with a "Fit" entry and explicit levels.
 *
 * Holds no state: the store knows whether zoom is fitting, and picking a
 * level releases fitting on its own.
 */
function PresentationZoomSelect({
  levels = ZOOM_LEVELS,
  className,
  ...props
}: PresentationZoomSelectProps) {
  const { zoom, isAutoFit, setZoom, setAutoFit } = useZoom();
  const percentage = `${Math.round(zoom * 100)}%`;
  const itemClassName = "py-0.5 pr-6 pl-1.5 text-xs [&_svg:not([class*='size-'])]:size-3";

  return (
    <Select
      value={isAutoFit ? "fit" : String(zoom)}
      onValueChange={(value) => {
        if (value === "fit") setAutoFit(true);
        else setZoom(Number(value));
      }}
    >
      <SelectTrigger
        size="sm"
        className={cn(
          "h-6 w-22 gap-1 rounded-md py-0 pr-1.5 pl-2 text-xs [&_svg:not([class*='size-'])]:size-3",
          className,
        )}
        {...props}
      >
        <SelectValue>{isAutoFit ? `Fit · ${percentage}` : percentage}</SelectValue>
      </SelectTrigger>
      <SelectContent className="min-w-22 p-0.5">
        <SelectItem value="fit" className={itemClassName}>
          Fit
        </SelectItem>
        {levels.map((level) => (
          <SelectItem key={level} value={String(level)} className={itemClassName}>
            {`${level * 100}%`}
          </SelectItem>
        ))}
      </SelectContent>
    </Select>
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

function PresentationSelection({ className, ...props }: PresentationPrimitive.Selection.Props) {
  return (
    <PresentationPrimitive.Selection
      data-slot="presentation-selection"
      className={cn("[--presentation-selection:var(--ring)]", className)}
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
  const id = React.useId();
  const [open, setOpen] = React.useState(false);
  const listRef = React.useRef<HTMLDivElement | null>(null);
  const toggleRef = React.useRef<HTMLButtonElement | null>(null);

  React.useEffect(() => {
    if (!open) return;

    /** Pointer id of a gesture that started outside, until it taps or scrolls. */
    let pendingPointerId: number | null = null;

    function isInside(target: Node) {
      // The toggle owns its own click, so closing here would fight it.
      return !!listRef.current?.contains(target) || !!toggleRef.current?.contains(target);
    }

    function onPointerDown(event: PointerEvent) {
      pendingPointerId = isInside(event.target as Node) ? null : event.pointerId;
    }

    /**
     * Dismiss on release rather than on press: a touch that turns into a scroll
     * is cancelled by the browser, so only a tap that stays outside gets here.
     */
    function onPointerUp(event: PointerEvent) {
      const isPendingPointer = pendingPointerId === event.pointerId;
      pendingPointerId = null;
      if (isPendingPointer && !isInside(event.target as Node)) setOpen(false);
    }

    function onPointerCancel() {
      pendingPointerId = null;
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
  }, [open]);

  return (
    <>
      <Button
        ref={toggleRef}
        type="button"
        variant="outline"
        size="icon-sm"
        aria-expanded={open}
        aria-controls={id}
        aria-label={open ? "Hide slides" : "Show slides"}
        className="absolute top-2 left-2 z-10 bg-background/80 backdrop-blur-sm md:hidden"
        onClick={() => setOpen((value) => !value)}
      >
        <PanelLeftIcon />
      </Button>
      <PresentationPrimitive.ThumbnailList
        id={id}
        ref={listRef}
        render={<aside />}
        data-slot="presentation-thumbnail-list"
        className={cn(
          "flex w-40 shrink-0 flex-col gap-2 overflow-y-auto border-r bg-background p-1.5",
          // Above the toggle so the panel covers it instead of the first
          // thumbnail sitting under a floating button.
          "max-md:absolute max-md:inset-y-0 max-md:left-0 max-md:z-20 max-md:transition-transform",
          !open && "max-md:-translate-x-full",
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
  PresentationZoomSelect,
};
