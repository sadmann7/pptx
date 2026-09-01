"use client";

import * as React from "react";

import type { PresentationStore } from "@diceui/pptx";
import { useCreatePresentationStore, useHistory, usePresentation, useSlide } from "@diceui/pptx";
import type { DragEndEvent, DragStartEvent, DropAnimation } from "@dnd-kit/core";
import {
  DndContext,
  DragOverlay,
  PointerSensor,
  closestCenter,
  defaultDropAnimation,
  defaultDropAnimationSideEffects,
  useSensor,
  useSensors,
} from "@dnd-kit/core";
import { restrictToFirstScrollableAncestor, restrictToVerticalAxis } from "@dnd-kit/modifiers";
import { SortableContext, useSortable, verticalListSortingStrategy } from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";
import { Button, buttonVariants } from "@pptx/ui/components/button";
import {
  Empty,
  EmptyContent,
  EmptyDescription,
  EmptyHeader,
  EmptyMedia,
  EmptyTitle,
} from "@pptx/ui/components/empty";
import { Input } from "@pptx/ui/components/input";
import { Label } from "@pptx/ui/components/label";
import {
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
} from "@pptx/ui/components/presentation";
import { Separator } from "@pptx/ui/components/separator";
import { Tooltip, TooltipContent, TooltipTrigger } from "@pptx/ui/components/tooltip";
import { cn } from "@pptx/ui/lib/utils";
import { CopyPlusIcon, PresentationIcon, Redo2Icon, Trash2Icon, Undo2Icon } from "lucide-react";

import { PresentationZoomSelect } from "@/components/presentation-zoom-select";
import type { SearchParams } from "@/types";

type PlaygroundLayout = "default" | "editing" | "compact";
type PlaygroundHiddenTarget = "file-input" | "debug-toolbar" | "toolbar";

/**
 * Keep the source item visible during the drop animation.
 *
 * The default side effect sets its opacity to `0`, delaying the thumbnail
 * item's focus ring from reappearing until roughly 250 ms after pointer release.
 */
const DROP_ANIMATION: DropAnimation = {
  ...defaultDropAnimation,
  sideEffects: defaultDropAnimationSideEffects({ styles: { active: {} } }),
};

const HIDDEN_TARGETS: Record<PlaygroundLayout, ReadonlySet<PlaygroundHiddenTarget>> = {
  default: new Set(),
  editing: new Set(["file-input", "debug-toolbar"]),
  compact: new Set(["file-input", "debug-toolbar", "toolbar"]),
};

function getPlaygroundLayout(value: Awaited<SearchParams>[number]): PlaygroundLayout {
  return value === "editing" || value === "compact" ? value : "default";
}

interface PresentationPlaygroundProps {
  searchParams: Promise<SearchParams>;
}

export function PresentationPlayground({ searchParams }: PresentationPlaygroundProps) {
  const id = React.useId();
  const store = useCreatePresentationStore();
  const layout = getPlaygroundLayout(React.use(searchParams).layout);
  const hidden = HIDDEN_TARGETS[layout];

  async function onFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      await store.load(file, { defaultSlideIndex: 0, readOnly: false });
    } catch {
      // Fail silently because `load()` already wrote the failure to `store.error`.
    }
  }

  return (
    <div className="mx-auto flex w-full max-w-(--fd-layout-width) flex-col gap-2 p-4">
      <div className={cn("flex items-center gap-2", hidden.has("file-input") && "sr-only")}>
        <Label htmlFor={`${id}-file-input`} className="sr-only">
          Open presentation
        </Label>
        <Input id={`${id}-file-input`} type="file" accept=".pptx" onChange={onFileChange} />
      </div>
      <div
        className={cn(
          "flex flex-col overflow-hidden rounded-md border",
          hidden.has("file-input")
            ? "h-[calc(100dvh-(--spacing(82)))]"
            : "h-[calc(100dvh-(--spacing(32)))]",
        )}
      >
        <PresentationProvider store={store}>
          {!hidden.has("debug-toolbar") && <PresentationDebugToolbar />}
          {!hidden.has("toolbar") && <PresentationToolbar store={store} />}
          <Presentation className="flex-1">
            <SortableThumbnailList store={store} />
            <PresentationContent>
              <PresentationLoading />
              <PresentationError />
              <PresentationEmpty inputId={`${id}-file-input`} />
              <PresentationViewport autoFitPadding={12} scrollZoom>
                <PresentationSlide>
                  <PresentationSelection undoRedoShortcuts />
                </PresentationSlide>
              </PresentationViewport>
            </PresentationContent>
          </Presentation>
        </PresentationProvider>
      </div>
    </div>
  );
}

interface PresentationEmptyProps {
  inputId: string;
}

function PresentationEmpty({ inputId }: PresentationEmptyProps) {
  const { status } = usePresentation();
  if (status !== "idle") return null;

  return (
    <Empty className="absolute inset-0 z-10">
      <EmptyHeader>
        <EmptyMedia variant="icon">
          <PresentationIcon />
        </EmptyMedia>
        <EmptyTitle>No presentation open</EmptyTitle>
        <EmptyDescription>
          Pick a .pptx file to render it here, then edit, reorder, and save it back.
        </EmptyDescription>
      </EmptyHeader>
      <EmptyContent>
        <label htmlFor={inputId} className={cn(buttonVariants({ size: "sm" }), "cursor-pointer")}>
          Choose file
        </label>
      </EmptyContent>
    </Empty>
  );
}

interface SortableThumbnailListProps {
  store: PresentationStore;
}

function SortableThumbnailList({ store }: SortableThumbnailListProps) {
  const { presentation } = usePresentation();
  const slideIds = presentation?.slides.map((slide) => slide.id) ?? [];

  /**
   * Order to paint while a drop is being committed. `store.edit()` is async, so
   * without this the strip would snap back to the old order for a frame between
   * the pointer release and the edit landing.
   */
  const [pendingIds, setPendingIds] = React.useState<string[] | null>(null);
  const orderedIds = pendingIds ?? slideIds;

  const [draggedId, setDraggedId] = React.useState<string | null>(null);

  // Use PointerSensor only to prevent keyboard drag from conflicting with ArrowUp/Down roving focus.
  const sensors = useSensors(
    useSensor(PointerSensor, {
      activationConstraint: { distance: 4 },
    }),
  );

  async function onDragEnd(event: DragEndEvent) {
    const { active, over } = event;
    setDraggedId(null);
    if (!over || active.id === over.id) return;

    const slideId = String(active.id);
    const toIndex = orderedIds.indexOf(String(over.id));
    if (toIndex === -1) return;

    const fromIndex = orderedIds.indexOf(slideId);
    if (fromIndex === -1) return;

    const next = [...orderedIds];
    next.splice(fromIndex, 1);
    next.splice(toIndex, 0, slideId);
    setPendingIds(next);

    try {
      await store.edit({ type: "moveSlide", slideId, toIndex });
    } finally {
      setPendingIds(null);
    }
  }

  return (
    <DndContext
      sensors={sensors}
      collisionDetection={closestCenter}
      modifiers={[restrictToVerticalAxis, restrictToFirstScrollableAncestor]}
      onDragStart={({ active }: DragStartEvent) => setDraggedId(String(active.id))}
      onDragCancel={() => setDraggedId(null)}
      onDragEnd={onDragEnd}
    >
      <SortableContext items={orderedIds} strategy={verticalListSortingStrategy}>
        <PresentationThumbnailList className="p-2">
          {() => (
            <>
              {orderedIds.map((slideId) => (
                <SortableThumbnailItem key={slideId} slideId={slideId} />
              ))}
              <DragOverlay dropAnimation={DROP_ANIMATION}>
                {draggedId ? (
                  <PresentationThumbnailItem
                    decorative
                    slideId={draggedId}
                    className="h-full cursor-grabbing bg-background shadow-lg"
                  >
                    <PresentationThumbnailItemNumber />
                    <PresentationThumbnailItemPreview />
                  </PresentationThumbnailItem>
                ) : null}
              </DragOverlay>
            </>
          )}
        </PresentationThumbnailList>
      </SortableContext>
    </DndContext>
  );
}

interface SortableThumbnailItemProps {
  slideId: string;
}

function SortableThumbnailItem({ slideId }: SortableThumbnailItemProps) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } = useSortable({
    id: slideId,
  });

  // dnd-kit sets role="button" and tabIndex={0}; both would override what the
  // list needs, replacing the `option` role and putting every thumbnail in the
  // tab order instead of the one roving tab stop.
  const { role: _role, tabIndex: _tabIndex, ...dragAttributes } = attributes;

  return (
    <PresentationThumbnailItem
      slideId={slideId}
      ref={setNodeRef}
      style={{ transform: CSS.Transform.toString(transform), transition }}
      // The overlay carries the thumbnail during the drag, so what stays in the
      // list is just the slot it will land in.
      className={isDragging ? "opacity-30" : undefined}
      onSelect={(event) => {
        // Pressing a thumbnail focuses it, and focus navigates. Suppress that
        // while dragging so the deck does not jump mid-gesture.
        if (isDragging) event.preventDefault();
      }}
      {...dragAttributes}
      {...listeners}
    >
      <PresentationThumbnailItemNumber />
      <PresentationThumbnailItemPreview />
    </PresentationThumbnailItem>
  );
}

interface PresentationToolbarProps {
  store: PresentationStore;
}

function PresentationToolbar({ store }: PresentationToolbarProps) {
  const { status } = usePresentation();
  const { slideId } = useSlide();
  const { canUndo, canRedo, isDirty, undo, redo } = useHistory();

  if (status !== "ready") return null;

  function run(action: () => Promise<unknown>) {
    action().catch((err) => console.error("[pg] edit failed:", err));
  }

  function onSave() {
    run(async () => {
      const bytes = await store.save();
      const blob = new Blob([bytes.slice().buffer], {
        type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "edited.pptx";
      a.click();
      URL.revokeObjectURL(url);
    });
  }

  return (
    <div className="flex flex-wrap items-center gap-2 border-b p-2">
      <Tooltip>
        <TooltipTrigger
          render={
            <Button
              aria-label="Duplicate slide"
              variant="ghost"
              size="icon-sm"
              className="aria-disabled:opacity-50"
              disabled={!slideId}
              focusableWhenDisabled
              onClick={() => slideId && run(() => store.edit({ type: "duplicateSlide", slideId }))}
            >
              <CopyPlusIcon />
            </Button>
          }
        />
        <TooltipContent>Duplicate slide</TooltipContent>
      </Tooltip>
      <Tooltip>
        <TooltipTrigger
          render={
            <Button
              aria-label="Delete slide"
              variant="ghost"
              size="icon-sm"
              className="aria-disabled:opacity-50"
              disabled={!slideId}
              focusableWhenDisabled
              onClick={() => slideId && run(() => store.edit({ type: "deleteSlide", slideId }))}
            >
              <Trash2Icon />
            </Button>
          }
        />
        <TooltipContent>Delete slide</TooltipContent>
      </Tooltip>
      <Separator orientation="vertical" />
      <Tooltip>
        <TooltipTrigger
          render={
            <Button
              aria-label="Undo"
              variant="ghost"
              size="icon-sm"
              className="aria-disabled:opacity-50"
              disabled={!canUndo}
              focusableWhenDisabled
              onClick={() => undo()}
            >
              <Undo2Icon />
            </Button>
          }
        />
        <TooltipContent>Undo</TooltipContent>
      </Tooltip>
      <Tooltip>
        <TooltipTrigger
          render={
            <Button
              aria-label="Redo"
              variant="ghost"
              size="icon-sm"
              className="aria-disabled:opacity-50"
              disabled={!canRedo}
              focusableWhenDisabled
              onClick={() => void redo()}
            >
              <Redo2Icon />
            </Button>
          }
        />
        <TooltipContent>Redo</TooltipContent>
      </Tooltip>
      <PresentationZoomSelect className="ml-auto" />
      <Button size="sm" onClick={onSave}>
        Export
        {isDirty && <span className="sr-only"> (unsaved changes)</span>}
      </Button>
    </div>
  );
}

interface ThumbnailPerfDetail {
  frames: number;
  renders: number;
  totalMs: number;
  maxFrameMs: number;
  backlog: number;
}

function useThumbnailPerf(): ThumbnailPerfDetail | null {
  const perfRef = React.useRef<ThumbnailPerfDetail | null>(null);

  const subscribe = React.useCallback((onStoreChange: () => void) => {
    const onPerf = (event: Event) => {
      perfRef.current = (event as CustomEvent<ThumbnailPerfDetail>).detail;
      onStoreChange();
    };
    window.addEventListener("pptx:thumbnail-perf", onPerf);
    return () => window.removeEventListener("pptx:thumbnail-perf", onPerf);
  }, []);

  return React.useSyncExternalStore(
    subscribe,
    () => perfRef.current,
    () => null,
  );
}

function ThumbnailPerfReadout() {
  const perf = useThumbnailPerf();
  if (!perf) return null;

  const avgFrameMs = perf.frames > 0 ? perf.totalMs / perf.frames : 0;
  return (
    <span>
      thumbs:{" "}
      <strong className="text-foreground">
        {perf.renders} rendered · {avgFrameMs.toFixed(1)}ms avg · {perf.maxFrameMs.toFixed(1)}ms max
        · {perf.backlog} queued
      </strong>
    </span>
  );
}

function PresentationDebugToolbar() {
  const { status, progress, error, presentation } = usePresentation();
  const { slide, index } = useSlide();

  return (
    // Wrapping keeps each readout whole on a narrow viewport: a flex item moves
    // to the next line rather than breaking, and only the long perf readout is
    // ever wide enough to wrap inside itself.
    <div className="flex flex-wrap items-center gap-x-3 gap-y-0.5 border-b border-border bg-muted px-3 py-1 font-mono text-xs text-muted-foreground">
      <span>
        status: <strong className="text-foreground">{status}</strong>
      </span>
      {status === "loading" && <span>{progress}%</span>}
      {status === "ready" && presentation && (
        <>
          <span>
            slides: <strong className="text-foreground">{presentation.slides.length}</strong>
          </span>
          <span>
            size:{" "}
            <strong className="text-foreground">
              {presentation.width.toFixed(0)} x {presentation.height.toFixed(0)}
            </strong>
          </span>
          {slide && (
            <span>
              slide {index + 1}:{" "}
              <strong className="text-foreground">{slide.nodes.length} nodes</strong>
            </span>
          )}
          <ThumbnailPerfReadout />
        </>
      )}
      {error && <span className="text-destructive">{error.message}</span>}
    </div>
  );
}
