"use client";

import * as React from "react";

import type { PresentationStore } from "@diceui/pptx";
import { useCreatePresentationStore, usePresentation, useZoom } from "@diceui/pptx";
import {
  DndContext,
  type DragEndEvent,
  DragOverlay,
  type DragStartEvent,
  PointerSensor,
  closestCenter,
  useSensor,
  useSensors,
} from "@dnd-kit/core";
import { restrictToFirstScrollableAncestor, restrictToVerticalAxis } from "@dnd-kit/modifiers";
import { SortableContext, useSortable, verticalListSortingStrategy } from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";
import { Button } from "@pptx/ui/components/button";
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
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@pptx/ui/components/select";

import { SAMPLE_DECK_PATH } from "@/lib/constants";

const ZOOM_LEVELS = [0.5, 0.75, 1, 1.5, 2];

export function PresentationEditingDemo() {
  const store = useCreatePresentationStore();
  const [history, setHistory] = React.useState({ canUndo: false, canRedo: false });

  /** Off once a zoom level is picked, so a resize stops overriding the choice. */
  const [isAutoFit, setIsAutoFit] = React.useState(true);

  React.useEffect(() => {
    fetch(SAMPLE_DECK_PATH)
      .then((res) => {
        // fetch resolves on 404, so an unchecked body would reach the parser as
        // an error page rather than a deck.
        if (!res.ok) throw new Error(`${SAMPLE_DECK_PATH}: ${res.status}`);
        return res.arrayBuffer();
      })
      // Editing and reordering both need the source package retained.
      .then((buf) => store.load(buf, { readOnly: false }))
      .catch(() => {
        // Fail silently to avoid blocking the main thread
      });
    // oxlint-disable-next-line react-hooks/exhaustive-deps -- store is a stable ref, intentionally omitted from deps
  }, []);

  return (
    <div className="not-prose flex h-100 flex-col overflow-hidden rounded-lg border">
      <PresentationProvider store={store}>
        <Toolbar
          store={store}
          history={history}
          isAutoFit={isAutoFit}
          onAutoFitChange={setIsAutoFit}
        />
        <Presentation
          className="min-h-0 flex-1"
          // Undo/redo availability without polling the store, so the toolbar
          // stays correct for edits made by dragging too.
          onHistoryChange={({ canUndo, canRedo }) => setHistory({ canUndo, canRedo })}
        >
          <SortableThumbnailList store={store} />
          <PresentationContent>
            <PresentationLoading />
            <PresentationError />
            <PresentationViewport autoFit={isAutoFit} autoFitPadding={10}>
              <PresentationSlide>
                <PresentationSelection />
              </PresentationSlide>
            </PresentationViewport>
          </PresentationContent>
        </Presentation>
      </PresentationProvider>
    </div>
  );
}

interface ToolbarProps {
  store: PresentationStore;
  history: { canUndo: boolean; canRedo: boolean };
  isAutoFit: boolean;
  onAutoFitChange: (isAutoFit: boolean) => void;
}

function Toolbar({ store, history, isAutoFit, onAutoFitChange }: ToolbarProps) {
  const { status } = usePresentation();

  return (
    <div className="flex items-center gap-2 border-b px-3 py-2">
      <span className="text-sm text-muted-foreground">
        {status === "ready"
          ? "Drag a shape to move it, or drag a thumbnail to reorder the deck"
          : "Loading sample deck…"}
      </span>
      <Button
        size="sm"
        variant="ghost"
        className="ml-auto"
        disabled={!history.canUndo}
        onClick={() => store.undo()}
      >
        Undo
      </Button>
      <Button
        size="sm"
        variant="ghost"
        disabled={!history.canRedo}
        onClick={() => void store.redo()}
      >
        Redo
      </Button>
      <ZoomSelect isAutoFit={isAutoFit} onAutoFitChange={onAutoFitChange} />
    </div>
  );
}

function ZoomSelect({
  isAutoFit,
  onAutoFitChange,
}: Pick<ToolbarProps, "isAutoFit" | "onAutoFitChange">) {
  const { zoom, setZoom } = useZoom();
  const percentage = `${Math.round(zoom * 100)}%`;

  return (
    <Select
      value={isAutoFit ? "fit" : String(zoom)}
      onValueChange={(value) => {
        const next = String(value);
        // Auto-fit refits on every resize, so it has to be released for an
        // explicit level to survive the next one.
        onAutoFitChange(next === "fit");
        if (next !== "fit") setZoom(Number(next));
      }}
    >
      <SelectTrigger
        size="sm"
        className="h-6 w-22 gap-1 rounded-md py-0 pr-1.5 pl-2 text-xs [&_svg:not([class*='size-'])]:size-3"
      >
        <SelectValue>{isAutoFit ? `Fit · ${percentage}` : percentage}</SelectValue>
      </SelectTrigger>
      <SelectContent className="min-w-22 p-0.5">
        <SelectItem
          value="fit"
          className="py-0.5 pr-6 pl-1.5 text-xs [&_svg:not([class*='size-'])]:size-3"
        >
          Fit
        </SelectItem>
        {ZOOM_LEVELS.map((level) => (
          <SelectItem
            key={level}
            value={String(level)}
            className="py-0.5 pr-6 pl-1.5 text-xs [&_svg:not([class*='size-'])]:size-3"
          >
            {`${level * 100}%`}
          </SelectItem>
        ))}
      </SelectContent>
    </Select>
  );
}

function SortableThumbnailList({ store }: { store: PresentationStore }) {
  const { presentation } = usePresentation();
  const slideIds = presentation?.slides.map((slide) => slide.id) ?? [];

  /**
   * Order to paint while a drop is being committed. `store.edit()` is async, so
   * without this the strip would snap back to the old order for a frame between
   * the pointer release and the edit landing.
   */
  const [pendingIds, setPendingIds] = React.useState<string[] | null>(null);
  const orderedIds = pendingIds ?? slideIds;

  /** Slide under the pointer, mirrored into the drag overlay. */
  const [draggedId, setDraggedId] = React.useState<string | null>(null);

  // Pointer only: the list owns ArrowUp/ArrowDown for roving focus, so a
  // keyboard drag sensor bound to the same keys would fight it.
  const sensors = useSensors(
    // A small threshold keeps a plain click selecting the slide instead of
    // starting a drag.
    useSensor(PointerSensor, { activationConstraint: { distance: 4 } }),
  );

  async function onDragEnd(event: DragEndEvent) {
    const { active, over } = event;
    setDraggedId(null);
    if (!over || active.id === over.id) return;

    const slideId = String(active.id);
    const toIndex = orderedIds.indexOf(String(over.id));
    if (toIndex === -1) return;

    const fromIndex = orderedIds.indexOf(slideId);
    const next = [...orderedIds];
    next.splice(fromIndex, 1);
    next.splice(toIndex, 0, slideId);
    setPendingIds(next);

    try {
      await store.edit({ type: "moveSlide", slideId, toIndex });
    } finally {
      // The store is the source of truth again once the edit settles.
      setPendingIds(null);
    }
  }

  return (
    <DndContext
      sensors={sensors}
      collisionDetection={closestCenter}
      // A transformed child still counts toward its scroll container's overflow,
      // so an unclamped drag past the last thumbnail grows scrollHeight, which
      // lets auto-scroll run, which grows the transform again: the strip scrolls
      // forever. Clamping the drag to the scroll port breaks that loop.
      modifiers={[restrictToVerticalAxis, restrictToFirstScrollableAncestor]}
      onDragStart={({ active }: DragStartEvent) => setDraggedId(String(active.id))}
      onDragCancel={() => setDraggedId(null)}
      onDragEnd={onDragEnd}
    >
      <SortableContext items={orderedIds} strategy={verticalListSortingStrategy}>
        <PresentationThumbnailList>
          {() => (
            <>
              {orderedIds.map((slideId) => (
                <SortableItem key={slideId} slideId={slideId} />
              ))}
              {/*
               * Inside the list so the floating copy can read the list context
               * it needs to paint a real miniature. It is fixed-positioned, so
               * the strip's overflow does not clip it.
               */}
              <DragOverlay>
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

function SortableItem({ slideId }: { slideId: string }) {
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
