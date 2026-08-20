"use client";

import * as React from "react";

import type { PresentationStore } from "@diceui/pptx";
import { useCreatePresentationStore, usePresentation } from "@diceui/pptx";
import {
  DndContext,
  type DragEndEvent,
  PointerSensor,
  closestCenter,
  useSensor,
  useSensors,
} from "@dnd-kit/core";
import { restrictToVerticalAxis } from "@dnd-kit/modifiers";
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

import { SAMPLE_DECK_PATH } from "@/lib/constants";

export function EditingDemo() {
  const store = useCreatePresentationStore();
  const [history, setHistory] = React.useState({ canUndo: false, canRedo: false });

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
        <Toolbar store={store} history={history} />
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
            <PresentationViewport autoFit autoFitPadding={10}>
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
}

function Toolbar({ store, history }: ToolbarProps) {
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
    </div>
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

  // Pointer only: the list owns ArrowUp/ArrowDown for roving focus, so a
  // keyboard drag sensor bound to the same keys would fight it.
  const sensors = useSensors(
    // A small threshold keeps a plain click selecting the slide instead of
    // starting a drag.
    useSensor(PointerSensor, { activationConstraint: { distance: 4 } }),
  );

  async function onDragEnd(event: DragEndEvent) {
    const { active, over } = event;
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
      modifiers={[restrictToVerticalAxis]}
      onDragEnd={onDragEnd}
    >
      <SortableContext items={orderedIds} strategy={verticalListSortingStrategy}>
        <PresentationThumbnailList className="m-1 hidden md:flex">
          {() => orderedIds.map((slideId) => <SortableItem key={slideId} slideId={slideId} />)}
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
      className={isDragging ? "z-10 opacity-80 shadow-lg" : undefined}
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
