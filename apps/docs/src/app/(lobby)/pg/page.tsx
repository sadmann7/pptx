"use client";

import * as React from "react";

import type { PresentationStore } from "@diceui/pptx";
import { useCreatePresentationStore, usePresentation, useSlide } from "@diceui/pptx";
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

export default function PgPage() {
  const id = React.useId();
  const store = useCreatePresentationStore();

  React.useEffect(() => {
    store.on("edit", (event) => {
      console.log("edit", event);
    });
    store.on("slideChange", (event) => {
      console.log("slideChange", event);
    });
    store.on("historyChange", (event) => {
      console.log("historyChange", event);
    });
    store.on("statusChange", (event) => {
      console.log("redo", event);
    });
    store.on("zoomChange", (event) => {
      console.log("zoomChange", event);
    });
  }, [store]);

  async function onFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    store.load(file, { defaultSlideIndex: 0, readOnly: false });
  }

  return (
    <div className="mx-auto flex h-[calc(100dvh-(--spacing(14)))] w-full max-w-(--fd-layout-width) flex-col gap-2">
      <div className="flex items-center gap-2 border-b border-border px-3 py-2">
        <Label htmlFor={`${id}-file-input`} className="sr-only">
          Open presentation
        </Label>
        <Input id={`${id}-file-input`} type="file" accept=".pptx" onChange={onFileChange} />
      </div>
      <PresentationProvider store={store}>
        <PresentationDebug />
        <PresentationToolbar store={store} />
        <Presentation className="flex-1">
          <SortableThumbnailList store={store} />
          <PresentationContent>
            <PresentationLoading />
            <PresentationError />
            <PresentationViewport>
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
                <SortableThumbnailItem key={slideId} slideId={slideId} />
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

function SortableThumbnailItem({ slideId }: { slideId: string }) {
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

interface PresentationToolbarProps {
  store: PresentationStore;
}

function PresentationToolbar({ store }: PresentationToolbarProps) {
  const { status } = usePresentation();
  const { slideId } = useSlide();

  const canUndo = React.useSyncExternalStore(
    store.subscribe,
    () => store.canUndo(),
    () => false,
  );

  const canRedo = React.useSyncExternalStore(
    store.subscribe,
    () => store.canRedo(),
    () => false,
  );

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
    <div className="flex items-center gap-2 border-b border-border px-3 py-1.5">
      <Button
        size="sm"
        variant="outline"
        onClick={() => slideId && run(() => store.edit({ type: "duplicateSlide", slideId }))}
      >
        Duplicate slide
      </Button>
      <Button
        size="sm"
        variant="outline"
        onClick={() => slideId && run(() => store.edit({ type: "deleteSlide", slideId }))}
      >
        Delete slide
      </Button>
      <Button size="sm" variant="ghost" disabled={!canUndo} onClick={() => store.undo()}>
        Undo
      </Button>
      <Button size="sm" variant="ghost" disabled={!canRedo} onClick={() => void store.redo()}>
        Redo
      </Button>
      <Button size="sm" className="ml-auto" onClick={onSave}>
        Save .pptx
      </Button>
    </div>
  );
}

function PresentationDebug() {
  const { status, progress, error, presentation } = usePresentation();
  const { slide, index } = useSlide();

  return (
    <div className="flex items-center gap-3 border-b border-border bg-muted px-3 py-1 font-mono text-xs text-muted-foreground">
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
