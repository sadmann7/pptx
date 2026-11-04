"use client";

import * as React from "react";

import type { PresentationStore, TextStyle } from "@diceui/pptx";
import {
  useCreatePresentationStore,
  useHistory,
  usePresentation,
  useSelection,
  useSlide,
} from "@diceui/pptx";
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
import { Tooltip, TooltipContent, TooltipTrigger } from "@pptx/ui/components/tooltip";
import {
  AArrowDownIcon,
  AArrowUpIcon,
  BoldIcon,
  ItalicIcon,
  PlusIcon,
  Redo2Icon,
  StrikethroughIcon,
  TextAlignCenterIcon,
  TextAlignEndIcon,
  TextAlignJustifyIcon,
  TextAlignStartIcon,
  UnderlineIcon,
  Undo2Icon,
} from "lucide-react";
import { toast } from "sonner";

import { PresentationZoomSelect } from "@/components/presentation-zoom-select";
import { DEMO_DECK_PATH } from "@/lib/constants";

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

export function PresentationEditingDemo() {
  const store = useCreatePresentationStore();

  React.useEffect(() => {
    fetch(DEMO_DECK_PATH)
      .then((res) => {
        // fetch resolves on 404, so an unchecked body would reach the parser as
        // an error page rather than a deck.
        if (!res.ok) throw new Error(`${DEMO_DECK_PATH}: ${res.status}.`);
        return res.arrayBuffer();
      })
      // Editing and reordering both need the source package retained.
      .then((buffer) => store.load(buffer, { readOnly: false }))
      .catch(() => {
        // Fail silently to avoid blocking the main thread.
      });
    // oxlint-disable-next-line react-hooks/exhaustive-deps -- store is a stable ref, intentionally omitted from deps
  }, []);

  return (
    <div className="not-prose flex h-[calc(100dvh-(--spacing(32)))] flex-col overflow-hidden rounded-lg border">
      <PresentationProvider store={store}>
        <PresentationToolbar store={store} />
        <FormattingToolbar store={store} />
        <Presentation className="min-h-0 flex-1">
          <SortableThumbnailList store={store} />
          <PresentationContent>
            <PresentationLoading />
            <PresentationError />
            <PresentationViewport autoFit autoFitPadding={10}>
              <PresentationSlide>
                <PresentationSelection undoRedoShortcuts />
              </PresentationSlide>
            </PresentationViewport>
          </PresentationContent>
        </Presentation>
      </PresentationProvider>
    </div>
  );
}

function PresentationToolbar({ store }: { store: PresentationStore }) {
  const id = React.useId();
  const { status } = usePresentation();
  const { canUndo, canRedo, undo, redo } = useHistory();

  async function onFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      await store.load(file, { readOnly: false });
    } catch {
      // Fail silently because `load()` already wrote the failure to `store.error`.
    }
  }

  return (
    <div className="flex items-center gap-2 border-b p-1.5">
      <Label htmlFor={`${id}-file`} className="sr-only">
        Open .pptx
      </Label>
      <Input
        id={`${id}-file`}
        type="file"
        accept=".pptx"
        className="h-8 max-w-56 text-xs"
        onChange={onFileChange}
      />
      <span className="min-w-0 truncate text-sm text-muted-foreground">
        {status === "ready"
          ? "Drag a shape to move it, or drag a thumbnail to reorder the deck"
          : "Loading sample deck…"}
      </span>
      <Tooltip>
        <TooltipTrigger
          render={
            <Button
              aria-label="Undo"
              variant="ghost"
              size="icon-sm"
              className="ml-auto active:not-aria-[haspopup]:translate-y-0 not-data-disabled:active:not-aria-[haspopup]:translate-y-px data-disabled:opacity-50"
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
              className="active:not-aria-[haspopup]:translate-y-0 not-data-disabled:active:not-aria-[haspopup]:translate-y-px data-disabled:opacity-50"
              disabled={!canRedo}
              focusableWhenDisabled
              onClick={() => void redo().catch(() => toast.error("Redo failed"))}
            >
              <Redo2Icon />
            </Button>
          }
        />
        <TooltipContent>Redo</TooltipContent>
      </Tooltip>
      <PresentationZoomSelect />
    </div>
  );
}

const TEXT_COLORS = ["#111827", "#DC2626", "#2563EB", "#16A34A"];

const ALIGNMENTS = [
  { align: "left", label: "Align left", icon: TextAlignStartIcon },
  { align: "center", label: "Align center", icon: TextAlignCenterIcon },
  { align: "right", label: "Align right", icon: TextAlignEndIcon },
  { align: "justify", label: "Justify", icon: TextAlignJustifyIcon },
] as const;

/**
 * Formats the selected shapes, or the selected text while editing. The
 * `data-pptx-keep-selection` attribute keeps pressing it from clearing the
 * selection or ending text editing.
 */
function FormattingToolbar({ store }: { store: PresentationStore }) {
  const { textStyle, setTextStyle, setParagraphStyle } = useSelection();
  const { slideId, goTo } = useSlide();
  const disabled = !textStyle;

  function onFormat(style: TextStyle) {
    setTextStyle(style).catch(() => toast.error("Formatting failed"));
  }

  function onFontSizeStep(step: number) {
    onFormat({ fontSize: Math.max(1, (textStyle?.fontSize ?? 18) + step) });
  }

  async function onAddSlide() {
    try {
      const { createdSlideId } = await store.edit({
        type: "addSlide",
        afterSlideId: slideId ?? undefined,
      });
      if (createdSlideId) goTo(createdSlideId);
    } catch {
      toast.error("Could not add a slide");
    }
  }

  const toggles = [
    { key: "bold", label: "Bold", icon: BoldIcon },
    { key: "italic", label: "Italic", icon: ItalicIcon },
    { key: "underline", label: "Underline", icon: UnderlineIcon },
    { key: "strikethrough", label: "Strikethrough", icon: StrikethroughIcon },
  ] as const;

  return (
    <div data-pptx-keep-selection="" className="flex items-center gap-1 border-b p-1.5">
      <Button variant="ghost" size="sm" onClick={() => void onAddSlide()}>
        <PlusIcon />
        New slide
      </Button>
      <div className="mx-1 h-5 w-px bg-border" />
      {toggles.map(({ key, label, icon: Icon }) => (
        <Button
          key={key}
          aria-label={label}
          aria-pressed={textStyle?.[key] ?? false}
          variant={textStyle?.[key] ? "secondary" : "ghost"}
          size="icon-sm"
          disabled={disabled}
          onClick={() => onFormat({ [key]: !textStyle?.[key] })}
        >
          <Icon />
        </Button>
      ))}
      <div className="mx-1 h-5 w-px bg-border" />
      <Button
        aria-label="Decrease font size"
        variant="ghost"
        size="icon-sm"
        disabled={disabled}
        onClick={() => onFontSizeStep(-2)}
      >
        <AArrowDownIcon />
      </Button>
      <span className="w-8 text-center text-xs text-muted-foreground tabular-nums">
        {textStyle?.fontSize ?? "–"}
      </span>
      <Button
        aria-label="Increase font size"
        variant="ghost"
        size="icon-sm"
        disabled={disabled}
        onClick={() => onFontSizeStep(2)}
      >
        <AArrowUpIcon />
      </Button>
      <div className="mx-1 h-5 w-px bg-border" />
      {TEXT_COLORS.map((color) => (
        <button
          key={color}
          type="button"
          aria-label={`Text color ${color}`}
          aria-pressed={textStyle?.color === color}
          disabled={disabled}
          className="size-5 rounded-full border ring-offset-1 ring-offset-background disabled:opacity-50 aria-pressed:ring-2 aria-pressed:ring-ring"
          style={{ backgroundColor: color }}
          onClick={() => onFormat({ color })}
        />
      ))}
      <div className="mx-1 h-5 w-px bg-border" />
      {ALIGNMENTS.map(({ align, label, icon: Icon }) => (
        <Button
          key={align}
          aria-label={label}
          aria-pressed={textStyle?.align === align}
          variant={textStyle?.align === align ? "secondary" : "ghost"}
          size="icon-sm"
          disabled={disabled}
          onClick={() =>
            void setParagraphStyle({ align }).catch(() => toast.error("Formatting failed"))
          }
        >
          <Icon />
        </Button>
      ))}
    </div>
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
        <PresentationThumbnailList className="p-1.5">
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
