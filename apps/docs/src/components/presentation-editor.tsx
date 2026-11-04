"use client";

import * as React from "react";

import type { NodeOrder, PresentationStore, TextStyle } from "@diceui/pptx";
import {
  getPresentationFonts,
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
import { Button, buttonVariants } from "@pptx/ui/components/button";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "@pptx/ui/components/dropdown-menu";
import {
  Empty,
  EmptyContent,
  EmptyDescription,
  EmptyHeader,
  EmptyMedia,
  EmptyTitle,
} from "@pptx/ui/components/empty";
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
  SelectGroup,
  SelectItem,
  SelectLabel,
  SelectTrigger,
  SelectValue,
} from "@pptx/ui/components/select";
import {
  Toolbar,
  ToolbarButton as ToolbarItem,
  ToolbarSeparator,
} from "@pptx/ui/components/toolbar";
import { Tooltip, TooltipContent, TooltipTrigger } from "@pptx/ui/components/tooltip";
import { cn } from "@pptx/ui/lib/utils";
import {
  AArrowDownIcon,
  AArrowUpIcon,
  BaselineIcon,
  BoldIcon,
  CopyIcon,
  CopyPlusIcon,
  DownloadIcon,
  FolderOpenIcon,
  ImageIcon,
  ItalicIcon,
  LayersIcon,
  PaintBucketIcon,
  PlusIcon,
  PresentationIcon,
  Redo2Icon,
  SquareDashedIcon,
  StrikethroughIcon,
  TextAlignCenterIcon,
  TextAlignEndIcon,
  TextAlignJustifyIcon,
  TextAlignStartIcon,
  Trash2Icon,
  TypeIcon,
  UnderlineIcon,
  Undo2Icon,
} from "lucide-react";
import { toast } from "sonner";

import { PresentationZoomSelect } from "@/components/presentation-zoom-select";

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

const COLORS = [
  { name: "Black", value: "#000000" },
  { name: "Gray", value: "#6B7280" },
  { name: "White", value: "#FFFFFF" },
  { name: "Red", value: "#DC2626" },
  { name: "Orange", value: "#EA580C" },
  { name: "Yellow", value: "#CA8A04" },
  { name: "Green", value: "#16A34A" },
  { name: "Blue", value: "#2563EB" },
  { name: "Purple", value: "#7C3AED" },
  { name: "Pink", value: "#DB2777" },
];

const ALIGNMENTS = [
  { align: "left", label: "Align left", icon: TextAlignStartIcon },
  { align: "center", label: "Align center", icon: TextAlignCenterIcon },
  { align: "right", label: "Align right", icon: TextAlignEndIcon },
  { align: "justify", label: "Justify", icon: TextAlignJustifyIcon },
] as const;

const TEXT_TOGGLES = [
  { key: "bold", label: "Bold", icon: BoldIcon },
  { key: "italic", label: "Italic", icon: ItalicIcon },
  { key: "underline", label: "Underline", icon: UnderlineIcon },
  { key: "strikethrough", label: "Strikethrough", icon: StrikethroughIcon },
] as const;

const ARRANGE = [
  { order: "front", label: "Bring to front" },
  { order: "forward", label: "Bring forward" },
  { order: "backward", label: "Send backward" },
  { order: "back", label: "Send to back" },
] as const satisfies readonly { order: NodeOrder; label: string }[];

/** Offered after the deck's own fonts. */
const COMMON_FONTS = ["Arial", "Calibri", "Georgia", "Helvetica", "Times New Roman", "Verdana"];

const PPTX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

function onEditError(error: unknown) {
  toast.error(error instanceof Error ? error.message : "Edit failed");
}

interface PresentationEditorProps {
  store: PresentationStore;
  className?: string;
  /** Rendered above the toolbar, such as a debug bar. */
  header?: React.ReactNode;
  /** Hides the toolbar, leaving direct manipulation and shortcuts. */
  toolbar?: boolean;
}

/**
 * A full slide editor: thumbnails to reorder, the slide with the selection
 * overlay, and a toolbar for everything the overlay doesn't do by hand.
 */
export function PresentationEditor({
  store,
  className,
  header,
  toolbar = true,
}: PresentationEditorProps) {
  const fileInputId = React.useId();
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  async function onFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    try {
      await store.load(file, { defaultSlideIndex: 0, readOnly: false });
    } catch {
      // `load()` already wrote the failure to `store.error`, which the editor shows.
    }
  }

  return (
    <div
      data-slot="presentation-editor"
      className={cn("flex flex-col overflow-hidden rounded-md border", className)}
    >
      <input
        ref={fileInputRef}
        id={fileInputId}
        type="file"
        accept=".pptx"
        hidden
        onChange={(event) => void onFileChange(event)}
      />
      <PresentationProvider store={store}>
        {header}
        {toolbar && <EditorToolbar store={store} onOpen={() => fileInputRef.current?.click()} />}
        <Presentation className="min-h-0 flex-1">
          <SortableThumbnailList store={store} />
          <PresentationContent>
            <PresentationLoading />
            <PresentationError />
            <EditorEmpty fileInputId={fileInputId} />
            <PresentationViewport autoFit autoFitPadding={12} scrollZoom>
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

function EditorEmpty({ fileInputId }: { fileInputId: string }) {
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
        <label
          htmlFor={fileInputId}
          className={cn(buttonVariants({ size: "sm" }), "cursor-pointer")}
        >
          Choose file
        </label>
      </EmptyContent>
    </Empty>
  );
}

interface ToolbarButtonProps extends React.ComponentProps<typeof Button> {
  label: string;
  /** Shown in the tooltip after the label. */
  shortcut?: string;
  /** Makes the button a toggle, announced as pressed or not. */
  pressed?: boolean;
}

/** An icon button in the toolbar's arrow-key navigation, with a tooltip. */
function ToolbarButton({
  label,
  shortcut,
  pressed,
  disabled,
  onClick,
  children,
  ...props
}: ToolbarButtonProps) {
  return (
    <Tooltip>
      <TooltipTrigger
        render={
          <ToolbarItem
            aria-label={label}
            aria-pressed={pressed}
            disabled={disabled}
            onClick={onClick}
            render={<Button variant={pressed ? "secondary" : "ghost"} size="icon-sm" {...props} />}
          >
            {children}
          </ToolbarItem>
        }
      />
      <TooltipContent>
        {label}
        {shortcut && <span className="ml-1.5 text-muted-foreground">{shortcut}</span>}
      </TooltipContent>
    </Tooltip>
  );
}

interface ColorMenuProps {
  label: string;
  icon: React.ComponentType<{ className?: string }>;
  disabled: boolean;
  value?: string | null;
  onSelect: (color: string) => void;
  /** Offers a "none" choice, e.g. to remove an outline. */
  noneLabel?: string;
  onNone?: () => void;
}

function ColorMenu({
  label,
  icon: Icon,
  disabled,
  value,
  onSelect,
  noneLabel,
  onNone,
}: ColorMenuProps) {
  return (
    <DropdownMenu>
      <DropdownMenuTrigger
        render={
          <ToolbarItem
            aria-label={label}
            disabled={disabled}
            render={<Button variant="ghost" size="icon-sm" className="relative" />}
          />
        }
      >
        <Icon />
        <span
          aria-hidden="true"
          className="absolute inset-x-1.5 bottom-1 h-0.5 rounded-full border border-border"
          style={{ backgroundColor: value ?? "transparent" }}
        />
      </DropdownMenuTrigger>
      <DropdownMenuContent data-pptx-keep-selection="" aria-label={label} className="w-auto">
        <div className="grid grid-cols-5 gap-1 p-1">
          {COLORS.map((color) => (
            <DropdownMenuItem
              key={color.value}
              aria-label={color.name}
              className="size-6 rounded-full border p-0 focus:ring-2 focus:ring-ring"
              style={{ backgroundColor: color.value }}
              onClick={() => onSelect(color.value)}
            />
          ))}
        </div>
        {onNone && (
          <>
            <DropdownMenuSeparator />
            <DropdownMenuItem onClick={onNone}>{noneLabel}</DropdownMenuItem>
          </>
        )}
      </DropdownMenuContent>
    </DropdownMenu>
  );
}

function FontSelect({
  disabled,
  value,
  onChange,
}: {
  disabled: boolean;
  value: string | null;
  onChange: (fontFamily: string) => void;
}) {
  const { presentation, revision } = usePresentation();
  const fonts = React.useMemo(
    () => (presentation ? getPresentationFonts(presentation) : null),
    // oxlint-disable-next-line react-hooks/exhaustive-deps -- edits mutate the deck in place; revision is the change signal
    [presentation, revision],
  );
  const deckFonts = [
    fonts?.theme.heading,
    fonts?.theme.body,
    ...(fonts?.embedded ?? []),
    ...(fonts?.used ?? []),
  ]
    .filter((font): font is string => Boolean(font))
    .filter((font, index, all) => all.indexOf(font) === index);
  const otherFonts = COMMON_FONTS.filter((font) => !deckFonts.includes(font));

  return (
    <Select
      value={value ?? ""}
      onValueChange={(next) => next && onChange(String(next))}
      disabled={disabled}
    >
      <ToolbarItem
        render={
          <SelectTrigger size="sm" aria-label="Font" className="h-7 w-36 text-xs">
            <SelectValue>{value ?? "Font"}</SelectValue>
          </SelectTrigger>
        }
      />
      <SelectContent data-pptx-keep-selection="" alignItemWithTrigger={false}>
        {deckFonts.length > 0 && (
          <SelectGroup>
            <SelectLabel>In this presentation</SelectLabel>
            {deckFonts.map((font) => (
              <SelectItem key={font} value={font} style={{ fontFamily: font }}>
                {font}
              </SelectItem>
            ))}
          </SelectGroup>
        )}
        <SelectGroup>
          <SelectLabel>Other fonts</SelectLabel>
          {otherFonts.map((font) => (
            <SelectItem key={font} value={font} style={{ fontFamily: font }}>
              {font}
            </SelectItem>
          ))}
        </SelectGroup>
      </SelectContent>
    </Select>
  );
}

interface EditorToolbarProps {
  store: PresentationStore;
  onOpen: () => void;
}

function EditorToolbar({ store, onOpen }: EditorToolbarProps) {
  const { status } = usePresentation();
  const { slideId, goTo } = useSlide();
  const { canUndo, canRedo, isDirty, undo, redo } = useHistory();
  const {
    nodes,
    textStyle,
    select,
    clear,
    setTextStyle,
    setParagraphStyle,
    insertTextBox,
    duplicate,
    remove,
    reorder,
  } = useSelection();
  const imageInputRef = React.useRef<HTMLInputElement>(null);

  const isReady = status === "ready" && slideId != null;
  const hasText = textStyle != null;
  const shapes = nodes.filter((node) => node.nodeType === "shape");
  const outlined = nodes.filter((node) => node.nodeType === "shape" || node.nodeType === "picture");

  function run(action: () => Promise<unknown>) {
    action().catch(onEditError);
  }

  function editNodes<T extends { nodeId: string }>(
    targets: { id: string }[],
    operation: (nodeId: string) => T & Parameters<PresentationStore["edit"]>[0],
  ) {
    if (!slideId || targets.length === 0) return;
    const operations = targets.map((node) => operation(node.id));
    run(() => store.edit(operations.length === 1 ? operations[0] : { type: "batch", operations }));
  }

  function onFormat(style: TextStyle) {
    run(() => setTextStyle(style));
  }

  async function onAddSlide() {
    const { createdSlideId } = await store.edit({
      type: "addSlide",
      afterSlideId: slideId ?? undefined,
    });
    if (createdSlideId) goTo(createdSlideId);
  }

  async function onImageChange(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file || !slideId) return;
    const { createdNodeIds } = await store.edit({
      type: "addImage",
      slideId,
      data: new Uint8Array(await file.arrayBuffer()),
      contentType: file.type,
      name: file.name,
    });
    if (createdNodeIds) select(createdNodeIds);
  }

  async function onExport() {
    const bytes = await store.save();
    const url = URL.createObjectURL(new Blob([bytes.slice().buffer], { type: PPTX_MIME_TYPE }));
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = "edited.pptx";
    anchor.click();
    URL.revokeObjectURL(url);
  }

  return (
    <Toolbar aria-label="Edit presentation" data-pptx-keep-selection="" className="border-b p-1.5">
      <ToolbarButton label="Open presentation" onClick={onOpen}>
        <FolderOpenIcon />
      </ToolbarButton>
      <ToolbarButton label="Export" disabled={!isReady} onClick={() => run(onExport)}>
        <DownloadIcon />
        {isDirty && <span className="sr-only"> (unsaved changes)</span>}
      </ToolbarButton>
      <ToolbarSeparator />

      <ToolbarButton label="Undo" shortcut="Ctrl/⌘Z" disabled={!canUndo} onClick={() => undo()}>
        <Undo2Icon />
      </ToolbarButton>
      <ToolbarButton label="Redo" shortcut="Ctrl/⌘⇧Z" disabled={!canRedo} onClick={() => run(redo)}>
        <Redo2Icon />
      </ToolbarButton>
      <ToolbarSeparator />

      <ToolbarButton label="New slide" disabled={!isReady} onClick={() => run(onAddSlide)}>
        <PlusIcon />
      </ToolbarButton>
      <ToolbarButton
        label="Duplicate slide"
        disabled={!isReady}
        onClick={() => slideId && run(() => store.edit({ type: "duplicateSlide", slideId }))}
      >
        <CopyPlusIcon />
      </ToolbarButton>
      <ToolbarButton
        label="Delete slide"
        disabled={!isReady}
        onClick={() => slideId && run(() => store.edit({ type: "deleteSlide", slideId }))}
      >
        <Trash2Icon />
      </ToolbarButton>
      <ToolbarSeparator />

      <ToolbarButton
        label="Text box"
        disabled={!isReady}
        onClick={() => run(() => insertTextBox())}
      >
        <TypeIcon />
      </ToolbarButton>
      <ToolbarButton
        label="Image"
        disabled={!isReady}
        onClick={() => imageInputRef.current?.click()}
      >
        <ImageIcon />
      </ToolbarButton>
      <input
        ref={imageInputRef}
        type="file"
        accept="image/png,image/jpeg,image/gif"
        hidden
        onChange={(event) => run(() => onImageChange(event))}
      />
      <ToolbarSeparator />

      <FontSelect
        disabled={!hasText}
        value={textStyle?.fontFamily ?? null}
        onChange={(fontFamily) => onFormat({ fontFamily })}
      />
      <ToolbarButton
        label="Decrease font size"
        disabled={!hasText}
        onClick={() => onFormat({ fontSize: Math.max(1, (textStyle?.fontSize ?? 18) - 2) })}
      >
        <AArrowDownIcon />
      </ToolbarButton>
      <span className="w-8 text-center text-xs text-muted-foreground tabular-nums">
        <span className="sr-only">Font size </span>
        {textStyle?.fontSize ?? "–"}
      </span>
      <ToolbarButton
        label="Increase font size"
        disabled={!hasText}
        onClick={() => onFormat({ fontSize: (textStyle?.fontSize ?? 18) + 2 })}
      >
        <AArrowUpIcon />
      </ToolbarButton>
      {TEXT_TOGGLES.map(({ key, label, icon: Icon }) => (
        <ToolbarButton
          key={key}
          label={label}
          pressed={textStyle?.[key] ?? false}
          disabled={!hasText}
          onClick={() => onFormat({ [key]: !textStyle?.[key] })}
        >
          <Icon />
        </ToolbarButton>
      ))}
      <ColorMenu
        label="Text color"
        icon={BaselineIcon}
        disabled={!hasText}
        value={textStyle?.color}
        onSelect={(color) => onFormat({ color })}
      />
      {ALIGNMENTS.map(({ align, label, icon: Icon }) => (
        <ToolbarButton
          key={align}
          label={label}
          pressed={textStyle?.align === align}
          disabled={!hasText}
          onClick={() => run(() => setParagraphStyle({ align }))}
        >
          <Icon />
        </ToolbarButton>
      ))}
      <ToolbarSeparator />

      <ColorMenu
        label="Fill"
        icon={PaintBucketIcon}
        disabled={shapes.length === 0}
        onSelect={(color) =>
          editNodes(shapes, (nodeId) => ({
            type: "setSolidFill",
            slideId: slideId!,
            nodeId,
            color,
          }))
        }
      />
      <ColorMenu
        label="Outline"
        icon={SquareDashedIcon}
        disabled={outlined.length === 0}
        onSelect={(color) =>
          editNodes(outlined, (nodeId) => ({
            type: "setOutline",
            slideId: slideId!,
            nodeId,
            color,
            width: 1.5,
          }))
        }
        noneLabel="No outline"
        onNone={() =>
          editNodes(outlined, (nodeId) => ({
            type: "setOutline",
            slideId: slideId!,
            nodeId,
            color: null,
          }))
        }
      />
      <DropdownMenu>
        <DropdownMenuTrigger
          render={
            <ToolbarItem
              aria-label="Arrange"
              disabled={nodes.length === 0}
              render={<Button variant="ghost" size="icon-sm" />}
            />
          }
        >
          <LayersIcon />
        </DropdownMenuTrigger>
        <DropdownMenuContent data-pptx-keep-selection="" aria-label="Arrange" className="w-44">
          {ARRANGE.map(({ order, label }) => (
            <DropdownMenuItem key={order} onClick={() => run(() => reorder(order))}>
              {label}
            </DropdownMenuItem>
          ))}
        </DropdownMenuContent>
      </DropdownMenu>
      <ToolbarButton
        label="Duplicate"
        shortcut="Ctrl/⌘D"
        disabled={nodes.length === 0}
        onClick={() => run(duplicate)}
      >
        <CopyIcon />
      </ToolbarButton>
      <ToolbarButton
        label="Delete"
        shortcut="Del"
        disabled={nodes.length === 0}
        onClick={() =>
          run(async () => {
            await remove();
            clear();
          })
        }
      >
        <Trash2Icon className="text-destructive" />
      </ToolbarButton>

      <ToolbarItem render={<PresentationZoomSelect className="ml-auto" aria-label="Zoom" />} />
    </Toolbar>
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
