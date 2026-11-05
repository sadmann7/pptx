"use client";

import * as React from "react";

import type { NodeOrder, ParagraphAlign, PresentationStore, TextStyle } from "@diceui/pptx";
import {
  getPresentationFonts,
  getPresetShapePath,
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
  ContextMenu,
  ContextMenuContent,
  ContextMenuItem,
  ContextMenuSeparator,
  ContextMenuSub,
  ContextMenuSubContent,
  ContextMenuSubTrigger,
  ContextMenuTrigger,
} from "@pptx/ui/components/context-menu";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuSeparator,
  DropdownMenuShortcut,
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
import { Kbd, KbdGroup } from "@pptx/ui/components/kbd";
import { Popover, PopoverContent } from "@pptx/ui/components/popover";
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
  ToolbarInput,
  ToolbarButton,
  ToolbarSeparator,
} from "@pptx/ui/components/toolbar";
import {
  Tooltip,
  TooltipContent,
  TooltipProvider,
  TooltipTrigger,
} from "@pptx/ui/components/tooltip";
import { useIsMac } from "@pptx/ui/hooks/use-is-mac";
import { cn } from "@pptx/ui/lib/utils";
import {
  BaselineIcon,
  BoldIcon,
  CopyPlusIcon,
  DownloadIcon,
  FolderOpenIcon,
  ImageIcon,
  ItalicIcon,
  LayersIcon,
  MinusIcon,
  PaintBucketIcon,
  PlusIcon,
  PresentationIcon,
  Redo2Icon,
  ShapesIcon,
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

// The default drop side effect hides the source item, so its focus ring
// would only reappear ~250 ms after pointer release.
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
  { order: "front", label: "Bring to front", shortcut: ["mod", "shift", "]"] },
  { order: "forward", label: "Bring forward", shortcut: ["mod", "]"] },
  { order: "backward", label: "Send backward", shortcut: ["mod", "["] },
  { order: "back", label: "Send to back", shortcut: ["mod", "shift", "["] },
] as const satisfies readonly { order: NodeOrder; label: string; shortcut: readonly string[] }[];

/** The font sizes OOXML allows, in points. */
const MIN_FONT_SIZE = 1;
const MAX_FONT_SIZE = 4000;

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
  /** Shows formatting in a toolbar floating above the selection. */
  floatingToolbar?: boolean;
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
  floatingToolbar = false,
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
            <CanvasContextMenu store={store}>
              <PresentationViewport autoFit autoFitPadding={12} scrollZoom>
                <PresentationSlide>
                  <PresentationSelection undoRedoShortcuts />
                </PresentationSlide>
              </PresentationViewport>
            </CanvasContextMenu>
            {floatingToolbar && <FloatingToolbar store={store} />}
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

const APPLE_KEYS: Record<string, string> = { mod: "⌘", shift: "⇧", alt: "⌥" };
const OTHER_KEYS: Record<string, string> = { mod: "Ctrl", shift: "Shift", alt: "Alt" };

function Shortcut({ keys, className }: { keys: readonly string[]; className?: string }) {
  const isMac = useIsMac();
  const names = isMac ? APPLE_KEYS : OTHER_KEYS;
  return (
    <KbdGroup className={className}>
      {keys.map((key) => (
        <Kbd key={key}>{names[key] ?? key}</Kbd>
      ))}
    </KbdGroup>
  );
}

interface ToolbarIconButtonProps extends React.ComponentProps<typeof Button> {
  label: string;
  /** Keys shown in the tooltip; `"mod"` is ⌘ on Apple platforms and Ctrl elsewhere. */
  shortcut?: string[];
  /** Makes the button a toggle, announced as pressed or not. */
  pressed?: boolean;
}

/** An icon button in the toolbar's arrow-key navigation, with a tooltip. */
function ToolbarIconButton({
  label,
  shortcut,
  pressed,
  disabled,
  onClick,
  children,
  ...props
}: ToolbarIconButtonProps) {
  return (
    <Tooltip>
      <TooltipTrigger
        render={
          <ToolbarButton
            aria-label={label}
            aria-pressed={pressed}
            disabled={disabled}
            onClick={onClick}
            render={<Button variant={pressed ? "secondary" : "ghost"} size="icon-sm" {...props} />}
          >
            {children}
          </ToolbarButton>
        }
      />
      <TooltipContent>
        {label}
        {shortcut && <Shortcut keys={shortcut} className="ml-1.5" />}
      </TooltipContent>
    </Tooltip>
  );
}

interface ToolbarMenuTriggerProps {
  label: string;
  disabled?: boolean;
  /** Shows the trigger pressed, e.g. while its tool is armed. */
  active?: boolean;
  className?: string;
  children: React.ReactNode;
}

/** A menu's trigger button in the toolbar, with a tooltip like the other buttons. */
function ToolbarMenuTrigger({
  label,
  disabled,
  active,
  className,
  children,
}: ToolbarMenuTriggerProps) {
  return (
    <Tooltip>
      <TooltipTrigger
        render={
          <DropdownMenuTrigger
            render={
              <ToolbarButton
                aria-label={label}
                disabled={disabled}
                render={
                  <Button
                    variant={active ? "secondary" : "ghost"}
                    size="icon-sm"
                    className={className}
                  />
                }
              />
            }
          />
        }
      >
        {children}
      </TooltipTrigger>
      <TooltipContent>{label}</TooltipContent>
    </Tooltip>
  );
}

interface ColorMenuProps {
  label: string;
  icon: React.ComponentType<{ className?: string }>;
  disabled?: boolean;
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
    <DropdownMenu modal={false}>
      <ToolbarMenuTrigger label={label} disabled={disabled} className="relative">
        <Icon />
        <span
          aria-hidden="true"
          className="absolute inset-x-1.5 bottom-1 h-0.5 rounded-full border border-border"
          style={{ backgroundColor: value ?? "transparent" }}
        />
      </ToolbarMenuTrigger>
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
  value,
  onChange,
}: {
  value: string | null;
  onChange: (fontFamily: string) => void;
}) {
  const { presentation, revision } = usePresentation();
  // Listing the deck's fonts reads every slide, so only while the list is open:
  // a drag or an edit would otherwise rescan the deck on every commit.
  const [open, setOpen] = React.useState(false);
  const fonts = React.useMemo(
    () => (open && presentation ? getPresentationFonts(presentation) : null),
    // oxlint-disable-next-line react-hooks/exhaustive-deps -- edits mutate the deck in place; revision is the change signal
    [open, presentation, revision],
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
      modal={false}
      open={open}
      onOpenChange={setOpen}
      value={value ?? ""}
      onValueChange={(next) => next && onChange(String(next))}
    >
      <Tooltip>
        <TooltipTrigger
          render={
            <ToolbarButton
              render={
                <SelectTrigger
                  size="sm"
                  aria-label="Font"
                  className="h-7 w-32 border-transparent bg-transparent text-xs hover:bg-accent dark:bg-transparent dark:hover:bg-accent"
                >
                  {/* Blank when the selection mixes fonts, as in Google Slides. */}
                  <SelectValue className="min-w-0">
                    <span className="truncate">{value ?? ""}</span>
                  </SelectValue>
                </SelectTrigger>
              }
            />
          }
        />
        <TooltipContent>{value ? `Font: ${value}` : "Font"}</TooltipContent>
      </Tooltip>
      <SelectContent
        data-pptx-keep-selection=""
        alignItemWithTrigger={false}
        align="start"
        className="w-56"
      >
        {deckFonts.length > 0 && (
          <SelectGroup>
            <SelectLabel>In this presentation</SelectLabel>
            {deckFonts.map((font) => (
              <FontItem key={font} font={font} />
            ))}
          </SelectGroup>
        )}
        <SelectGroup>
          <SelectLabel>Other fonts</SelectLabel>
          {otherFonts.map((font) => (
            <FontItem key={font} font={font} />
          ))}
        </SelectGroup>
      </SelectContent>
    </Select>
  );
}

/** A font in the font list, in its own face, truncated to the list's width. */
function FontItem({ font }: { font: string }) {
  return (
    <SelectItem
      value={font}
      title={font}
      // Lets the item text shrink so a long name truncates.
      className="*:first:min-w-0 *:first:shrink"
      style={{ fontFamily: font }}
    >
      <span className="truncate">{font}</span>
    </SelectItem>
  );
}

/** Which ends of a horizontal scroller have content scrolled out of view. */
function useOverflowEdges(ref: React.RefObject<HTMLElement | null>) {
  const [edges, setEdges] = React.useState({ start: false, end: false });

  React.useEffect(() => {
    const element = ref.current;
    if (!element) return;
    function update() {
      if (!element) return;
      const start = element.scrollLeft > 1;
      const end = element.scrollLeft + element.clientWidth < element.scrollWidth - 1;
      setEdges((previous) =>
        previous.start === start && previous.end === end ? previous : { start, end },
      );
    }
    update();
    const observer = new ResizeObserver(update);
    observer.observe(element);
    for (const child of element.children) observer.observe(child);
    const mutations = new MutationObserver(update);
    mutations.observe(element, { childList: true });
    element.addEventListener("scroll", update, { passive: true });
    return () => {
      observer.disconnect();
      mutations.disconnect();
      element.removeEventListener("scroll", update);
    };
  }, [ref]);

  return edges;
}

/**
 * The font size as one control, as in Google Slides: step with the buttons or
 * type a size and press Enter. Blank when the selection mixes sizes.
 */
function FontSizeControl({
  value,
  onChange,
}: {
  value: number | null;
  onChange: (fontSize: number) => void;
}) {
  const [draft, setDraft] = React.useState<string | null>(null);
  const clamp = (size: number) => Math.min(MAX_FONT_SIZE, Math.max(MIN_FONT_SIZE, size));

  function commit() {
    if (draft === null) return;
    const size = Number.parseFloat(draft);
    setDraft(null);
    if (Number.isFinite(size) && size !== value) onChange(clamp(Math.round(size)));
  }

  return (
    <div className="flex items-center">
      <ToolbarIconButton
        label="Decrease font size"
        disabled={value !== null && value <= MIN_FONT_SIZE}
        onClick={() => onChange(clamp((value ?? 18) - 1))}
      >
        <MinusIcon />
      </ToolbarIconButton>
      <Tooltip>
        <TooltipTrigger
          render={
            <ToolbarInput
              aria-label="Font size"
              inputMode="numeric"
              value={draft ?? (value === null ? "" : String(value))}
              className="h-7 w-9 rounded-md bg-transparent text-center text-xs tabular-nums outline-none hover:bg-accent focus-visible:ring-3 focus-visible:ring-ring/50"
              onChange={(event) => setDraft(event.target.value.replace(/[^\d.]/g, ""))}
              onFocus={(event) => event.target.select()}
              onBlur={commit}
              onKeyDown={(event) => {
                if (event.key === "Enter") {
                  event.preventDefault();
                  commit();
                } else if (event.key === "Escape") {
                  setDraft(null);
                }
              }}
            />
          }
        />
        <TooltipContent>Font size</TooltipContent>
      </Tooltip>
      <ToolbarIconButton
        label="Increase font size"
        disabled={value !== null && value >= MAX_FONT_SIZE}
        onClick={() => onChange(clamp((value ?? 18) + 1))}
      >
        <PlusIcon />
      </ToolbarIconButton>
    </div>
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
    mode,
    nodes,
    textStyle,
    select,
    clear,
    setTextStyle,
    setParagraphStyle,
    drawTool,
    startDrawing,
    duplicate,
    remove,
    reorder,
  } = useSelection();
  const imageInputRef = React.useRef<HTMLInputElement>(null);
  const rowRef = React.useRef<HTMLDivElement>(null);
  const overflow = useOverflowEdges(rowRef);

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
    if (!createdSlideId) return;
    goTo(createdSlideId);
    focusThumbnail(createdSlideId);
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
    // A delay before the first tooltip, then instant ones while moving along the row.
    <TooltipProvider delay={400}>
      <Toolbar
        aria-label="Edit presentation"
        data-pptx-keep-selection=""
        className="flex-nowrap border-b p-1.5"
      >
        {/* One scrolling row, so the contextual tools never wrap onto a new line. */}
        <div
          ref={rowRef}
          data-overflow-start={overflow.start ? "" : undefined}
          data-overflow-end={overflow.end ? "" : undefined}
          className="-m-0.75 no-scrollbar flex min-w-0 flex-1 items-center gap-0.5 overflow-x-auto p-0.75 *:shrink-0 data-overflow-end:mask-r-from-[calc(100%-2rem)] data-overflow-start:mask-l-from-[calc(100%-2rem)]"
        >
          <ToolbarIconButton label="Open presentation" onClick={onOpen}>
            <FolderOpenIcon />
          </ToolbarIconButton>
          <ToolbarIconButton
            label={isDirty ? "Export (unsaved changes)" : "Export"}
            disabled={!isReady}
            className="relative"
            onClick={() => run(onExport)}
          >
            <DownloadIcon />
            {isDirty && (
              <span className="absolute top-1 right-1 size-1.5 rounded-full bg-primary">
                <span className="sr-only"> (unsaved changes)</span>
              </span>
            )}
          </ToolbarIconButton>
          <ToolbarSeparator />

          <ToolbarIconButton
            label="Undo"
            shortcut={["mod", "Z"]}
            disabled={!canUndo}
            onClick={() => undo()}
          >
            <Undo2Icon />
          </ToolbarIconButton>
          <ToolbarIconButton
            label="Redo"
            shortcut={["mod", "shift", "Z"]}
            disabled={!canRedo}
            onClick={() => run(redo)}
          >
            <Redo2Icon />
          </ToolbarIconButton>
          <ToolbarButton
            disabled={!isReady}
            render={
              <PresentationZoomSelect
                aria-label="Zoom"
                disabled={!isReady}
                className="h-7 border-transparent bg-transparent hover:bg-accent dark:bg-transparent dark:hover:bg-accent"
              />
            }
          />
          <ToolbarSeparator />

          <ToolbarIconButton label="New slide" disabled={!isReady} onClick={() => run(onAddSlide)}>
            <PlusIcon />
          </ToolbarIconButton>
          <ToolbarIconButton
            label="Text box"
            disabled={!isReady}
            pressed={drawTool?.type === "textBox"}
            onClick={() => startDrawing(drawTool?.type === "textBox" ? null : { type: "textBox" })}
          >
            <TypeIcon />
          </ToolbarIconButton>
          <ShapeMenu
            disabled={!isReady}
            active={drawTool?.type === "shape"}
            onSelect={(geometry) => startDrawing({ type: "shape", geometry })}
          />
          <ToolbarIconButton
            label="Image"
            disabled={!isReady}
            onClick={() => imageInputRef.current?.click()}
          >
            <ImageIcon />
          </ToolbarIconButton>
          <input
            ref={imageInputRef}
            type="file"
            accept="image/png,image/jpeg,image/gif"
            hidden
            onChange={(event) => run(() => onImageChange(event))}
          />

          {isReady && nodes.length === 0 && (
            <>
              <ToolbarSeparator />
              <ColorMenu
                label="Background"
                icon={PaintBucketIcon}
                onSelect={(color) =>
                  run(() => store.edit({ type: "setSlideBackground", slideId, color }))
                }
                noneLabel="Reset to layout"
                onNone={() =>
                  run(() => store.edit({ type: "setSlideBackground", slideId, color: null }))
                }
              />
            </>
          )}

          {nodes.length > 0 && mode !== "text" && (
            <>
              <ToolbarSeparator />
              {shapes.length > 0 && (
                <ColorMenu
                  label="Fill"
                  icon={PaintBucketIcon}
                  onSelect={(color) =>
                    editNodes(shapes, (nodeId) => ({
                      type: "setSolidFill",
                      slideId: slideId!,
                      nodeId,
                      color,
                    }))
                  }
                />
              )}
              {outlined.length > 0 && (
                <ColorMenu
                  label="Outline"
                  icon={SquareDashedIcon}
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
              )}
            </>
          )}

          {hasText && (
            <>
              <ToolbarSeparator />
              <FontSelect
                value={textStyle.fontFamily ?? null}
                onChange={(fontFamily) => onFormat({ fontFamily })}
              />
              <FontSizeControl
                value={textStyle.fontSize}
                onChange={(fontSize) => onFormat({ fontSize })}
              />
              <ToolbarSeparator />
              {TEXT_TOGGLES.map(({ key, label, icon: Icon }) => (
                <ToolbarIconButton
                  key={key}
                  label={label}
                  pressed={textStyle[key]}
                  onClick={() => onFormat({ [key]: !textStyle[key] })}
                >
                  <Icon />
                </ToolbarIconButton>
              ))}
              <ColorMenu
                label="Text color"
                icon={BaselineIcon}
                value={textStyle.color}
                onSelect={(color) => onFormat({ color })}
              />
              <AlignMenu
                value={textStyle.align}
                onSelect={(align) => run(() => setParagraphStyle({ align }))}
              />
            </>
          )}

          {nodes.length > 0 && mode !== "text" && (
            <>
              <ToolbarSeparator />
              <DropdownMenu modal={false}>
                <ToolbarMenuTrigger label="Arrange">
                  <LayersIcon />
                </ToolbarMenuTrigger>
                <DropdownMenuContent
                  data-pptx-keep-selection=""
                  aria-label="Arrange"
                  className="w-48"
                >
                  {ARRANGE.map(({ order, label, shortcut }) => (
                    <DropdownMenuItem key={order} onClick={() => run(() => reorder(order))}>
                      {label}
                      <DropdownMenuShortcut>
                        <Shortcut keys={shortcut} />
                      </DropdownMenuShortcut>
                    </DropdownMenuItem>
                  ))}
                  <DropdownMenuSeparator />
                  <DropdownMenuItem onClick={() => run(duplicate)}>
                    Duplicate
                    <DropdownMenuShortcut>
                      <Shortcut keys={["mod", "D"]} />
                    </DropdownMenuShortcut>
                  </DropdownMenuItem>
                  <DropdownMenuItem
                    variant="destructive"
                    onClick={() =>
                      run(async () => {
                        await remove();
                        clear();
                      })
                    }
                  >
                    Delete
                    <DropdownMenuShortcut>
                      <Shortcut keys={["Delete"]} />
                    </DropdownMenuShortcut>
                  </DropdownMenuItem>
                </DropdownMenuContent>
              </DropdownMenu>
            </>
          )}
        </div>
      </Toolbar>
    </TooltipProvider>
  );
}

const SHAPES = [
  { geometry: "rect", label: "Rectangle" },
  { geometry: "roundRect", label: "Rounded rectangle" },
  { geometry: "ellipse", label: "Oval" },
  { geometry: "triangle", label: "Triangle" },
  { geometry: "rtTriangle", label: "Right triangle" },
  { geometry: "diamond", label: "Diamond" },
  { geometry: "pentagon", label: "Pentagon" },
  { geometry: "hexagon", label: "Hexagon" },
  { geometry: "star5", label: "Star" },
  { geometry: "heart", label: "Heart" },
  { geometry: "chevron", label: "Chevron" },
  { geometry: "wedgeRoundRectCallout", label: "Callout" },
  { geometry: "rightArrow", label: "Right arrow" },
  { geometry: "leftArrow", label: "Left arrow" },
  { geometry: "upArrow", label: "Up arrow" },
  { geometry: "downArrow", label: "Down arrow" },
];

const SHAPE_PREVIEW_SIZE = 18;

function ShapeMenu({
  disabled,
  active,
  onSelect,
}: {
  disabled: boolean;
  active: boolean;
  onSelect: (geometry: string) => void;
}) {
  return (
    <DropdownMenu modal={false}>
      <ToolbarMenuTrigger label="Shape" disabled={disabled} active={active}>
        <ShapesIcon />
      </ToolbarMenuTrigger>
      <DropdownMenuContent
        data-pptx-keep-selection=""
        aria-label="Shape"
        // The armed slide takes focus, so Escape disarms it.
        finalFocus={false}
        className="w-auto"
      >
        <div className="grid grid-cols-4 gap-0.5">
          {SHAPES.map(({ geometry, label }) => (
            <DropdownMenuItem
              key={geometry}
              aria-label={label}
              title={label}
              className="size-8 justify-center p-0"
              onClick={() => onSelect(geometry)}
            >
              <svg
                aria-hidden="true"
                viewBox={`-1 -1 ${SHAPE_PREVIEW_SIZE + 2} ${SHAPE_PREVIEW_SIZE + 2}`}
                className="size-4.5 overflow-visible"
              >
                <path
                  d={getPresetShapePath(geometry, SHAPE_PREVIEW_SIZE, SHAPE_PREVIEW_SIZE)}
                  fill="none"
                  stroke="currentColor"
                  strokeWidth={1.5}
                  strokeLinejoin="round"
                />
              </svg>
            </DropdownMenuItem>
          ))}
        </div>
      </DropdownMenuContent>
    </DropdownMenu>
  );
}

function AlignMenu({
  value,
  onSelect,
}: {
  value: ParagraphAlign | null;
  onSelect: (align: ParagraphAlign) => void;
}) {
  const current = ALIGNMENTS.find(({ align }) => align === value) ?? ALIGNMENTS[0];
  return (
    <DropdownMenu modal={false}>
      <ToolbarMenuTrigger label="Align">
        <current.icon />
      </ToolbarMenuTrigger>
      <DropdownMenuContent data-pptx-keep-selection="" aria-label="Align" className="w-40">
        <DropdownMenuRadioGroup
          value={value ?? ""}
          onValueChange={(align) => onSelect(align as ParagraphAlign)}
        >
          {ALIGNMENTS.map(({ align, label, icon: Icon }) => (
            <DropdownMenuRadioItem key={align} value={align}>
              <Icon />
              {label}
            </DropdownMenuRadioItem>
          ))}
        </DropdownMenuRadioGroup>
      </DropdownMenuContent>
    </DropdownMenu>
  );
}

interface SortableThumbnailListProps {
  store: PresentationStore;
}

function SortableThumbnailList({ store }: SortableThumbnailListProps) {
  const { presentation } = usePresentation();
  const slideIds = presentation?.slides.map((slide) => slide.id) ?? [];
  const actions = useSlideActions(store);
  const [menuSlideId, setMenuSlideId] = React.useState<string | null>(null);
  const [isMenuOpen, setIsMenuOpen] = React.useState(false);
  const menuItemRef = React.useRef<HTMLElement | null>(null);

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

  const list = (
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
                <SortableThumbnailItem
                  key={slideId}
                  slideId={slideId}
                  actions={actions}
                  isMenuOpen={isMenuOpen && slideId === menuSlideId}
                />
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

  return (
    <ContextMenu
      onOpenChange={(open, { reason }) => {
        setIsMenuOpen(open);
        // The click outside moves focus itself; returning it first would flash the ring.
        if (reason === "outside-press") menuItemRef.current = null;
      }}
    >
      <ContextMenuTrigger
        className="contents"
        onContextMenuCapture={(event) => {
          // Like PowerPoint, right-clicking a thumbnail selects it, and it
          // keeps focus once the menu closes.
          const item = (event.target as Element).closest<HTMLElement>("[role=option]");
          item?.focus();
          menuItemRef.current = item;
          setMenuSlideId(item?.dataset.slideId ?? null);
        }}
      >
        {list}
      </ContextMenuTrigger>
      <SlideMenuContent
        slideId={menuSlideId}
        actions={actions}
        // Focus returns to the thumbnail, unless an action moves it on.
        finalFocus={() => menuItemRef.current ?? false}
        onAction={() => {
          menuItemRef.current = null;
        }}
      />
    </ContextMenu>
  );
}

interface SlideActions {
  add: (afterSlideId?: string) => Promise<void>;
  duplicate: (slideId: string) => Promise<void>;
  remove: (slideId: string) => Promise<void>;
  move: (slideId: string, to: SlideMove) => Promise<void>;
  canMove: (slideId: string, to: SlideMove) => boolean;
  canRemove: boolean;
}

type SlideMove = "up" | "down" | "start" | "end";

/** Focuses a slide's thumbnail once the edit that made or moved it renders. */
function focusThumbnail(slideId: string) {
  requestAnimationFrame(() => {
    document
      .querySelector<HTMLElement>(
        `[data-slot=presentation-editor] [role=option][data-slide-id="${slideId}"]`,
      )
      ?.focus();
  });
}

function useSlideActions(store: PresentationStore): SlideActions {
  const { presentation } = usePresentation();
  const slides = presentation?.slides ?? [];

  function indexOf(slideId: string) {
    return slides.findIndex((slide) => slide.id === slideId);
  }

  function targetIndex(slideId: string, to: SlideMove) {
    if (to === "up") return indexOf(slideId) - 1;
    if (to === "down") return indexOf(slideId) + 1;
    return to === "start" ? 0 : slides.length - 1;
  }

  async function goToCreated(edit: Promise<{ createdSlideId?: string }>) {
    const { createdSlideId } = await edit;
    if (!createdSlideId) return;
    store.goTo(createdSlideId);
    focusThumbnail(createdSlideId);
  }

  return {
    add: (afterSlideId) => goToCreated(store.edit({ type: "addSlide", afterSlideId })),
    duplicate: (slideId) => goToCreated(store.edit({ type: "duplicateSlide", slideId })),
    async remove(slideId) {
      await store.edit({ type: "deleteSlide", slideId });
      const activeSlideId = store.getActiveSlide()?.id;
      if (activeSlideId) focusThumbnail(activeSlideId);
    },
    async move(slideId, to) {
      await store.edit({ type: "moveSlide", slideId, toIndex: targetIndex(slideId, to) });
      focusThumbnail(slideId);
    },
    canMove(slideId, to) {
      const index = targetIndex(slideId, to);
      return index >= 0 && index < slides.length && index !== indexOf(slideId);
    },
    canRemove: slides.length > 1,
  };
}

/** PowerPoint's shortcuts for a focused thumbnail. */
function onThumbnailKeyDown(
  event: React.KeyboardEvent<HTMLElement>,
  slideId: string,
  actions: SlideActions,
) {
  const mod = event.metaKey || event.ctrlKey;
  let action: (() => Promise<void>) | null = null;
  if ((event.key === "Delete" || event.key === "Backspace") && !mod) {
    if (actions.canRemove) action = () => actions.remove(slideId);
  } else if (event.ctrlKey && !event.metaKey && event.key.toLowerCase() === "m") {
    // Ctrl on every platform: ⌘M minimizes the window on macOS.
    action = () => actions.add(slideId);
  } else if (mod && event.key.toLowerCase() === "d") {
    action = () => actions.duplicate(slideId);
  } else if (mod && (event.key === "ArrowUp" || event.key === "ArrowDown")) {
    const up = event.key === "ArrowUp";
    const to = event.shiftKey ? (up ? "start" : "end") : up ? "up" : "down";
    if (actions.canMove(slideId, to)) action = () => actions.move(slideId, to);
  }
  if (!action) return;
  event.preventDefault();
  action().catch(onEditError);
}

interface SlideMenuContentProps {
  /** The slide right-clicked, or `null` for the empty space below the thumbnails. */
  slideId: string | null;
  actions: SlideActions;
  /** Runs when an item is chosen, before its edit. */
  onAction: () => void;
}

function SlideMenuContent({
  slideId,
  actions,
  onAction,
  ...props
}: SlideMenuContentProps & React.ComponentProps<typeof ContextMenuContent>) {
  function run(action: () => Promise<unknown>) {
    onAction();
    action().catch(onEditError);
  }

  return (
    <ContextMenuContent aria-label="Slide" className="w-44" {...props}>
      <ContextMenuItem onClick={() => run(() => actions.add(slideId ?? undefined))}>
        <PlusIcon />
        New slide
      </ContextMenuItem>
      {slideId && (
        <>
          <ContextMenuItem onClick={() => run(() => actions.duplicate(slideId))}>
            <CopyPlusIcon />
            Duplicate slide
          </ContextMenuItem>
          <ContextMenuSeparator />
          <ContextMenuItem
            variant="destructive"
            disabled={!actions.canRemove}
            onClick={() => run(() => actions.remove(slideId))}
          >
            <Trash2Icon />
            Delete slide
          </ContextMenuItem>
        </>
      )}
    </ContextMenuContent>
  );
}

/**
 * The common formatting for the selection, floating just above it, as in Canva
 * and Keynote. Hidden while the selection is moved, resized, or rotated.
 */
function FloatingToolbar({ store }: { store: PresentationStore }) {
  const { slideId } = useSlide();
  const { mode, nodes, anchor, drawTool, textStyle, setTextStyle, setParagraphStyle } =
    useSelection();
  const isTextMode = mode === "text";
  const shapes = isTextMode ? [] : nodes.filter((node) => node.nodeType === "shape");
  const outlined = isTextMode
    ? []
    : nodes.filter((node) => node.nodeType === "shape" || node.nodeType === "picture");
  const hasShapeTools = shapes.length > 0 || outlined.length > 0;
  const isOpen =
    anchor !== null &&
    drawTool === null &&
    (mode === "selected" || isTextMode) &&
    (hasShapeTools || textStyle !== null);

  function run(action: () => Promise<unknown>) {
    action().catch(onEditError);
  }

  function editNodes(
    targets: { id: string }[],
    operation: (nodeId: string) => Parameters<PresentationStore["edit"]>[0],
  ) {
    if (!slideId || targets.length === 0) return;
    const operations = targets.map((node) => operation(node.id));
    run(() => store.edit(operations.length === 1 ? operations[0] : { type: "batch", operations }));
  }

  return (
    <Popover open={isOpen}>
      <PopoverContent
        anchor={anchor}
        side="top"
        align="start"
        // Clears the rotate handle, which sits 20px above the selection.
        sideOffset={24}
        collisionPadding={8}
        initialFocus={false}
        finalFocus={false}
        data-pptx-keep-selection=""
        // Shown on every selection, so it appears without animating.
        className="w-auto p-1 data-open:animate-none data-closed:animate-none"
      >
        <TooltipProvider delay={400}>
          <Toolbar aria-label="Format selection" className="flex-nowrap">
            {shapes.length > 0 && slideId && (
              <ColorMenu
                label="Fill"
                icon={PaintBucketIcon}
                onSelect={(color) =>
                  editNodes(shapes, (nodeId) => ({ type: "setSolidFill", slideId, nodeId, color }))
                }
              />
            )}
            {outlined.length > 0 && slideId && (
              <ColorMenu
                label="Outline"
                icon={SquareDashedIcon}
                onSelect={(color) =>
                  editNodes(outlined, (nodeId) => ({
                    type: "setOutline",
                    slideId,
                    nodeId,
                    color,
                    width: 1.5,
                  }))
                }
                noneLabel="No outline"
                onNone={() =>
                  editNodes(outlined, (nodeId) => ({
                    type: "setOutline",
                    slideId,
                    nodeId,
                    color: null,
                  }))
                }
              />
            )}
            {textStyle && (
              <>
                {hasShapeTools && <ToolbarSeparator />}
                {TEXT_TOGGLES.slice(0, 3).map(({ key, label, icon: Icon }) => (
                  <ToolbarIconButton
                    key={key}
                    label={label}
                    pressed={textStyle[key]}
                    onClick={() => run(() => setTextStyle({ [key]: !textStyle[key] }))}
                  >
                    <Icon />
                  </ToolbarIconButton>
                ))}
                <ColorMenu
                  label="Text color"
                  icon={BaselineIcon}
                  value={textStyle.color}
                  onSelect={(color) => run(() => setTextStyle({ color }))}
                />
                <AlignMenu
                  value={textStyle.align}
                  onSelect={(align) => run(() => setParagraphStyle({ align }))}
                />
              </>
            )}
          </Toolbar>
        </TooltipProvider>
      </PopoverContent>
    </Popover>
  );
}

interface CanvasContextMenuProps {
  store: PresentationStore;
  children: React.ReactNode;
}

/** Right-clicking a shape acts on it, and on empty canvas acts on the slide. */
function CanvasContextMenu({ store, children }: CanvasContextMenuProps) {
  const { slideId } = useSlide();
  const { mode, nodes, editText, duplicate, remove, reorder, clear } = useSelection();
  const actions = useSlideActions(store);
  const returnFocusRef = React.useRef<HTMLElement | null>(null);
  // Fixed when the menu opens, so it can't swap menus while closing.
  const [isShapeMenu, setIsShapeMenu] = React.useState(false);
  const [isTextMenu, setIsTextMenu] = React.useState(false);
  const textNode = nodes.length === 1 && nodes[0]?.nodeType === "shape" ? nodes[0] : null;

  function run(action: () => Promise<unknown>) {
    action().catch(onEditError);
  }

  return (
    <ContextMenu>
      <ContextMenuTrigger
        className="contents"
        onContextMenuCapture={(event) => {
          setIsShapeMenu(nodes.length > 0 || mode === "text");
          setIsTextMenu(mode === "text");
          // Text being edited gets its caret back; otherwise the overlay, for shortcuts.
          returnFocusRef.current =
            mode === "text" && document.activeElement instanceof HTMLElement
              ? document.activeElement
              : (event.target as Element).closest<HTMLElement>("[data-pptx-selection]");
        }}
      >
        {children}
      </ContextMenuTrigger>
      {isShapeMenu ? (
        <ContextMenuContent
          data-pptx-keep-selection=""
          // Pressing outside only closes the menu, keeping the selection and text editing.
          portalProps={{ render: <div data-pptx-keep-selection="" /> }}
          aria-label="Shape"
          className="w-44"
          finalFocus={() => returnFocusRef.current ?? false}
        >
          {textNode && !isTextMenu && (
            <ContextMenuItem
              onClick={() => {
                returnFocusRef.current = null;
                editText(textNode.id);
              }}
            >
              <TypeIcon />
              Edit text
            </ContextMenuItem>
          )}
          <ContextMenuItem onClick={() => run(duplicate)}>
            <CopyPlusIcon />
            Duplicate
          </ContextMenuItem>
          <ContextMenuSub>
            <ContextMenuSubTrigger>
              <LayersIcon />
              Order
            </ContextMenuSubTrigger>
            <ContextMenuSubContent data-pptx-keep-selection="" className="w-44">
              {ARRANGE.map(({ order, label }) => (
                <ContextMenuItem key={order} onClick={() => run(() => reorder(order))}>
                  {label}
                </ContextMenuItem>
              ))}
            </ContextMenuSubContent>
          </ContextMenuSub>
          <ContextMenuSeparator />
          <ContextMenuItem
            variant="destructive"
            onClick={() =>
              run(async () => {
                await remove();
                clear();
              })
            }
          >
            <Trash2Icon />
            Delete
          </ContextMenuItem>
        </ContextMenuContent>
      ) : (
        <SlideMenuContent
          slideId={slideId}
          actions={actions}
          finalFocus={() => returnFocusRef.current ?? false}
          onAction={() => {
            returnFocusRef.current = null;
          }}
        />
      )}
    </ContextMenu>
  );
}

interface SortableThumbnailItemProps {
  slideId: string;
  actions: SlideActions;
  isMenuOpen: boolean;
}

function SortableThumbnailItem({ slideId, actions, isMenuOpen }: SortableThumbnailItemProps) {
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
      // The menu takes focus while open; keep the ring on the slide it acts on.
      className={cn(isDragging && "opacity-30", isMenuOpen && "ring-ring")}
      onSelect={(event) => {
        // Pressing a thumbnail focuses it, and focus navigates. Suppress that
        // while dragging so the deck does not jump mid-gesture.
        if (isDragging) event.preventDefault();
      }}
      onKeyDown={(event) => onThumbnailKeyDown(event, slideId, actions)}
      {...dragAttributes}
      {...listeners}
    >
      <PresentationThumbnailItemNumber />
      <PresentationThumbnailItemPreview />
    </PresentationThumbnailItem>
  );
}
