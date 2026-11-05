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

interface ToolbarButtonProps extends React.ComponentProps<typeof Button> {
  label: string;
  /** Keys shown in the tooltip; `"mod"` is ⌘ on Apple platforms and Ctrl elsewhere. */
  shortcut?: string[];
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
              <ToolbarItem
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
            <ToolbarItem
              render={
                <SelectTrigger size="sm" aria-label="Font" className="h-7 w-32 text-xs">
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
    <Toolbar
      aria-label="Edit presentation"
      data-pptx-keep-selection=""
      // One scrolling row on narrow screens, rather than rows that squeeze the slide.
      className="border-b p-1.5 max-sm:flex-nowrap max-sm:overflow-x-auto max-sm:*:shrink-0"
    >
      <ToolbarButton label="Open presentation" onClick={onOpen}>
        <FolderOpenIcon />
      </ToolbarButton>
      <ToolbarButton label="Export" disabled={!isReady} onClick={() => run(onExport)}>
        <DownloadIcon />
        {isDirty && <span className="sr-only"> (unsaved changes)</span>}
      </ToolbarButton>
      <ToolbarSeparator />

      <ToolbarButton
        label="Undo"
        shortcut={["mod", "Z"]}
        disabled={!canUndo}
        onClick={() => undo()}
      >
        <Undo2Icon />
      </ToolbarButton>
      <ToolbarButton
        label="Redo"
        shortcut={["mod", "shift", "Z"]}
        disabled={!canRedo}
        onClick={() => run(redo)}
      >
        <Redo2Icon />
      </ToolbarButton>
      <ToolbarSeparator />

      <ToolbarButton label="New slide" disabled={!isReady} onClick={() => run(onAddSlide)}>
        <PlusIcon />
      </ToolbarButton>
      <ToolbarButton
        label="Text box"
        disabled={!isReady}
        pressed={drawTool?.type === "textBox"}
        onClick={() => startDrawing(drawTool?.type === "textBox" ? null : { type: "textBox" })}
      >
        <TypeIcon />
      </ToolbarButton>
      <ShapeMenu
        disabled={!isReady}
        active={drawTool?.type === "shape"}
        onSelect={(geometry) => startDrawing({ type: "shape", geometry })}
      />
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
          <DropdownMenu modal={false}>
            <ToolbarMenuTrigger label="Arrange">
              <LayersIcon />
            </ToolbarMenuTrigger>
            <DropdownMenuContent data-pptx-keep-selection="" aria-label="Arrange" className="w-48">
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

      {hasText && (
        <>
          <ToolbarSeparator />
          <FontSelect
            value={textStyle.fontFamily ?? null}
            onChange={(fontFamily) => onFormat({ fontFamily })}
          />
          <ToolbarButton
            label="Decrease font size"
            disabled={textStyle.fontSize !== null && textStyle.fontSize <= MIN_FONT_SIZE}
            onClick={() =>
              onFormat({ fontSize: Math.max(MIN_FONT_SIZE, (textStyle.fontSize ?? 18) - 1) })
            }
          >
            <MinusIcon />
          </ToolbarButton>
          <span className="min-w-7 px-0.5 text-center text-xs tabular-nums">
            <span className="sr-only">Font size </span>
            {textStyle.fontSize ?? "–"}
          </span>
          <ToolbarButton
            label="Increase font size"
            disabled={textStyle.fontSize !== null && textStyle.fontSize >= MAX_FONT_SIZE}
            onClick={() =>
              onFormat({ fontSize: Math.min(MAX_FONT_SIZE, (textStyle.fontSize ?? 18) + 1) })
            }
          >
            <PlusIcon />
          </ToolbarButton>
          <ToolbarSeparator />
          {TEXT_TOGGLES.map(({ key, label, icon: Icon }) => (
            <ToolbarButton
              key={key}
              label={label}
              pressed={textStyle[key]}
              onClick={() => onFormat({ [key]: !textStyle[key] })}
            >
              <Icon />
            </ToolbarButton>
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

      <ToolbarItem
        disabled={!isReady}
        render={
          <PresentationZoomSelect className="ml-auto" aria-label="Zoom" disabled={!isReady} />
        }
      />
    </Toolbar>
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
                <SortableThumbnailItem key={slideId} slideId={slideId} actions={actions} />
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
    <ContextMenu>
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

interface SortableThumbnailItemProps {
  slideId: string;
  actions: SlideActions;
}

function SortableThumbnailItem({ slideId, actions }: SortableThumbnailItemProps) {
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
      onKeyDown={(event) => onThumbnailKeyDown(event, slideId, actions)}
      {...dragAttributes}
      {...listeners}
    >
      <PresentationThumbnailItemNumber />
      <PresentationThumbnailItemPreview />
    </PresentationThumbnailItem>
  );
}
