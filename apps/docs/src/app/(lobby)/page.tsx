import { buttonVariants } from "@pptx/ui/components/button";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@pptx/ui/components/tabs";
import { cn } from "@pptx/ui/lib/utils";
import { highlight } from "fumadocs-core/highlight";
import {
  Anchor,
  ArrowRight,
  CaseSensitive,
  Check,
  ChevronRight,
  Component,
  Database,
  Download,
  FileArchive,
  FileCode,
  GalleryVertical,
  Keyboard,
  Layers,
  Monitor,
  Move,
  Package,
  Pencil,
  Scan,
  Shapes,
  Table,
  Type,
  Undo2,
} from "lucide-react";
import Link from "next/link";

import { PresentationEditingDemo } from "@/components/presentation-editing-demo";
import { siteConfig } from "@/lib/site";

const PARTS = [
  { path: "[Content_Types].xml", size: "1.8 KB" },
  { path: "ppt/presentation.xml", size: "3.2 KB" },
  { path: "ppt/slides/slide1.xml", size: "6.4 KB" },
  { path: "ppt/slides/slide2.xml", size: "5.1 KB", rewritten: true },
  { path: "ppt/slides/slide3.xml", size: "7.9 KB" },
  { path: "ppt/slides/slide4.xml", size: "4.6 KB", rewritten: true },
  { path: "ppt/slideLayouts/slideLayout1.xml", size: "2.7 KB" },
  { path: "ppt/slideMasters/slideMaster1.xml", size: "12.3 KB" },
  { path: "ppt/theme/theme1.xml", size: "8.1 KB" },
  { path: "ppt/fonts/font1.fntdata", size: "184 KB" },
  { path: "ppt/media/image1.png", size: "412 KB" },
];

const ROUND_TRIP_POINTS = [
  "Embedded ODTTF and MTX fonts decoded and registered",
  "Layout, master, and theme inheritance resolved",
  "Typed edit operations with full undo and redo",
];

const FEATURES = [
  {
    icon: Shapes,
    title: "Faithful rendering",
    description:
      "Preset and custom geometries, gradients, pattern fills, shadows, tables, charts, and images.",
  },
  {
    icon: Move,
    title: "Direct manipulation",
    description:
      "Move, resize with aspect lock, rotate, snap to guides, multi-select, and marquee.",
  },
  {
    icon: Type,
    title: "Inline text editing",
    description: "Edit text in place, then format runs and paragraphs through useSelection.",
  },
  {
    icon: Table,
    title: "Table editing",
    description:
      "Insert tables, add rows and columns, merge and split cells, and style fills and borders.",
  },
  {
    icon: Keyboard,
    title: "Keyboard shortcuts",
    description: "Nudge, delete, duplicate, copy, paste, and arrange with PowerPoint's bindings.",
  },
  {
    icon: CaseSensitive,
    title: "Embedded fonts",
    description: "ODTTF and MTX fonts are decoded from the file so slides keep their typefaces.",
  },
  {
    icon: GalleryVertical,
    title: "Thumbnail strip",
    description:
      "Lazy miniatures on a shared render queue, so big decks never block the main thread.",
  },
  {
    icon: Scan,
    title: "Auto-fit zoom",
    description: "The slide scales to fill its viewport and re-fits whenever the window resizes.",
  },
  {
    icon: Undo2,
    title: "Undo and redo",
    description: "Every typed edit lands in history, with opt-in Ctrl/Cmd+Z and Ctrl/Cmd+Y.",
  },
];

const EDITOR_CODE = `import { Presentation } from "@diceui/pptx";

export function Editor({ file }) {
  return (
    <Presentation.Root file={file} readOnly={false}>
      <Presentation.ThumbnailList />
      <Presentation.Viewport autoFit>
        <Presentation.Slide>
          <Presentation.Selection undoRedoShortcuts />
        </Presentation.Slide>
      </Presentation.Viewport>
    </Presentation.Root>
  );
}`;

const TOOLBAR_CODE = `import { useSelection } from "@diceui/pptx";

export function Toolbar() {
  const { textStyle, setTextStyle } = useSelection();

  return (
    <div data-pptx-keep-selection="">
      <button
        aria-pressed={textStyle?.bold ?? false}
        disabled={!textStyle}
        onClick={() => void setTextStyle({ bold: !textStyle?.bold })}
      >
        Bold
      </button>
    </div>
  );
}`;

const CODE_FILES = [
  { name: "editor.tsx", code: EDITOR_CODE, highlightedLine: 9 },
  { name: "toolbar.tsx", code: TOOLBAR_CODE, highlightedLine: 11 },
];

const PRIMITIVES = [
  { name: "Root", description: "Loads the file and owns the store" },
  { name: "ThumbnailList", description: "Lazy slide previews" },
  { name: "Viewport", description: "Zoom, fit, and scroll navigation" },
  { name: "Slide", description: "Renders the active slide" },
  { name: "Selection", description: "Box, handles, snap guides, marquee" },
];

const REACT_LAYER = [
  { icon: Database, name: "Store", description: "Loading, navigation, zoom, history" },
  { icon: Component, name: "Primitives", description: "Root, Viewport, Slide, Selection" },
  { icon: Anchor, name: "Hooks", description: "useSelection, useZoom, and more" },
];

const CORE_STEPS = [
  { icon: Package, name: "Unzip" },
  { icon: FileCode, name: "Parse OOXML" },
  { icon: Layers, name: "Resolve inheritance" },
  { icon: Monitor, name: "Render to DOM" },
  { icon: Pencil, name: "Apply edits" },
  { icon: Download, name: "Save" },
];

export default function IndexPage() {
  return (
    <main className="flex flex-col">
      <Hero />
      <RoundTrip />
      <Features />
      <ComposableApi />
      <Architecture />
      <Footer />
    </main>
  );
}

function Hero() {
  return (
    <section className="flex flex-col items-center gap-16 px-4 pt-16 pb-24 md:px-8 md:pt-28 md:pb-36">
      <div className="flex max-w-4xl flex-col items-center gap-6 text-center">
        <Link
          href="/docs/editing"
          className="flex animate-fade-up items-center gap-2 rounded-full border bg-card py-1 pr-3 pl-1 text-[13px] text-muted-foreground opacity-0 transition-colors [animation-fill-mode:forwards] hover:text-foreground"
        >
          <span className="rounded-full bg-foreground px-2 py-0.5 text-[11px] font-semibold text-background">
            New
          </span>
          New in 0.4: table editing
          <ArrowRight className="size-3.5" />
        </Link>
        <h1 className="animate-fade-up text-4xl font-semibold tracking-tight text-balance opacity-0 [animation-delay:0.1s] [animation-fill-mode:forwards] md:text-[68px]/[71px] md:tracking-[-2.6px]">
          Open, edit, and save .pptx without leaving the browser
        </h1>
        <p className="max-w-[40rem] animate-fade-up text-lg text-balance text-muted-foreground opacity-0 [animation-delay:0.2s] [animation-fill-mode:forwards] md:text-[19px]/[29px]">
          Composable React primitives for viewing and editing PowerPoint decks. Real rendering,
          lossless saves, and no styles you have to fight.
        </p>
        <div className="flex animate-fade-up gap-3 pt-2 opacity-0 [animation-delay:0.3s] [animation-fill-mode:forwards]">
          <CtaButtons />
        </div>
      </div>
      <div className="flex w-full justify-center bg-[radial-gradient(ellipse_45%_40%_at_50%_55%,color-mix(in_srgb,var(--presentation-selection,#1a73e8)_8%,transparent),transparent)] py-4 md:py-0">
        <PresentationEditingDemo className="h-[22rem] w-full max-w-300 rounded-xl shadow-2xl shadow-black/40 md:h-[40.75rem]" />
      </div>
    </section>
  );
}

function CtaButtons() {
  return (
    <>
      <Link href="/docs" className={cn(buttonVariants({ size: "lg" }), "h-11 px-5 text-[15px]")}>
        Get started
        <ArrowRight />
      </Link>
      <Link
        href={siteConfig.links.github}
        target="_blank"
        rel="noreferrer"
        className={cn(buttonVariants({ variant: "outline", size: "lg" }), "h-11 px-5 text-[15px]")}
      >
        <GitHubIcon />
        GitHub
      </Link>
    </>
  );
}

function RoundTrip() {
  const rewritten = PARTS.filter((part) => part.rewritten).length;

  return (
    <Section className="grid items-center gap-16 lg:grid-cols-[minmax(0,26rem)_minmax(0,1fr)] lg:gap-30">
      <div className="flex flex-col gap-6">
        <SectionHeading eyebrow="Lossless save">
          Edit one slide.
          <br />
          Save the same file.
        </SectionHeading>
        <p className="text-[17px]/[26px] text-muted-foreground">
          Untouched parts round-trip byte-for-byte. Only the slides you edit are re-serialized, so
          charts, notes, and macros you never rendered survive the save.
        </p>
        <ul className="flex flex-col gap-3 pt-2 text-[15px]">
          {ROUND_TRIP_POINTS.map((point) => (
            <li key={point} className="flex items-center gap-3">
              <Check className="size-3.5 shrink-0 text-muted-foreground" />
              {point}
            </li>
          ))}
        </ul>
      </div>
      <div className="overflow-hidden rounded-xl border bg-card font-mono">
        <div className="flex items-center justify-between gap-4 border-b px-5 py-3.5">
          <div className="flex min-w-0 items-center gap-2.5 text-[13px]">
            <FileArchive className="size-4 shrink-0 text-muted-foreground" />
            <span className="truncate">after-the-needle-drops.pptx</span>
          </div>
          <div className="hidden items-center gap-4 text-xs text-muted-foreground sm:flex">
            <span className="flex items-center gap-1.5">
              <span className="size-1.5 rounded-full bg-green-500" />
              rewritten
            </span>
            <span className="flex items-center gap-1.5">
              <span className="size-1.5 rounded-full bg-muted-foreground/60" />
              byte-identical
            </span>
          </div>
        </div>
        <ul className="py-2.5">
          {PARTS.map((part) => (
            <li
              key={part.path}
              className={cn(
                "flex items-center gap-3 px-5 py-2 text-[13px]",
                part.rewritten && "bg-green-500/5",
              )}
            >
              <span
                className={cn(
                  "size-1.5 shrink-0 rounded-full",
                  part.rewritten ? "bg-green-500" : "bg-muted-foreground/40",
                )}
              />
              <span
                className={cn(
                  "min-w-0 flex-1 truncate",
                  part.rewritten ? "text-foreground" : "text-muted-foreground",
                )}
              >
                {part.path}
              </span>
              <span
                className={cn(
                  "hidden text-xs sm:inline",
                  part.rewritten
                    ? "text-green-600 dark:text-green-400"
                    : "text-muted-foreground/75",
                )}
              >
                {part.rewritten ? "rewritten" : "unchanged"}
              </span>
              <span className="w-16 text-right text-xs text-muted-foreground/75">{part.size}</span>
            </li>
          ))}
        </ul>
        <div className="flex items-center justify-between gap-4 border-t bg-muted/40 px-5 py-3.5 text-xs">
          <span className="text-muted-foreground">
            {rewritten} of {PARTS.length} parts re-serialized
          </span>
          <span>await store.save()</span>
        </div>
      </div>
    </Section>
  );
}

function Features() {
  return (
    <Section className="flex flex-col gap-16">
      <div className="flex flex-col justify-between gap-6 md:flex-row md:items-end">
        <SectionHeading eyebrow="Everything a deck needs">
          Renders like PowerPoint.
          <br />
          Edits like PowerPoint.
        </SectionHeading>
        <p className="max-w-[24rem] text-[17px]/[26px] text-muted-foreground">
          Keyboard shortcuts, snapping, and inline text editing behave the way people already
          expect.
        </p>
      </div>
      <div className="grid gap-px overflow-hidden rounded-xl border bg-border sm:grid-cols-2 lg:grid-cols-3">
        {FEATURES.map((feature) => (
          <div key={feature.title} className="flex flex-col gap-1.5 bg-background p-7 pb-7.5">
            <IconTile icon={feature.icon} />
            <h3 className="mt-4 text-[17px] font-semibold tracking-[-0.2px]">{feature.title}</h3>
            <p className="text-[15px]/[23px] text-muted-foreground">{feature.description}</p>
          </div>
        ))}
      </div>
    </Section>
  );
}

async function ComposableApi() {
  const files = await Promise.all(
    CODE_FILES.map(async (file) => ({
      ...file,
      rendered: await highlight(file.code, {
        lang: "tsx",
        themes: { light: "material-theme-lighter", dark: "material-theme-darker" },
        defaultColor: false,
        transformers: [
          {
            line(node, line) {
              if (line === file.highlightedLine) node.properties["data-highlighted"] = "";
            },
          },
        ],
        components: {
          pre: (props) => (
            <pre
              className="overflow-x-auto py-4 text-[13.5px]/6 [&_code]:grid [&_code]:[counter-reset:line] [&_span]:text-(--shiki-light) dark:[&_span]:text-(--shiki-dark)"
              tabIndex={0}
            >
              {props.children}
            </pre>
          ),
        },
      }),
    })),
  );

  return (
    <Section className="grid items-center gap-16 lg:grid-cols-[minmax(0,1fr)_minmax(0,27.5rem)] lg:gap-24">
      <Tabs
        defaultValue={files[0]?.name}
        className="order-last min-w-0 gap-0 overflow-hidden rounded-xl border bg-card font-mono lg:order-none"
      >
        <TabsList variant="line" className="h-auto w-full justify-start rounded-none border-b px-2">
          {files.map((file) => (
            <TabsTrigger
              key={file.name}
              value={file.name}
              className="flex-none px-3 py-3 text-[13px]"
            >
              {file.name}
            </TabsTrigger>
          ))}
        </TabsList>
        {files.map((file) => (
          <TabsContent
            key={file.name}
            value={file.name}
            className="[&_.line]:pr-6 [&_.line]:[counter-increment:line] [&_.line]:before:mr-6 [&_.line]:before:inline-block [&_.line]:before:w-10 [&_.line]:before:text-right [&_.line]:before:text-muted-foreground/50 [&_.line]:before:content-[counter(line)] [&_.line[data-highlighted]]:bg-[color-mix(in_srgb,var(--presentation-selection,#1a73e8)_10%,transparent)] [&_.line[data-highlighted]]:shadow-[inset_2px_0_0_var(--presentation-selection,#1a73e8)]"
          >
            {file.rendered}
          </TabsContent>
        ))}
      </Tabs>
      <div className="flex flex-col gap-6">
        <SectionHeading eyebrow="Composable API">Primitives, not a product.</SectionHeading>
        <p className="text-[17px]/[26px] text-muted-foreground">
          Assemble the parts like shadcn/ui. Each one is unstyled and accessible, so the editor
          matches your app instead of the other way round.
        </p>
        <dl className="flex flex-col pt-4 text-sm">
          {PRIMITIVES.map((primitive) => (
            <div
              key={primitive.name}
              className="flex items-center justify-between gap-4 border-t border-foreground/15 py-3"
            >
              <dt className="font-mono">{primitive.name}</dt>
              <dd className="text-right text-muted-foreground">{primitive.description}</dd>
            </div>
          ))}
        </dl>
      </div>
    </Section>
  );
}

function Architecture() {
  return (
    <Section className="flex flex-col items-center gap-16">
      <div className="flex flex-col items-center gap-6 text-center">
        <SectionHeading eyebrow="Architecture">Two layers. React is optional.</SectionHeading>
        <p className="max-w-[35rem] text-[17px]/[26px] text-balance text-muted-foreground">
          The core parses, renders, edits, and saves. The React layer adds a store and the
          primitives on top.
        </p>
      </div>
      <div className="flex w-full flex-col items-center">
        <Layer name="@diceui/pptx" tag="React" description="Store built on useSyncExternalStore">
          <div className="grid gap-3 md:grid-cols-3">
            {REACT_LAYER.map((item) => (
              <div
                key={item.name}
                className="flex flex-col gap-2 rounded-lg border bg-card px-4.5 py-4"
              >
                <span className="flex items-center gap-2 text-[15px] font-semibold">
                  <item.icon className="size-4" />
                  {item.name}
                </span>
                <span className="font-mono text-[13px] text-muted-foreground">
                  {item.description}
                </span>
              </div>
            ))}
          </div>
        </Layer>
        <div className="flex flex-col items-center">
          <span className="h-4 w-px bg-border" />
          <span className="rounded-full border px-2.5 py-0.5 font-mono text-[11px] text-muted-foreground">
            wraps
          </span>
          <span className="h-4 w-px bg-border" />
        </div>
        <Layer
          name="@diceui/pptx-core"
          tag="Framework-agnostic"
          description="Runs anywhere the DOM does"
        >
          <ol className="grid grid-cols-2 gap-3 md:grid-cols-3 lg:grid-cols-6 lg:gap-8">
            {CORE_STEPS.map((step, index) => (
              <li
                key={step.name}
                className="relative flex flex-col gap-3 rounded-lg border px-3.5 py-3"
              >
                {index > 0 && (
                  <ChevronRight className="absolute top-1/2 -left-5.75 hidden size-3.5 -translate-y-1/2 text-muted-foreground/50 lg:block" />
                )}
                <span className="flex items-center justify-between text-muted-foreground">
                  <step.icon className="size-4" />
                  <span className="font-mono text-[11px] text-muted-foreground/75">
                    {String(index + 1).padStart(2, "0")}
                  </span>
                </span>
                <span className="text-sm font-medium">{step.name}</span>
              </li>
            ))}
          </ol>
          <p className="font-mono text-xs text-muted-foreground/75">
            slide → layout → master → theme
          </p>
        </Layer>
      </div>
    </Section>
  );
}

function Footer() {
  return (
    <footer className="border-t px-4 py-7 text-center text-[13px] text-muted-foreground md:px-8">
      Built by{" "}
      <Link
        href={siteConfig.links.x}
        target="_blank"
        rel="noreferrer"
        className="font-medium underline underline-offset-4 hover:text-foreground"
      >
        sadman
      </Link>{" "}
      for{" "}
      <Link
        href={siteConfig.links.diceui}
        target="_blank"
        rel="noreferrer"
        className="font-medium underline underline-offset-4 hover:text-foreground"
      >
        Dice UI
      </Link>
      .
    </footer>
  );
}

function Section({ className, children }: { className?: string; children: React.ReactNode }) {
  return (
    <section className="border-t px-4 py-24 md:px-8 md:py-32">
      <div className={cn("mx-auto w-full max-w-300", className)}>{children}</div>
    </section>
  );
}

function SectionHeading({ eyebrow, children }: { eyebrow: string; children: React.ReactNode }) {
  return (
    <div className="flex flex-col gap-5">
      <p className="font-mono text-xs tracking-[0.1em] text-muted-foreground/75 uppercase">
        {eyebrow}
      </p>
      <h2 className="text-3xl font-semibold tracking-tight md:text-[44px]/[48px] md:tracking-[-1.4px]">
        {children}
      </h2>
    </div>
  );
}

function IconTile({ icon: Icon }: { icon: React.ComponentType<{ className?: string }> }) {
  return (
    <span className="flex size-9 items-center justify-center rounded-lg border bg-card">
      <Icon className="size-4" />
    </span>
  );
}

function Layer({
  name,
  tag,
  description,
  children,
}: {
  name: string;
  tag: string;
  description: string;
  children: React.ReactNode;
}) {
  return (
    <div className="flex w-full flex-col gap-6 rounded-xl border bg-card/40 p-6 md:p-8">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div className="flex items-center gap-3">
          <span className="font-mono text-lg font-medium">{name}</span>
          <span className="rounded-full border px-2 py-0.5 text-xs text-muted-foreground">
            {tag}
          </span>
        </div>
        <span className="text-sm text-muted-foreground/75">{description}</span>
      </div>
      {children}
    </div>
  );
}

function GitHubIcon() {
  return (
    <svg viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
      <path d="M12 .297c-6.63 0-12 5.373-12 12 0 5.303 3.438 9.8 8.205 11.385.6.113.82-.258.82-.577 0-.285-.01-1.04-.015-2.04-3.338.724-4.042-1.61-4.042-1.61C4.422 18.07 3.633 17.7 3.633 17.7c-1.087-.744.084-.729.084-.729 1.205.084 1.838 1.236 1.838 1.236 1.07 1.835 2.809 1.305 3.495.998.108-.776.417-1.305.76-1.605-2.665-.3-5.466-1.332-5.466-5.93 0-1.31.465-2.38 1.235-3.22-.135-.303-.54-1.523.105-3.176 0 0 1.005-.322 3.3 1.23.96-.267 1.98-.399 3-.405 1.02.006 2.04.138 3 .405 2.28-1.552 3.285-1.23 3.285-1.23.645 1.653.24 2.873.12 3.176.765.84 1.23 1.91 1.23 3.22 0 4.61-2.805 5.625-5.475 5.92.42.36.81 1.096.81 2.22 0 1.606-.015 2.896-.015 3.286 0 .315.21.69.825.57C20.565 22.092 24 17.592 24 12.297c0-6.627-5.373-12-12-12" />
    </svg>
  );
}
