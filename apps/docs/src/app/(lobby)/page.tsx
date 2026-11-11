import * as React from "react";

import { buttonVariants } from "@pptx/ui/components/button";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@pptx/ui/components/tabs";
import { cn } from "@pptx/ui/lib/utils";
import { highlight } from "fumadocs-core/highlight";
import JSZip from "jszip";
import {
  ArrowDown,
  ArrowRight,
  CaseSensitive,
  ChartColumn,
  Move,
  Shapes,
  Table,
  Type,
} from "lucide-react";
import Link from "next/link";
import { readFile } from "node:fs/promises";
import path from "node:path";

import { DemoDeckEditor, DemoDeckParts, DemoDeckProvider } from "@/components/landing-demo";
import type { DeckPart } from "@/components/landing-demo";
import { DEMO_DECK_PATH } from "@/lib/constants";
import { siteConfig } from "@/lib/site";

const FEATURES = [
  {
    icon: Shapes,
    title: "Faithful rendering",
    description: "Shapes, fills, shadows, and images, inherited through layouts and themes.",
  },
  {
    icon: ChartColumn,
    title: "Charts",
    description: "Bar, line, area, pie, scatter, and more, drawn from the deck's data.",
  },
  {
    icon: Move,
    title: "Direct manipulation",
    description: "Move, resize, rotate, and snap, with PowerPoint's shortcuts.",
  },
  {
    icon: Type,
    title: "Inline text editing",
    description: "Edit text in place and format it with useSelection.",
  },
  {
    icon: Table,
    title: "Table editing",
    description: "Add rows and columns, merge cells, and style borders.",
  },
  {
    icon: CaseSensitive,
    title: "Embedded fonts",
    description: "Fonts packed in the file load, so slides keep their type.",
  },
];

const COVERAGE_STATS = [
  { value: "187", label: "preset geometries" },
  { value: "74", label: "built-in table styles" },
  { value: "9", label: "chart types" },
  { value: "6", label: "image formats" },
];

const COVERAGE = [
  {
    area: "Shapes and geometry",
    notes: "All 187 presets, custom geometry, groups, and connectors",
    status: "supported",
  },
  {
    area: "Fills, lines, and effects",
    notes: "Solid, gradient, pattern, and picture fills, plus shadows",
    status: "supported",
  },
  {
    area: "Text",
    notes: "Bullets, autofit, vertical text, and embedded fonts",
    status: "supported",
  },
  { area: "Tables", notes: "Merges, borders, and 74 built-in styles", status: "supported" },
  { area: "Charts", notes: "Nine types; 3-D charts render flat", status: "partial" },
  {
    area: "SmartArt, EMF, and OLE",
    notes: "Drawn from the fallback stored in the file",
    status: "partial",
  },
  { area: "Transitions and animations", notes: "Slides render as static frames", status: "none" },
  { area: "Notes, comments, 3-D, and WMF", notes: "Kept on save, not rendered", status: "none" },
] as const;

const COVERAGE_STATUS = {
  supported: { label: "Supported", className: "bg-green-500" },
  partial: { label: "Partial", className: "bg-amber-500" },
  none: { label: "Not yet", className: "bg-muted-foreground" },
};

const EDITOR_CODE = `import { Presentation } from "@diceui/pptx";

export function Editor({ file }) {
  return (
    <Presentation.Root file={file}>
      <Presentation.ThumbnailList />
      <Presentation.Viewport>
        <Presentation.Slide>
          <Presentation.Selection />
        </Presentation.Slide>
      </Presentation.Viewport>
    </Presentation.Root>
  );
}`;

const TOOLBAR_CODE = `import { useSelection } from "@diceui/pptx";

export function Toolbar() {
  const { textStyle, setTextStyle } = useSelection();

  return (
    <div data-presentation-keep-selection="">
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

const ARCHITECTURE_PATHS = [
  {
    label: "With React",
    app: "Next.js, Vite",
    layer: "@diceui/pptx",
    tag: "Components and hooks",
    parts: ["Root", "Viewport", "Slide", "Selection"],
  },
  {
    label: "Without React",
    app: "Vue, Svelte, plain JS",
    layer: "Your bindings",
    tag: "Call the core directly",
    parts: ["PptxViewer", "renderSlide()", "applyEdit()", "writePptx()"],
  },
];

const CORE_STEPS = ["Parse", "Render", "Edit", "Save"];

export default function IndexPage() {
  return (
    <DemoDeckProvider>
      <main className="flex flex-col">
        <Hero />
        <RoundTrip />
        <Features />
        <Coverage />
        <ComposableApi />
        <Architecture />
        <Footer />
      </main>
    </DemoDeckProvider>
  );
}

function Hero() {
  return (
    <section className="flex flex-col items-center gap-10 px-5 pt-14 pb-14 md:gap-18 md:px-8 md:pt-28 md:pb-24">
      <div className="flex max-w-4xl flex-col items-center gap-5 text-center md:gap-6">
        <Link
          href="/docs/editing"
          className="flex animate-fade-up items-center gap-2 rounded-full border bg-card py-1 pr-3 pl-1 text-[13px] text-muted-foreground opacity-0 transition-colors [animation-fill-mode:forwards] hover:text-foreground"
        >
          <span className="rounded-full bg-primary px-2 py-0.5 text-[11px] font-semibold text-primary-foreground">
            New
          </span>
          <span className="sm:hidden">Table editing in 0.4</span>
          <span className="max-sm:hidden">New in 0.4: table editing</span>
          <ArrowRight className="size-3.5" />
        </Link>
        <h1 className="animate-fade-up text-4xl font-semibold tracking-tight text-balance opacity-0 [animation-delay:0.1s] [animation-fill-mode:forwards] md:text-[68px]/[71px] md:tracking-[-2.6px]">
          Open, edit, and save .pptx without leaving the browser
        </h1>
        <p className="max-w-[40rem] animate-fade-up text-lg text-balance text-muted-foreground opacity-0 [animation-delay:0.2s] [animation-fill-mode:forwards] md:text-[19px]/[29px]">
          Composable React primitives for viewing and editing PowerPoint decks.
        </p>
        <div className="grid w-full animate-fade-up grid-cols-2 gap-3 pt-2 opacity-0 [animation-delay:0.3s] [animation-fill-mode:forwards] sm:flex sm:w-auto">
          <CtaButtons />
        </div>
      </div>
      <div className="flex w-full justify-center bg-[radial-gradient(ellipse_45%_40%_at_50%_55%,color-mix(in_srgb,var(--presentation-selection,#1a73e8)_8%,transparent),transparent)] md:py-0">
        <DemoDeckEditor
          defaultSelectedNodeName="mat-cover-title"
          className="h-[22rem] w-full max-w-300 rounded-xl shadow-2xl shadow-black/40 md:h-[40.75rem]"
        />
      </div>
    </section>
  );
}

function CtaButtons() {
  return (
    <>
      <Link
        href="/docs"
        className={cn(buttonVariants({ size: "lg" }), "h-10 rounded-md px-4 font-semibold")}
      >
        Get started
        <ArrowRight className="size-3.75" />
      </Link>
      <Link
        href="/docs/components"
        className={cn(
          buttonVariants({ variant: "outline", size: "lg" }),
          "h-10 rounded-md bg-card px-4 dark:border-border dark:bg-card",
        )}
      >
        View components
      </Link>
    </>
  );
}

async function RoundTrip() {
  const parts = await getDeckParts();

  return (
    <Section className="grid items-center gap-8 lg:grid-cols-[minmax(0,26rem)_minmax(0,1fr)] lg:gap-30">
      <div className="flex flex-col gap-6">
        <SectionHeading eyebrow="Lossless save">
          Edit one slide.
          <br />
          Save the same file.
        </SectionHeading>
        <p className="text-[17px]/[26px] text-muted-foreground">
          Only the parts you edit are rewritten. Everything else is saved byte-for-byte.
        </p>
        <p className="flex items-center gap-2 font-mono text-xs text-muted-foreground">
          <span className="size-1.75 rounded-full bg-green-500" />
          Live: edit the slide above<span className="max-sm:hidden"> and this list updates</span>
        </p>
      </div>
      <DemoDeckParts fileName={path.basename(DEMO_DECK_PATH)} parts={parts} />
    </Section>
  );
}

async function getDeckParts(): Promise<DeckPart[]> {
  const deck = await readFile(path.join(process.cwd(), "public", DEMO_DECK_PATH));
  const zip = await JSZip.loadAsync(deck);
  const files = Object.values(zip.files).filter((file) => !file.dir);

  return Promise.all(
    files.map(async (file) => ({
      path: file.name,
      size: (await file.async("uint8array")).length,
    })),
  );
}

function Features() {
  return (
    <Section className="flex flex-col gap-8 md:gap-16">
      <SectionHeading eyebrow="Everything a deck needs">
        Renders like PowerPoint.
        <br />
        Edits like PowerPoint.
      </SectionHeading>
      <div className="grid gap-px overflow-hidden rounded-xl border bg-border sm:grid-cols-2 lg:grid-cols-3">
        {FEATURES.map((feature) => (
          <div
            key={feature.title}
            className="flex gap-3.5 bg-background p-4 sm:flex-col sm:gap-1.5 sm:p-7 sm:pb-7.5"
          >
            <IconTile icon={feature.icon} />
            <div className="flex flex-col gap-1 sm:mt-4 sm:gap-1.5">
              <h3 className="text-[15px] font-semibold tracking-[-0.2px] sm:text-[17px]">
                {feature.title}
              </h3>
              <p className="text-sm/5 text-muted-foreground sm:text-[15px]/[23px]">
                {feature.description}
              </p>
            </div>
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
              className="overflow-x-auto py-4 text-xs/5.5 sm:text-[13.5px]/6 [&_code]:grid [&_code]:[counter-reset:line] [&_span]:text-(--shiki-light) dark:[&_span]:text-(--shiki-dark)"
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
    <Section className="grid items-center gap-8 lg:grid-cols-[minmax(0,1fr)_minmax(0,27.5rem)] lg:gap-24">
      <Tabs
        defaultValue={files[0]?.name}
        className="dark min-w-0 gap-0 overflow-hidden rounded-xl border bg-card font-mono text-foreground"
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
            className="[&_.line]:pr-6 [&_.line]:[counter-increment:line] [&_.line]:before:mr-4 [&_.line]:before:inline-block [&_.line]:before:w-8 [&_.line]:before:text-right [&_.line]:before:text-neutral-700 [&_.line]:before:content-[counter(line)] sm:[&_.line]:before:mr-6 sm:[&_.line]:before:w-10 [&_.line[data-highlighted]]:bg-[color-mix(in_srgb,var(--presentation-selection,#1a73e8)_10%,transparent)] [&_.line[data-highlighted]]:shadow-[inset_2px_0_0_var(--presentation-selection,#1a73e8)]"
          >
            {file.rendered}
          </TabsContent>
        ))}
      </Tabs>
      <div className="flex flex-col gap-6 max-lg:order-first">
        <SectionHeading eyebrow="Composable API">Compose your editor.</SectionHeading>
        <p className="text-[17px]/[26px] text-muted-foreground">
          Unstyled, accessible parts you arrange and style like shadcn/ui.
        </p>
        <dl className="flex flex-col pt-4 text-sm">
          {PRIMITIVES.map((primitive) => (
            <div
              key={primitive.name}
              className="flex justify-between gap-x-4 gap-y-1 border-t py-3 max-sm:flex-col sm:items-center"
            >
              <dt className="font-mono">{primitive.name}</dt>
              <dd className="text-muted-foreground sm:text-right">{primitive.description}</dd>
            </div>
          ))}
        </dl>
      </div>
    </Section>
  );
}

function Coverage() {
  return (
    <Section className="flex flex-col gap-8 md:gap-16">
      <SectionHeading eyebrow="OOXML coverage">
        Measured against
        <br />
        ECMA-376.
      </SectionHeading>
      <div className="flex flex-col gap-8 md:gap-16">
        <dl className="grid grid-cols-2 border-y md:grid-cols-4">
          {COVERAGE_STATS.map((stat, index) => (
            <div
              key={stat.label}
              className={cn(
                "flex flex-col-reverse gap-1.5 py-5 md:py-7",
                index % 2 === 0 ? "pr-5 max-md:border-r" : "pl-5",
                index < 2 && "max-md:border-b",
                index > 0 && "md:border-l md:pl-7",
                index < COVERAGE_STATS.length - 1 && "md:pr-7",
              )}
            >
              <dt className="text-[13px] text-muted-foreground md:text-[15px] md:font-medium md:text-foreground">
                {stat.label}
              </dt>
              <dd className="text-[34px]/none font-semibold tracking-[-1.2px] md:text-5xl/none md:tracking-[-1.8px]">
                {stat.value}
              </dd>
            </div>
          ))}
        </dl>
        <div className="overflow-hidden rounded-xl border">
          <div className="grid grid-cols-[15rem_minmax(0,1fr)_7.5rem] gap-6 border-b bg-card px-6 py-3 font-mono text-[11px] tracking-[0.1em] text-muted-foreground uppercase max-md:hidden">
            <span>Area</span>
            <span>What renders</span>
            <span>Status</span>
          </div>
          <ul>
            {COVERAGE.map((row) => {
              const status = COVERAGE_STATUS[row.status];

              return (
                <li
                  key={row.area}
                  className="grid grid-cols-[minmax(0,1fr)_auto] gap-x-6 gap-y-1 border-b px-4 py-3.5 last:border-b-0 md:grid-cols-[15rem_minmax(0,1fr)_7.5rem] md:items-center md:px-6 md:py-4"
                >
                  <span className="text-sm font-medium md:text-[15px]">{row.area}</span>
                  <span className="text-[13px]/5 text-muted-foreground max-md:order-last max-md:col-span-2 md:text-sm">
                    {row.notes}
                  </span>
                  <span
                    className={cn(
                      "flex items-center gap-2 font-mono text-[11px] md:text-xs",
                      row.status === "none" && "text-muted-foreground",
                    )}
                  >
                    <span className={cn("size-1.5 rounded-full", status.className)} />
                    {status.label}
                  </span>
                </li>
              );
            })}
          </ul>
        </div>
      </div>
    </Section>
  );
}

function Architecture() {
  return (
    <Section className="flex flex-col gap-8 md:items-center md:gap-16">
      <div className="flex flex-col gap-3.5 md:items-center md:gap-6 md:text-center">
        <SectionHeading eyebrow="Architecture">React is optional.</SectionHeading>
        <p className="max-w-[40rem] text-[17px]/[26px] text-balance text-muted-foreground">
          The engine runs on its own. React adds the components.
        </p>
      </div>
      <div className="grid w-full gap-4 md:grid-cols-2">
        {ARCHITECTURE_PATHS.map((item) => (
          <div
            key={item.label}
            className="flex flex-col rounded-xl border bg-background p-4 md:p-6"
          >
            <p className="pb-4 font-mono text-[11px] tracking-[0.1em] text-muted-foreground uppercase">
              {item.label}
            </p>
            <Tier name="Your app" tag={item.app} className="h-11 font-sans font-medium md:h-14" />
            <TierConnector />
            <Tier
              name={item.layer}
              tag={item.tag}
              className={item.layer.startsWith("@") ? undefined : "font-sans"}
            >
              {item.parts.map((part) => (
                <Chip key={part}>{part}</Chip>
              ))}
            </Tier>
            <TierConnector />
            <Tier name="@diceui/pptx-core" tag="The engine">
              {CORE_STEPS.map((step, index) => (
                <React.Fragment key={step}>
                  {index > 0 && <ArrowRight className="size-3 text-muted-foreground" />}
                  <Chip>{step}</Chip>
                </React.Fragment>
              ))}
            </Tier>
          </div>
        ))}
      </div>
    </Section>
  );
}

function Tier({
  name,
  tag,
  className,
  children,
}: {
  name: string;
  tag: string;
  className?: string;
  children?: React.ReactNode;
}) {
  return (
    <div
      className={cn(
        "flex flex-col justify-center gap-2 rounded-lg border bg-muted/50 px-3.5 font-mono md:gap-2.5 md:px-5",
        children ? "min-h-19 py-3 md:min-h-24" : undefined,
        className,
      )}
    >
      <div className="flex items-center justify-between gap-3">
        <span className="text-[15px] font-medium">{name}</span>
        <span className="text-right font-sans text-xs font-normal text-muted-foreground">
          {tag}
        </span>
      </div>
      {children && <div className="flex flex-wrap items-center gap-1.5">{children}</div>}
    </div>
  );
}

function TierConnector() {
  return (
    <div className="flex h-7 items-center justify-center text-muted-foreground">
      <ArrowDown className="size-3.5" />
    </div>
  );
}

function Chip({ children }: { children: React.ReactNode }) {
  return (
    <span className="rounded-sm border px-2 py-1 font-mono text-xs font-medium">{children}</span>
  );
}

function Footer() {
  return (
    <footer className="border-t px-5 py-6 text-center text-[13px] text-muted-foreground md:px-8">
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
    <section className="px-5 py-14 md:px-8 md:py-24">
      <div className={cn("mx-auto w-full max-w-300", className)}>{children}</div>
    </section>
  );
}

function SectionHeading({ eyebrow, children }: { eyebrow: string; children: React.ReactNode }) {
  return (
    <div className="flex flex-col gap-3.5 md:gap-5">
      <p className="font-mono text-xs tracking-[0.1em] text-muted-foreground uppercase">
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
    <span className="flex size-9 shrink-0 items-center justify-center rounded-lg border bg-muted">
      <Icon className="size-4" />
    </span>
  );
}
