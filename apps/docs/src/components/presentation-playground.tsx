"use client";

import * as React from "react";

import type { PresentationStore } from "@diceui/pptx";
import {
  useCreatePresentationStore,
  usePresentation,
  useSlide,
} from "@diceui/pptx";
import { cn } from "@pptx/ui/lib/utils";

import { PresentationEditor } from "@/components/presentation-editor";
import type { SearchParams } from "@/types";

type PlaygroundLayout = "default" | "editing" | "compact";

function getPlaygroundLayout(
  value: Awaited<SearchParams>[number],
): PlaygroundLayout {
  return value === "editing" || value === "compact" ? value : "default";
}

interface PresentationPlaygroundProps {
  searchParams: Promise<SearchParams>;
}

// `?layout=editing` drops the debug bar, `?layout=compact` the toolbar as well,
// and `?floatingToolbar` floats the toolbar above the selection.
export function PresentationPlayground({
  searchParams,
}: PresentationPlaygroundProps) {
  const store = useCreatePresentationStore();
  const params = React.use(searchParams);
  const layout = getPlaygroundLayout(params.layout);

  return (
    <div className="mx-auto flex w-full max-w-(--fd-layout-width) flex-col p-4">
      <PresentationEditor
        store={store}
        className={cn(
          "h-[calc(100dvh-(--spacing(22)))] md:h-[calc(100dvh-(--spacing(24)))]",
        )}
        header={layout === "default" ? <PresentationDebugToolbar /> : undefined}
        toolbar={layout !== "compact"}
        floatingToolbar={params.floatingToolbar !== undefined}
      />
      {process.env.NODE_ENV === "development" && params.data && (
        <DataToggle store={store} initial={params.data} />
      )}
    </div>
  );
}

const DATASETS = [
  {
    value: "demo",
    label: "Demo data",
    src: "/assets/after-the-needle-drops-mat.pptx",
  },
  { value: "worst", label: "Worst case", src: "/assets/worst-case/worst.pptx" },
  { value: "one", label: "One slide", src: "/assets/worst-case/one.pptx" },
  { value: "many", label: "300 slides", src: "/assets/worst-case/many.pptx" },
  { value: "empty", label: "Empty", src: null },
] as const;

type Dataset = (typeof DATASETS)[number]["value"];

// Dev-only, and only with `?data=`. Loads through `store.load`, the way a
// user's file arrives.
function DataToggle({
  store,
  initial,
}: {
  store: PresentationStore;
  initial: Awaited<SearchParams>[number];
}) {
  const [dataset, setDataset] = React.useState<Dataset | null>(
    () => DATASETS.find((option) => option.value === initial)?.value ?? null,
  );

  React.useEffect(() => {
    if (!dataset) return;
    const src = DATASETS.find((option) => option.value === dataset)?.src;
    if (!src) {
      store.reset();
      return;
    }
    let cancelled = false;
    fetch(src)
      .then((response) => response.blob())
      .then((blob) => {
        if (!cancelled)
          return store.load(blob, { defaultSlideIndex: 0, readOnly: false });
      })
      .catch(() => {});
    return () => {
      cancelled = true;
    };
  }, [dataset, store]);

  function select(value: Dataset) {
    setDataset(value);
    const url = new URL(window.location.href);
    url.searchParams.set("data", value);
    window.history.replaceState(null, "", url);
  }

  return (
    <div className="fixed bottom-4 left-1/2 z-50 flex -translate-x-1/2 gap-0.5 rounded-full border border-input bg-background p-0.5 font-sans text-xs shadow-sm">
      {DATASETS.map((option) => (
        <button
          key={option.value}
          type="button"
          aria-pressed={dataset === option.value}
          onClick={() => select(option.value)}
          className={cn(
            "rounded-full px-3 py-1 whitespace-nowrap text-muted-foreground transition-colors hover:bg-muted hover:text-foreground",
            dataset === option.value && "bg-muted text-foreground",
          )}
        >
          {option.label}
        </button>
      ))}
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
        {perf.renders} rendered · {avgFrameMs.toFixed(1)}ms avg ·{" "}
        {perf.maxFrameMs.toFixed(1)}ms max · {perf.backlog} queued
      </strong>
    </span>
  );
}

function PresentationDebugToolbar() {
  const { status, progress, error, presentation } = usePresentation();
  const { slide, index } = useSlide();

  return (
    <div className="flex flex-wrap items-center gap-x-3 gap-y-0.5 border-b border-border px-3 py-1 font-mono text-xs text-muted-foreground">
      <span>
        status: <strong className="text-foreground">{status}</strong>
      </span>
      {status === "loading" && <span>{progress}%</span>}
      {status === "ready" && presentation && (
        <>
          <span>
            slides:{" "}
            <strong className="text-foreground">
              {presentation.slides.length}
            </strong>
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
              <strong className="text-foreground">
                {slide.nodes.length} nodes
              </strong>
            </span>
          )}
          <ThumbnailPerfReadout />
        </>
      )}
      {error && <span className="text-destructive">{error.message}</span>}
    </div>
  );
}
