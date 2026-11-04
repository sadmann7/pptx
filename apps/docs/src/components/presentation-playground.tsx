"use client";

import * as React from "react";

import { useCreatePresentationStore, usePresentation, useSlide } from "@diceui/pptx";
import { cn } from "@pptx/ui/lib/utils";

import { PresentationEditor } from "@/components/presentation-editor";
import type { SearchParams } from "@/types";

type PlaygroundLayout = "default" | "editing" | "compact";

function getPlaygroundLayout(value: Awaited<SearchParams>[number]): PlaygroundLayout {
  return value === "editing" || value === "compact" ? value : "default";
}

interface PresentationPlaygroundProps {
  searchParams: Promise<SearchParams>;
}

/**
 * The editor at full height. `?layout=editing` drops the debug bar, and
 * `?layout=compact` the toolbar as well.
 */
export function PresentationPlayground({ searchParams }: PresentationPlaygroundProps) {
  const store = useCreatePresentationStore();
  const layout = getPlaygroundLayout(React.use(searchParams).layout);

  return (
    <div className="mx-auto flex w-full max-w-(--fd-layout-width) flex-col p-4">
      <PresentationEditor
        store={store}
        className={cn("h-[calc(100dvh-(--spacing(22)))]")}
        header={layout === "default" ? <PresentationDebugToolbar /> : undefined}
        toolbar={layout !== "compact"}
      />
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
        {perf.renders} rendered · {avgFrameMs.toFixed(1)}ms avg · {perf.maxFrameMs.toFixed(1)}ms max
        · {perf.backlog} queued
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
