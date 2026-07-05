"use client";

import * as React from "react";

import { useCreatePresentationStore, usePresentation, useSlide } from "@diceui/pptx";

import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import {
  Presentation,
  PresentationContent,
  PresentationError,
  PresentationLoading,
  PresentationProvider,
  PresentationSlide,
  PresentationThumbnailList,
  PresentationViewport,
} from "@/components/ui/presentation";

export default function PgPage() {
  const id = React.useId();
  const store = useCreatePresentationStore();

  async function onFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    store.load(file, { defaultSlideIndex: 0 });
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
        <Presentation className="flex-1">
          <PresentationThumbnailList />
          <PresentationContent>
            <PresentationLoading />
            <PresentationError />
            <PresentationViewport autoFit>
              <PresentationSlide />
            </PresentationViewport>
          </PresentationContent>
        </Presentation>
      </PresentationProvider>
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

/**
 * Temporary instrumentation readout: the thumbnail list dispatches a
 * `pptx:thumbnail-perf` CustomEvent after each render-queue drain frame
 * (dev builds only). Shows renders completed, avg/max frame cost, backlog.
 */
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
