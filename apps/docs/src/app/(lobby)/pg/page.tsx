"use client";

import * as React from "react";

import type { PresentationData, PresentationStore } from "@diceui/pptx";
import { useCreatePresentationStore, usePresentation, useSlide } from "@diceui/pptx";

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import {
  Presentation,
  PresentationContent,
  PresentationError,
  PresentationLoading,
  PresentationProvider,
  PresentationSelection,
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
        <EditToolbar store={store} />
        <Presentation className="flex-1">
          <PresentationThumbnailList />
          <PresentationContent>
            <PresentationLoading />
            <PresentationError />
            <PresentationViewport autoFit autoFitPadding={10}>
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

/** Locate the first editable text run on the active slide. */
function findFirstTextRun(
  presentation: PresentationData,
  slideId: string,
): { nodeId: string; text: string } | null {
  const slide = presentation.slides.find((s) => s.id === slideId);
  if (!slide) return null;
  for (const node of slide.nodes) {
    if (node.nodeType !== "shape") continue;
    if (!("textBody" in node) || !node.textBody) continue;
    const run = node.textBody.paragraphs[0]?.runs[0];
    if (run) return { nodeId: node.id, text: run.text };
  }
  return null;
}

/**
 * Temporary editing testbed until the interactive editor UI exists:
 * exercises `store.edit()`, undo/redo, and `store.save()` end to end.
 */
function EditToolbar({ store }: { store: PresentationStore }) {
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

  function onEditText() {
    const presentation = store.getState().presentation;
    if (!presentation || !slideId) return;
    const target = findFirstTextRun(presentation, slideId);
    if (!target) {
      console.warn("[pg] no text run on this slide to edit");
      return;
    }
    const text = window.prompt("New text for the first run:", target.text);
    if (text === null) return;
    run(() =>
      store.edit({
        type: "setTextRun",
        slideId,
        nodeId: target.nodeId,
        paragraphIndex: 0,
        runIndex: 0,
        text,
      }),
    );
  }

  function onSave() {
    run(async () => {
      const bytes = await store.save();
      const blob = new Blob([bytes.slice().buffer as ArrayBuffer], {
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
      <Button size="sm" variant="outline" onClick={onEditText}>
        Edit text…
      </Button>
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
