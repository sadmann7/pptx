"use client";

import * as React from "react";

import { useCreatePresentationStore, usePresentation, useSlide } from "@diceui/pptx";

import {
  Presentation,
  PresentationBody,
  PresentationContent,
  PresentationError,
  PresentationLoading,
  PresentationSlide,
  PresentationThumbnailList,
  PresentationViewport,
} from "@/components/ui/presentation";

export default function PgPage() {
  const store = useCreatePresentationStore();
  const [defaultSlideIndex, setDefaultSlideIndex] = React.useState(2);

  function onFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    store.load(file, { defaultSlideIndex });
  }

  return (
    <div className="flex flex-col">
      <div className="flex items-center gap-3 border-b border-border px-3 py-2">
        <label className="text-sm font-medium text-muted-foreground">Open PPTX</label>
        <input type="file" accept=".pptx" className="text-sm" onChange={onFileChange} />
        <label className="ml-auto flex items-center gap-1.5 text-sm text-muted-foreground">
          Start slide
          <input
            type="number"
            min={0}
            value={defaultSlideIndex}
            onChange={(e) => setDefaultSlideIndex(Math.max(0, Number(e.target.value)))}
            className="w-14 rounded border border-border bg-background px-2 py-0.5 text-foreground"
          />
        </label>
      </div>
      <Presentation store={store} className="max-h-[calc(100dvh-100px)] flex-1">
        <PresentationDebug />
        <PresentationBody>
          <PresentationThumbnailList />
          <PresentationContent>
            <PresentationLoading />
            <PresentationError />
            <PresentationViewport autoFit autoFitPadding={32}>
              <PresentationSlide />
            </PresentationViewport>
          </PresentationContent>
        </PresentationBody>
      </Presentation>
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
        </>
      )}
      {error && <span className="text-destructive">{error.message}</span>}
    </div>
  );
}
