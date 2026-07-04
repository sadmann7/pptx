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
