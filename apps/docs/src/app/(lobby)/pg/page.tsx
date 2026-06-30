"use client";

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
import { usePresentation, useSlide } from "@diceui/pptx";
import * as React from "react";

export default function PgPage() {
  const [file, setFile] = React.useState<File | null>(null);

  return (
    <div className="flex flex-col bg-background text-foreground">
      <div className="flex items-center gap-3 border-b border-border px-3 py-2">
        <label className="text-sm font-medium text-muted-foreground">Open PPTX</label>
        <input
          type="file"
          accept=".pptx"
          className="text-sm"
          onChange={(event) => {
            const f = event.target.files?.[0];
            if (f) setFile(f);
          }}
        />
      </div>
      <Presentation
        file={file}
        className="max-h-[calc(100dvh-100px)] flex-1"
        onError={(error) => console.error("[pptx] parse error:", error)}
      >
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
