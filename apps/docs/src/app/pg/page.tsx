"use client";

import { usePresentation, useSlide } from "@diceui/pptx";
import {
  Presentation,
  PresentationLoading,
  PresentationError,
  PresentationThumbnailList,
  PresentationViewport,
  PresentationSlide,
} from "@/components/ui/presentation";
import * as React from "react";
import { ThemeSwitch } from "fumadocs-ui/layouts/shared/slots/theme-switch";

export default function PgPage() {
  const [file, setFile] = React.useState<File | null>(null);

  return (
    <div className="flex h-dvh flex-col bg-background text-foreground">
      <ThemeSwitch className="absolute top-2 right-2" />
      <div className="flex items-center gap-3 border-b border-border px-3 py-2">
        <label className="text-sm font-medium text-muted-foreground">Open PPTX</label>
        <input
          type="file"
          accept=".pptx"
          className="text-sm"
          onChange={(e) => {
            const f = e.target.files?.[0];
            if (f) setFile(f);
          }}
        />
      </div>
      <Presentation file={file} onError={(e) => console.error("[pptx] parse error:", e)}>
        <PresentationDebug />
        <div className="flex flex-1 overflow-hidden">
          <PresentationThumbnailList className="border-r border-border" />
          <div className="relative flex flex-1 flex-col overflow-hidden">
            <PresentationLoading />
            <PresentationError />
            <PresentationViewport className="flex-1" autoFit autoFitPadding={32}>
              <PresentationSlide />
            </PresentationViewport>
          </div>
        </div>
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
