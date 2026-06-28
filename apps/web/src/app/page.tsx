"use client";

import { Presentation, usePresentation, useSlide } from "@pptx/react";
import * as React from "react";

export default function IndexPage() {
  const [file, setFile] = React.useState<File | null>(null);

  return (
    <div className="flex h-dvh flex-col bg-background text-foreground">
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
      <Presentation.Root file={file} onError={(e) => console.error("[pptx] parse error:", e)}>
        <PresentationDebug />
        <div className="flex flex-1 overflow-hidden">
          <Presentation.ThumbnailList
            className="border-r border-border"
            thumbWidth={160}
            style={{ width: 160 }}
          >
            {({ slides }) =>
              slides.map((slide) => (
                <Presentation.ThumbnailItem key={slide.id} slideId={slide.id}>
                  <Presentation.ThumbnailItemPreview />
                  <Presentation.ThumbnailItemIndex className="font-mono text-muted-foreground" />
                </Presentation.ThumbnailItem>
              ))
            }
          </Presentation.ThumbnailList>
          <div className="relative flex flex-1 flex-col overflow-hidden">
            <Presentation.Loading className="absolute inset-0 z-10 flex items-center justify-center bg-background/80 text-sm text-muted-foreground">
              {(progress) => <span>Loading… {progress}%</span>}
            </Presentation.Loading>
            <Presentation.Error className="absolute inset-0 z-10 flex items-center justify-center bg-background/80 text-sm text-destructive">
              {(err) => <span>{err.message}</span>}
            </Presentation.Error>
            <Presentation.Viewport className="flex-1" autoFit autoFitPadding={32}>
              <Presentation.Slide />
            </Presentation.Viewport>
          </div>
        </div>
      </Presentation.Root>
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
