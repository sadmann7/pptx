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
          <Presentation.Thumbnails className="border-r border-border" style={{ width: 160 }} />
          <div className="flex flex-1 flex-col overflow-hidden">
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
  const { slide } = useSlide();

  React.useEffect(() => {
    if (slide && status === "ready" && slide.index === 0) {
      console.group(`[pptx] FINAL elements for slide 0 (${slide.elements.length} total)`);
      slide.elements.forEach((el, i) => {
        type AnyEl = {
          type: string;
          id: string;
          placeholder?: { type: string; idx: number };
          position: { x: number; y: number };
          size: { width: number; height: number };
          paragraphs?: { runs?: { type: string; text?: string }[] }[];
          src?: string;
        };
        const e = el as AnyEl;
        const text =
          e.type === "text"
            ? e.paragraphs
                ?.map((p) => p.runs?.map((r) => r.text).join(""))
                .join(" / ")
                .slice(0, 60)
            : e.type === "image"
              ? `src=${!!e.src}`
              : "";
        const posStr = `x=${e.position.x.toFixed(1)} y=${e.position.y.toFixed(1)}`;
        const sizeStr = `w=${e.size.width.toFixed(1)} h=${e.size.height.toFixed(1)}`;
        const phStr = e.placeholder ? `[${e.placeholder.type}/${e.placeholder.idx}]` : "";
        console.log(
          `  [${i}] ${e.type}${phStr} id="${e.id}" | ${posStr} | ${sizeStr}${text ? ` | "${text}"` : ""}`,
        );
      });
      console.groupEnd();
    }
  }, [slide, status]);

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
              {presentation.slideSize.width.toFixed(0)} × {presentation.slideSize.height.toFixed(0)}{" "}
              pt
            </strong>
          </span>
          {slide && (
            <span>
              elements: <strong className="text-foreground">{slide.elements.length}</strong>
            </span>
          )}
        </>
      )}
      {error && <span className="text-destructive">{error.message}</span>}
    </div>
  );
}
