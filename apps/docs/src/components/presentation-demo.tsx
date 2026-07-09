"use client";

import * as React from "react";

import { useCreatePresentationStore } from "@diceui/pptx";

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

export function PresentationDemo() {
  const id = React.useId();
  const store = useCreatePresentationStore();

  React.useEffect(() => {
    fetch("/assets/sample.pptx")
      .then((res) => res.arrayBuffer())
      .then((buf) => store.load(buf))
      .catch(() => {
        // Fail silently to avoid blocking the main thread
      });
    // eslint-disable-next-line react-hooks/exhaustive-deps -- store is a stable ref, intentionally omitted from deps
  }, []);

  function onFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    store.load(file, { readOnly: false });
  }

  return (
    <div className="not-prose flex h-[560px] flex-col overflow-hidden rounded-lg border">
      <div className="flex items-center gap-2 border-b px-3 py-2">
        <Label htmlFor={`${id}-file`} className="shrink-0 text-sm text-muted-foreground">
          Open .pptx
        </Label>
        <Input
          id={`${id}-file`}
          type="file"
          accept=".pptx"
          className="h-8 max-w-xs cursor-pointer text-xs"
          onChange={onFileChange}
        />
      </div>
      <PresentationProvider store={store}>
        <Presentation className="min-h-0 flex-1">
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
