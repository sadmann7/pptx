"use client";

import * as React from "react";

import { useCreatePresentationStore } from "@diceui/pptx";
import { Input } from "@pptx/ui/components/input";
import { Label } from "@pptx/ui/components/label";
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
} from "@pptx/ui/components/presentation";

import { DEMO_DECK_PATH } from "@/lib/constants";

export function PresentationDemo() {
  const id = React.useId();
  const store = useCreatePresentationStore();

  React.useEffect(() => {
    fetch(DEMO_DECK_PATH)
      .then((res) => {
        // fetch resolves on 404, so an unchecked body would reach the parser as
        // an error page rather than a deck.
        if (!res.ok) throw new Error(`${DEMO_DECK_PATH}: ${res.status}.`);
        return res.arrayBuffer();
      })
      .then((buffer) => store.load(buffer))
      .catch(() => {
        // Fail silently to avoid blocking the main thread.
      });
    // eslint-disable-next-line react-hooks/exhaustive-deps -- store is a stable ref, intentionally omitted from deps
  }, []);

  async function onFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      await store.load(file, { readOnly: false });
    } catch {
      // Fail silently because `load()` already wrote the failure to `store.error`.
    }
  }

  return (
    <div className="not-prose flex h-100 flex-col overflow-hidden rounded-lg border">
      <div className="flex items-center gap-2 border-b p-1.5">
        <Label htmlFor={`${id}-file`} className="sr-only">
          Open .pptx
        </Label>
        <Input
          id={`${id}-file`}
          type="file"
          accept=".pptx"
          className="h-8 max-w-xs text-xs"
          onChange={onFileChange}
        />
      </div>
      <PresentationProvider store={store}>
        <Presentation className="min-h-0 flex-1">
          <PresentationThumbnailList className="p-1.5" />
          <PresentationContent>
            <PresentationLoading />
            <PresentationError />
            <PresentationViewport>
              <PresentationSlide>
                <PresentationSelection undoRedoShortcuts />
              </PresentationSlide>
            </PresentationViewport>
          </PresentationContent>
        </Presentation>
      </PresentationProvider>
    </div>
  );
}
