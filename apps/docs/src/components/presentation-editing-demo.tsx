"use client";

import * as React from "react";

import { useCreatePresentationStore } from "@diceui/pptx";

import { PresentationEditor } from "@/components/presentation-editor";
import { DEMO_DECK_PATH } from "@/lib/constants";

/** The playground's editor, on the sample deck. */
export function PresentationEditingDemo() {
  const store = useCreatePresentationStore();

  React.useEffect(() => {
    fetch(DEMO_DECK_PATH)
      .then((res) => {
        // fetch resolves on 404, so an unchecked body would reach the parser as
        // an error page rather than a deck.
        if (!res.ok) throw new Error(`${DEMO_DECK_PATH}: ${res.status}.`);
        return res.arrayBuffer();
      })
      // Editing and reordering both need the source package retained.
      .then((buffer) => store.load(buffer, { readOnly: false }))
      .catch(() => {
        // The editor shows load errors from `store.error`.
      });
  }, [store]);

  return <PresentationEditor store={store} className="not-prose h-[36rem] rounded-lg" />;
}
