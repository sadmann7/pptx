/**
 * Mounts and unmounts a viewer that either owns its store (`Presentation.Root`
 * with `file`) or gets one from `useCreatePresentationStore`.
 */
import * as React from "react";
import { flushSync } from "react-dom";

import { Presentation, useCreatePresentationStore } from "@diceui/pptx";
import { createRoot } from "react-dom/client";

const root = createRoot(document.getElementById("lifecycle-root")!);

function onLoad() {
  window.__viewerLoaded = true;
}

function onError(error: unknown) {
  window.__renderError = String(error);
}

function Slide() {
  return (
    <Presentation.Viewport>
      <Presentation.Slide />
    </Presentation.Viewport>
  );
}

function RootViewer({ file }: { file: ArrayBuffer }) {
  return (
    <Presentation.Root file={file} onLoad={onLoad} onError={onError}>
      <Slide />
    </Presentation.Root>
  );
}

function HookViewer({ file }: { file: ArrayBuffer }) {
  const store = useCreatePresentationStore();

  React.useEffect(() => {
    store.load(file).then(onLoad, (error: unknown) => {
      if (error instanceof DOMException && error.name === "AbortError") return;
      onError(error);
    });
  }, [store, file]);

  return (
    <Presentation.Provider store={store}>
      <Presentation.Root>
        <Slide />
      </Presentation.Root>
    </Presentation.Provider>
  );
}

window.__mountViewer = async (name: string, store: ViewerStore = "root") => {
  const file = await fetch(name).then((response) => response.arrayBuffer());
  const Viewer = store === "hook" ? HookViewer : RootViewer;
  window.__viewerLoaded = false;
  flushSync(() => root.render(<Viewer file={file} />));
};

window.__unmountViewer = () => {
  flushSync(() => root.render(null));
};
