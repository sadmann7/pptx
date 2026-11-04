/**
 * Editing harness: one editable slide with a formatting toolbar outside it,
 * the way an app wires `useSelection`, so editing.spec.ts can press real
 * toolbar buttons and check the selection and text editing survive.
 *
 * The toolbar and a portalled "popover" carry `data-pptx-keep-selection`; the
 * "outside" strip does not. The spec reads the model through `__editStore`.
 *
 * Query params:
 *   file  deck to load, served from fixtures/ (default basic.pptx)
 */
import * as React from "react";
import { createPortal } from "react-dom";

import { Presentation, useCreatePresentationStore, useSelection } from "@diceui/pptx";
import { createRoot } from "react-dom/client";

const file = new URLSearchParams(location.search).get("file") ?? "basic.pptx";

function Toolbar() {
  const { mode, nodeIds, textStyle, setTextStyle, setParagraphStyle, insertTextBox } =
    useSelection();

  return (
    <>
      <div data-pptx-keep-selection="" style={{ display: "flex", gap: 4, padding: 4 }}>
        <button
          type="button"
          data-testid="italic"
          onClick={() => void setTextStyle({ italic: true })}
        >
          Italic
        </button>
        <button
          type="button"
          data-testid="underline"
          onClick={() => void setTextStyle({ underline: true })}
        >
          Underline
        </button>
        <button
          type="button"
          data-testid="align-left"
          onClick={() => void setParagraphStyle({ align: "left" })}
        >
          Left
        </button>
        <button type="button" data-testid="text-box" onClick={() => void insertTextBox()}>
          Text box
        </button>
        <output data-testid="state">
          {JSON.stringify({ mode, nodeIds, bold: textStyle?.bold ?? null })}
        </output>
      </div>
      {createPortal(
        <div data-pptx-keep-selection="" style={{ position: "fixed", right: 8, bottom: 8 }}>
          <button
            type="button"
            data-testid="swatch"
            onClick={() => void setTextStyle({ color: "00FF00" })}
          >
            Green
          </button>
        </div>,
        document.body,
      )}
    </>
  );
}

function App() {
  const store = useCreatePresentationStore();

  React.useEffect(() => {
    window.__editStore = store;
    fetch(`/${file}`)
      .then((response) => response.arrayBuffer())
      .then((buffer) => store.load(buffer, { readOnly: false, embedFonts: false }))
      .then(() => {
        window.__editingReady = true;
      })
      .catch((error: unknown) => {
        window.__renderError = String(error);
      });
  }, [store]);

  return (
    <Presentation.Provider store={store}>
      <Toolbar />
      <div data-testid="outside" style={{ height: 40, background: "#eee" }} />
      <div style={{ position: "relative", padding: 16 }}>
        <Presentation.Slide>
          <Presentation.Selection undoRedoShortcuts />
        </Presentation.Slide>
      </div>
    </Presentation.Provider>
  );
}

createRoot(document.getElementById("editing-root")!).render(<App />);
