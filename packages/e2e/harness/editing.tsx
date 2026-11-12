/**
 * One editable slide with a formatting toolbar outside it, wired the way an app
 * uses `useSelection`.
 *
 * The toolbar sits in `Presentation.Controls` with an unmarked "popover" it portals;
 * a separate strip carries only `data-presentation-controls`, and the "outside"
 * strip is neither. The spec reads the model through `__editStore`.
 *
 * Query params:
 *   file  deck to load, served from fixtures/ (default basic.pptx)
 */
import * as React from "react";
import { createPortal } from "react-dom";

import {
  Presentation,
  useCreatePresentationStore,
  useSelection,
} from "@diceui/pptx";
import { createRoot } from "react-dom/client";

const file = new URLSearchParams(location.search).get("file") ?? "basic.pptx";

function Toolbar() {
  const {
    mode,
    nodeIds,
    cell,
    cellRange,
    textStyle,
    setTextStyle,
    setParagraphStyle,
    insertTextBox,
    insertTable,
    changeTable,
    setTableCellStyle,
    startDrawing,
  } = useSelection();

  return (
    <Presentation.Controls style={{ display: "flex", gap: 4, padding: 4 }}>
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
      <button
        type="button"
        data-testid="text-box"
        onClick={() => void insertTextBox()}
      >
        Text box
      </button>
      <button
        type="button"
        data-testid="draw-text-box"
        onClick={() => startDrawing({ type: "textBox" })}
      >
        Draw text box
      </button>
      <button
        type="button"
        data-testid="draw-ellipse"
        onClick={() => startDrawing({ type: "shape", geometry: "ellipse" })}
      >
        Draw ellipse
      </button>
      <button
        type="button"
        data-testid="insert-table"
        onClick={() => void insertTable(2, 3)}
      >
        Table
      </button>
      <button
        type="button"
        data-testid="merge-cells"
        onClick={() => void changeTable({ type: "mergeCells" })}
      >
        Merge
      </button>
      <button
        type="button"
        data-testid="cell-fill"
        onClick={() => void setTableCellStyle({ fill: "#FF0000" })}
      >
        Fill
      </button>
      <output data-testid="state">
        {JSON.stringify({
          mode,
          nodeIds,
          cell,
          cellRange,
          bold: textStyle?.bold ?? null,
        })}
      </output>
      {createPortal(
        <div style={{ position: "fixed", right: 8, bottom: 8 }}>
          <button
            type="button"
            data-testid="swatch"
            onClick={() => void setTextStyle({ color: "00FF00" })}
          >
            Green
          </button>
          <input data-testid="popover-input" aria-label="Popover field" />
        </div>,
        document.body,
      )}
    </Presentation.Controls>
  );
}

function MarkedStrip() {
  const { setTextStyle } = useSelection();

  return (
    <div data-presentation-controls="" style={{ padding: 4 }}>
      <button
        type="button"
        data-testid="marked-bold"
        onClick={() => void setTextStyle({ bold: true })}
      >
        Bold
      </button>
    </div>
  );
}

function App() {
  const store = useCreatePresentationStore();

  React.useEffect(() => {
    window.__editStore = store;
    fetch(`/${file}`)
      .then((response) => response.arrayBuffer())
      .then((buffer) =>
        store.load(buffer, { readOnly: false, embedFonts: false }),
      )
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
      <MarkedStrip />
      <div data-testid="outside" style={{ height: 40, background: "#eee" }} />
      <div
        data-testid="outside-stopped"
        onPointerDown={(event) => event.stopPropagation()}
        style={{ height: 40, background: "#ddd" }}
      />
      <div style={{ position: "relative", padding: 16 }}>
        <Presentation.Slide>
          <Presentation.Selection undoRedoShortcuts />
        </Presentation.Slide>
      </div>
    </Presentation.Provider>
  );
}

createRoot(document.getElementById("editing-root")!).render(<App />);
