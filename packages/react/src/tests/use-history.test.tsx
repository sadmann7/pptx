// @vitest-environment jsdom
/**
 * `useHistory` tests.
 *
 * Runs under jsdom (not the default happy-dom) because happy-dom's XML parser
 * drops namespaced attributes (e.g. `r:id` on `p:sldId`), which the slide-level
 * edit operations these tests drive depend on.
 */
import { act, render } from "@testing-library/react";
import { beforeAll, describe, expect, it } from "vitest";

import type { UseHistoryResult } from "../context";
import { useHistory } from "../context";
import { Presentation } from "../index";
import type { Store } from "../store";
import { createStore } from "../store";
import { buildMinimalPptx } from "./minimal-pptx";

let fixture: ArrayBuffer;

beforeAll(async () => {
  fixture = await buildMinimalPptx();
});

async function renderHistory(): Promise<{ store: Store; history: () => UseHistoryResult }> {
  const store = createStore();
  // embedFonts off: the fixture has none, and the font pipeline needs workers.
  await store.load(fixture, { readOnly: false, embedFonts: false });

  const latest: { history?: UseHistoryResult } = {};

  function Probe() {
    latest.history = useHistory();
    return null;
  }

  render(
    <Presentation.Provider store={store}>
      <Probe />
    </Presentation.Provider>,
  );

  return { store, history: () => latest.history! };
}

describe("useHistory", () => {
  it("tracks availability and the dirty flag across edit, undo, and redo", async () => {
    const { store, history } = await renderHistory();
    const slideId = store.getState().presentation!.slides[0].id;

    expect(history().canUndo).toBe(false);
    expect(history().canRedo).toBe(false);
    expect(history().isDirty).toBe(false);

    await act(async () => {
      await store.edit({ type: "duplicateSlide", slideId });
    });
    expect(history().canUndo).toBe(true);
    expect(history().canRedo).toBe(false);
    expect(history().isDirty).toBe(true);

    // Undoing back to the loaded state is not a change worth saving.
    act(() => {
      history().undo();
    });
    expect(history().canUndo).toBe(false);
    expect(history().canRedo).toBe(true);
    expect(history().isDirty).toBe(false);

    await act(async () => {
      await history().redo();
    });
    expect(history().canUndo).toBe(true);
    expect(history().canRedo).toBe(false);
    expect(history().isDirty).toBe(true);
  });

  it("clears the dirty flag once the deck is saved", async () => {
    const { store, history } = await renderHistory();
    const slideId = store.getState().presentation!.slides[0].id;

    await act(async () => {
      await store.edit({ type: "duplicateSlide", slideId });
    });
    expect(history().isDirty).toBe(true);

    // `save()` moves the clean depth without writing to store state, so this is
    // the update a state-only subscription would miss.
    await act(async () => {
      await store.save();
    });
    expect(history().isDirty).toBe(false);
    expect(history().canUndo).toBe(true);
  });

  it("throws when used outside a Presentation tree", () => {
    function Naked() {
      useHistory();
      return null;
    }
    expect(() => render(<Naked />)).toThrow(/must be used within/);
  });
});
