/**
 * Shared helpers for component tests: fixture loading and store-scoped
 * rendering. Test files mirror source files (viewport.tsx →
 * viewport.test.tsx); keep cross-file setup here instead of duplicating it.
 */
import * as React from "react";

import { render, type RenderResult } from "@testing-library/react";

import { Presentation } from "../index";
import type { Store } from "../store";
import { createStore } from "../store";
import { buildMinimalPptx } from "./minimal-pptx";

let fixturePromise: Promise<ArrayBuffer> | null = null;

export function loadFixture(): Promise<ArrayBuffer> {
  fixturePromise ??= buildMinimalPptx();
  return fixturePromise;
}

export async function loadedStore(): Promise<Store> {
  const store = createStore();
  await store.load(await loadFixture());
  return store;
}

export async function editableStore(): Promise<Store> {
  const store = createStore();
  // The fixture has no fonts, and the font pipeline needs workers.
  await store.load(await loadFixture(), { readOnly: false, embedFonts: false });
  return store;
}

export function withStore(store: Store, ui: React.ReactNode): RenderResult {
  return render(<Presentation.Provider store={store}>{ui}</Presentation.Provider>);
}
