/**
 * Shared helpers for component tests: fixture loading and store-scoped
 * rendering. Test files mirror source files (viewport.tsx →
 * viewport.test.tsx); keep cross-file setup here instead of duplicating it.
 */
import * as React from "react";

import { render } from "@testing-library/react";

import { Presentation } from "../index";
import type { Store } from "../store";
import { createStore } from "../store";
import { buildMinimalPptx } from "./minimal-pptx";

let fixturePromise: Promise<ArrayBuffer> | null = null;

/** The minimal deck fixture, built once and shared across test files. */
export function loadFixture(): Promise<ArrayBuffer> {
  fixturePromise ??= buildMinimalPptx();
  return fixturePromise;
}

/** A store with the minimal fixture fully loaded. */
export async function loadedStore(): Promise<Store> {
  const store = createStore();
  await store.load(await loadFixture());
  return store;
}

/** Render UI inside a `Presentation.Provider` bound to the given store. */
export function withStore(store: Store, ui: React.ReactNode) {
  return render(<Presentation.Provider store={store}>{ui}</Presentation.Provider>);
}
