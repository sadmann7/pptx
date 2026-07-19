import * as React from "react";

import { act, screen } from "@testing-library/react";
import { beforeAll, describe, expect, it } from "vitest";

import { Presentation } from "../index";
import { createStore } from "../store";
import { loadFixture, withStore } from "./test-utils";

let fixture: ArrayBuffer;

beforeAll(async () => {
  fixture = await loadFixture();
});

describe("Presentation.Loading", () => {
  it("renders only while loading, with progress via render function", async () => {
    const store = createStore();
    const pending = store.load(fixture);

    withStore(store, <Presentation.Loading>{(p) => <span>at {p}%</span>}</Presentation.Loading>);
    expect(screen.getByText(/at \d+%/)).toBeDefined();

    await act(async () => {
      await pending;
    });
    expect(screen.queryByText(/at \d+%/)).toBeNull();
  });

  it("renders nothing when idle", () => {
    const store = createStore();
    const { container } = withStore(store, <Presentation.Loading>busy</Presentation.Loading>);
    expect(container.textContent).toBe("");
  });
});
