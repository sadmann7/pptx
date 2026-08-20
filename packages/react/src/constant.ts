import type * as React from "react";

import type { StoreState } from "./store";

/**
 *  Idle state shared by the store's initial value and the ssr snapshot.
 */
export const DEFAULT_STORE_STATE: StoreState = {
  status: "idle",
  presentation: null,
  activeSlideId: null,
  zoom: 1,
  zoomLevel: "fit",
  progress: 0,
  error: null,
  revision: 0,
};

/**
 * Visually hidden style for screen readers.
 */
export const VISUALLY_HIDDEN_STYLE: React.CSSProperties = {
  position: "absolute",
  width: "1px",
  height: "1px",
  padding: "0",
  margin: "-1px",
  overflow: "hidden",
  clipPath: "inset(50%)",
  whiteSpace: "nowrap",
  borderWidth: 0,
};
