import type { PresentationState } from "./store";

/**
 *  Idle state shared by the store's initial value and the ssr snapshot.
 */
export const DEFAULT_PRESENTATION_STATE: PresentationState = {
  status: "idle",
  presentation: null,
  activeSlideId: null,
  zoom: 1,
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

/**
 * Resets inherited typography so host-page styles (e.g. `.prose`, Tailwind
 * base) don't bleed into rendered slide content.
 */
export const TYPOGRAPHY_RESET_STYLE: React.CSSProperties = {
  fontSize: "initial",
  fontFamily: "initial",
  fontWeight: "normal",
  lineHeight: "normal",
  color: "initial",
  letterSpacing: "normal",
  textDecoration: "none",
  textTransform: "none",
};
