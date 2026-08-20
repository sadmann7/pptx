import type { CSSProperties } from "react";

export const color = {
  ink: "#0e1117",
  white: "#f5f6f8",
} as const;

export const PANEL_BORDER_WIDTH = 1;
export const PANEL_RADIUS = 18;
export const PANEL_FILL = "rgba(255,255,255,.04)";
export const PANEL_STROKE = "rgba(255,255,255,.12)";

export function panelShadow(strength: number) {
  return `0 16px 40px rgba(6,8,13,${0.34 * strength})`;
}

export function panelBox(reveal: number): CSSProperties {
  return {
    borderRadius: PANEL_RADIUS,
    background: PANEL_FILL,
    border: `${PANEL_BORDER_WIDTH}px solid ${PANEL_STROKE}`,
    boxShadow: panelShadow(reveal),
  };
}
