export const PANEL_BORDER_WIDTH = 1;

export function panelShadow(strength: number) {
  return `0 16px 40px rgba(6,8,13,${0.34 * strength})`;
}
