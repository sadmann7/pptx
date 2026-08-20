/**
 * Panels sit on a saturated gradient, where a wide black shadow desaturates the
 * colour beneath it and reads as a smudge instead of depth. Tinting it with the
 * backdrop's own darkest value and keeping it tight lets it deepen the gradient
 * rather than cover it.
 */
export function panelShadow(strength: number) {
  return `0 16px 40px rgba(6,8,13,${0.34 * strength})`;
}
