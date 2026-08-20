import { Easing, interpolate } from "remotion";

export const FADE_DURATION = 12;
export const CONTENT_FADE = 10;

export const clamp = {
  extrapolateLeft: "clamp",
  extrapolateRight: "clamp",
} as const;

export const ease = Easing.bezier(0.22, 1, 0.36, 1);
export const cameraEase = Easing.bezier(0.7, 0, 0.25, 1);

export function progress(frame: number, from: number, to: number) {
  return interpolate(frame, [from, to], [0, 1], { ...clamp, easing: ease });
}
