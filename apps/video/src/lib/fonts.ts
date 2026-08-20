import type { CSSProperties } from "react";

import { loadFont as loadGeistSans } from "@remotion/google-fonts/Geist";
import { loadFont as loadGeistMono } from "@remotion/google-fonts/GeistMono";

const sans = loadGeistSans("normal", {
  weights: ["400", "500", "600", "700", "800"],
  subsets: ["latin"],
});

const mono = loadGeistMono("normal", {
  weights: ["400", "500", "600"],
  subsets: ["latin"],
});

export const geistSans = sans.fontFamily;
export const geistMono = mono.fontFamily;

export const fontVars = {
  "--font-geist-sans": geistSans,
  "--font-geist-mono": geistMono,
} as CSSProperties;
