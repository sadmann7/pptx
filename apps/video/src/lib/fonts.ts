import { loadFont as loadGeistSans } from "@remotion/google-fonts/Geist";
import { loadFont as loadGeistMono } from "@remotion/google-fonts/GeistMono";

export const { fontFamily: geistSans } = loadGeistSans("normal", {
  weights: ["400", "500", "600", "700", "800"],
  subsets: ["latin"],
});

export const { fontFamily: geistMono } = loadGeistMono("normal", {
  weights: ["400", "500", "600"],
  subsets: ["latin"],
});
