/**
 * The contract between the harness page and everything that drives it.
 *
 * The harness (main.ts) assigns these; specs and the debugging scripts poll
 * them to know when a render settled and to reach into the loaded deck.
 */
import type { SerializedPresentation } from "@diceui/pptx-core";
import type { EmbeddedFontError } from "@diceui/pptx-core/fonts";

declare global {
  /** Which path the font harness decodes with; see harness/fonts.ts. */
  type BenchMode = "worker" | "main" | "sliced";

  /** A font part the harness sabotaged, and what the loader should report. */
  interface BrokenFontPart {
    path: string;
    kind: "missing" | "empty" | "undecodable" | "unregisterable";
    stage: EmbeddedFontError["stage"];
    /** Families that lose their font with this part. */
    typefaces: string[];
  }

  interface Window {
    /** True once the current slide, including async media and charts, settled. */
    __renderDone?: boolean;
    /** Holds the error instead of setting __renderDone when loading or rendering threw. */
    __renderError?: string;
    __slideCount?: number;
    __slideWidth?: number;
    __slideHeight?: number;
    /** Re-renders another slide of the already loaded deck. */
    __showSlide?: (index: number) => Promise<void>;
    __getStructure?: () => SerializedPresentation;
    /** True once the thumbnail harness has a loaded deck mounted. */
    __thumbnailsReady?: boolean;
    /**
     * Per-slide cost of each render mode, one entry per slide, in ms
     * (thumbnail harness). Rendering and mounting are separate because the
     * list pays both and they cost about the same on a detailed slide.
     */
    __benchRenderModes?: () => {
      slide: { render: number[]; mount: number[] };
      thumbnail: { render: number[]; mount: number[] };
    };
    /** True once the embedded-font harness has a deck ready to decode. */
    __fontsReady?: boolean;
    /** Unique font parts the loaded deck will decode, and their total size. */
    __fontParts?: number;
    __fontBytes?: number;
    /** Decodes every font part on the main thread, timing each one (ms). */
    __benchDecodeParts?: () => { path: string; bytes: number; ms: number }[];
    /**
     * Breaks one font part per failure kind, loads the rest, and reports what
     * the loader passed to `onError`. Leaves the deck broken, so a page that
     * calls this must not also be measured.
     */
    __benchFontErrors?: () => Promise<{
      broken: BrokenFontPart[];
      errors: EmbeddedFontError[];
      /** Faces added to `document.fonts`, i.e. the parts left intact. */
      faces: number;
    }>;
    /**
     * Runs the font pipeline once in the mode the page was loaded with, and
     * reports both the wall clock and what the main thread lost to it.
     */
    __benchFonts?: () => Promise<{
      mode: BenchMode;
      parts: number;
      bytes: number;
      readyMs: number;
      completeMs: number;
      /** Faces added to `document.fonts`, i.e. parts the browser accepted. */
      faces: number;
      frames: number;
      worstFrameMs: number;
      blockedMs: number;
      longTasks: number;
      worstTaskMs: number;
      totalTaskMs: number;
    }>;
  }
}
