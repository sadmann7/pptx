/**
 * Globals the harness pages assign and the specs and scripts poll.
 */
import type { PresentationStore } from "@diceui/pptx";
import type { SerializedPresentation } from "@diceui/pptx-core";
import type { EmbeddedFontError } from "@diceui/pptx-core/fonts";

declare global {
  type BenchMode = "worker" | "main" | "sliced";

  type ViewerStore = "root" | "hook";

  interface BrokenFontPart {
    path: string;
    kind: "missing" | "empty" | "undecodable" | "unregisterable";
    stage: EmbeddedFontError["stage"];
    typefaces: string[];
  }

  interface Window {
    // Waits for async media and charts too.
    __renderDone?: boolean;
    // Set instead of __renderDone when loading or rendering threw.
    __renderError?: string;
    __slideCount?: number;
    __slideWidth?: number;
    __slideHeight?: number;
    __showSlide?: (index: number) => Promise<void>;
    __getStructure?: () => SerializedPresentation;
    __editingReady?: boolean;
    __editStore?: PresentationStore;
    __thumbnailsReady?: boolean;
    // Rendering and mounting are timed apart because the list pays both and they
    // cost about the same on a detailed slide.
    __benchRenderModes?: () => {
      slide: { render: number[]; mount: number[] };
      thumbnail: { render: number[]; mount: number[] };
    };
    __mountViewer?: (file: string, store?: ViewerStore) => Promise<void>;
    __unmountViewer?: () => void;
    __viewerLoaded?: boolean;
    __fontsReady?: boolean;
    __fontParts?: number;
    __fontBytes?: number;
    __benchDecodeParts?: () => { path: string; bytes: number; ms: number }[];
    // Leaves the deck broken, so a page that calls this must not also be measured.
    __benchFontErrors?: () => Promise<{
      broken: BrokenFontPart[];
      errors: EmbeddedFontError[];
      faces: number;
    }>;
    __benchFonts?: () => Promise<{
      mode: BenchMode;
      parts: number;
      bytes: number;
      readyMs: number;
      completeMs: number;
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
