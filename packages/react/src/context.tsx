import React from "react";
import type { PresentationData, SlideData } from "@pptx/parser";
import type { PresentationState } from "./store";
import { PresentationStore } from "./store";

export const PresentationContext = React.createContext<PresentationStore | null>(null);

function usePresentationStore(): PresentationStore {
  const store = React.useContext(PresentationContext);
  if (!store) throw new Error("[pptx/react] Hooks must be used inside <Presentation.Root>");
  return store;
}

const SERVER_SNAPSHOT: PresentationState = {
  status: "idle", presentation: null, currentIndex: 0, zoom: 1, progress: 0, error: null,
};

export interface UsePresentationResult {
  presentation: PresentationData | null;
  status: PresentationState["status"];
  error: Error | null;
  progress: number;
}

export function usePresentation(): UsePresentationResult {
  const store = usePresentationStore();
  const state = React.useSyncExternalStore(
    store.subscribe.bind(store),
    store.getState.bind(store),
    () => SERVER_SNAPSHOT,
  );
  return { presentation: state.presentation, status: state.status, error: state.error, progress: state.progress };
}

export interface UseSlideResult {
  slide: SlideData | null;
  index: number;
  total: number;
  isFirst: boolean;
  isLast: boolean;
  goTo: (index: number) => void;
  next: () => void;
  prev: () => void;
}

export function useSlide(): UseSlideResult {
  const store = usePresentationStore();
  const currentIndex = React.useSyncExternalStore(store.subscribe.bind(store), () => store.getState().currentIndex, () => 0);
  const total = React.useSyncExternalStore(store.subscribe.bind(store), () => store.getState().presentation?.slides.length ?? 0, () => 0);
  const slide = React.useSyncExternalStore(store.subscribe.bind(store), () => {
    const { presentation, currentIndex: idx } = store.getState();
    return presentation?.slides[idx] ?? null;
  }, () => null);
  const goTo = React.useCallback((i: number) => store.goTo(i), [store]);
  const next = React.useCallback(() => store.next(), [store]);
  const prev = React.useCallback(() => store.prev(), [store]);
  return { slide, index: currentIndex, total, isFirst: currentIndex === 0, isLast: total > 0 && currentIndex === total - 1, goTo, next, prev };
}

export interface UseZoomResult {
  zoom: number;
  setZoom: (zoom: number) => void;
  zoomIn: (step?: number) => void;
  zoomOut: (step?: number) => void;
  fitTo: (containerWidth: number, containerHeight: number, padding?: number) => void;
}

export function useZoom(): UseZoomResult {
  const store = usePresentationStore();
  const zoom = React.useSyncExternalStore(store.subscribe.bind(store), () => store.getState().zoom, () => 1);
  return {
    zoom,
    setZoom: React.useCallback((z: number) => store.setZoom(z), [store]),
    zoomIn: React.useCallback((step?: number) => store.zoomIn(step), [store]),
    zoomOut: React.useCallback((step?: number) => store.zoomOut(step), [store]),
    fitTo: React.useCallback((w: number, h: number, p?: number) => store.fitTo(w, h, p), [store]),
  };
}

export function usePresentationStoreRef(): PresentationStore { return usePresentationStore(); }
