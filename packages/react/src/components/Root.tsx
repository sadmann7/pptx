import React from 'react'
import type { ParseOptions, PresentationInput } from '@pptx/parser'
import { PresentationContext } from '../context'
import { PresentationStore } from '../store'

export interface RootProps {
  /**
   * The PPTX file to load. Accepts File, ArrayBuffer, Uint8Array, or Blob.
   * Changing this value triggers a re-parse.
   */
  file: PresentationInput | null | undefined
  children: React.ReactNode
  /** Called when the presentation finishes loading */
  onLoad?: (store: PresentationStore) => void
  /** Called when parsing fails */
  onError?: (error: Error) => void
  /** Called as each slide is parsed. Useful for progress UIs. */
  onProgress?: (current: number, total: number) => void
  /** Additional parse options forwarded to @pptx/parser */
  parseOptions?: Omit<ParseOptions, 'onProgress'>
}

/**
 * <Presentation.Root> — the context provider.
 *
 * Creates and owns one PresentationStore instance for its subtree.
 * When `file` changes, re-parses automatically.
 */
export function Root({
  file,
  children,
  onLoad,
  onError,
  onProgress,
  parseOptions,
}: RootProps) {
  // One store per Root instance, never recreated
  const store = React.useMemo(() => new PresentationStore(), [])

  // Expose callbacks via refs so they never cause re-triggering of the effect
  const onLoadRef = React.useRef(onLoad)
  const onErrorRef = React.useRef(onError)
  const onProgressRef = React.useRef(onProgress)
  onLoadRef.current = onLoad
  onErrorRef.current = onError
  onProgressRef.current = onProgress

  React.useEffect(() => {
    if (!file) {
      store.reset()
      return
    }

    store
      .load(file, {
        ...parseOptions,
        onProgress: onProgressRef.current,
      } as Parameters<typeof store.load>[1])
      .then(() => {
        onLoadRef.current?.(store)
      })
      .catch((err: unknown) => {
        const error = err instanceof Error ? err : new Error(String(err))
        onErrorRef.current?.(error)
      })
  // parseOptions is intentionally excluded — consumers should memoize it
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [store, file])

  return (
    <PresentationContext.Provider value={store}>
      {children}
    </PresentationContext.Provider>
  )
}
