import React from 'react'
import { usePresentationStoreRef } from '../context'

export interface ViewportProps {
  children: React.ReactNode
  className?: string
  style?: React.CSSProperties
  /**
   * If true, the viewport computes and applies `fitTo` whenever the container
   * is resized. Requires the ResizeObserver API (available in all modern browsers).
   * @default false
   */
  autoFit?: boolean
  /** Padding (in px) subtracted from each edge when auto-fitting. @default 24 */
  autoFitPadding?: number
}

/**
 * <Presentation.Viewport>
 *
 * A scrollable container that holds <Presentation.Slide> (and any overlays).
 * When `autoFit` is true, it observes its own size and keeps the slide fitted.
 */
export function Viewport({
  children,
  className,
  style,
  autoFit = false,
  autoFitPadding = 24,
}: ViewportProps) {
  const ref = React.useRef<HTMLDivElement>(null)
  const store = usePresentationStoreRef()

  // Auto-fit: observe container size and call store.fitTo whenever it changes
  React.useEffect(() => {
    if (!autoFit || !ref.current) return

    const el = ref.current

    const fit = () => {
      store.fitTo(el.clientWidth, el.clientHeight, autoFitPadding)
    }

    fit() // initial fit

    if (typeof ResizeObserver === 'undefined') return

    const observer = new ResizeObserver(fit)
    observer.observe(el)
    return () => observer.disconnect()
  }, [autoFit, autoFitPadding, store])

  const containerStyle: React.CSSProperties = {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    overflow: 'auto',
    position: 'relative',
    ...style,
  }

  return (
    <div ref={ref} className={className} style={containerStyle}>
      {children}
    </div>
  )
}
