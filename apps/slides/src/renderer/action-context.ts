/**
 * Shared context for action modules extracted from App.tsx.
 *
 * App builds this bundle fresh on every render (`ctxRef.current = {...}`), and
 * exposes each extracted action as a stable `useCallback(() => impl(ctxRef.current, ...))`
 * wrapper, so module functions always see the latest state without stale closures.
 */
import type React from 'react'
import type { RenderSlide } from '@genoffice/pptx-render'

type Set<T> = React.Dispatch<React.SetStateAction<T>>

// ── State payload types shared between App useState declarations and the ctx ──

export interface SlideShowState {
  startAt: number
  customOrder?: number[]
  rehearse?: boolean
  /** Slide Show > Record: the show records the narration of each slide */
  record?: boolean
}

/** App state bundle passed to every extracted action; refreshed each render. */
export interface ActionCtx {
  // Document
  slides: RenderSlide[]
  setSlides: Set<RenderSlide[]>
  current: number
  setCurrent: Set<number>
  /** slides[current] */
  slide: RenderSlide | undefined
  path: string | null
  setPath: Set<string | null>
  setStatus: Set<string>

  // Show / presenter / rehearse
  slideShow: SlideShowState | null
  setSlideShow: Set<SlideShowState | null>
  presenter: { startAt: number } | null
  setPresenter: Set<{ startAt: number } | null>
}
