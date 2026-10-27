/**
 * Slide-show tab actions extracted from App.tsx: starting shows,
 * presenter view, custom shows, rehearsal timings, and hiding slides.
 * Functions take the ActionCtx built fresh per call.
 */
import type { ActionCtx } from './action-context'

/** Instant black curtain under the upcoming show (body::after overlay): painted on the
 *  very next frame after the click, it hides the React mount + window-snap latency.
 *  SlideShowView lifts it once the show is revealed (and on unmount, for early exits). */
export function dropShowCurtain(): void {
  document.body.classList.add('show-curtain')
}
export function liftShowCurtain(): void {
  document.body.classList.remove('show-curtain')
}

export function startSlideShow(ctx: ActionCtx, fromStart: boolean): void {
  if (ctx.slides.length === 0 || ctx.slideShow || ctx.presenter) return
  dropShowCurtain()
  // From start: jump to the first unhidden slide (if all hidden, still start at slide 1)
  const first = ctx.slides.findIndex((s) => !s.hidden)
  ctx.setSlideShow({ startAt: fromStart ? Math.max(0, first) : ctx.current })
}

export function exitSlideShow(ctx: ActionCtx, lastIndex: number): void {
  ctx.setSlideShow(null)
  ctx.setCurrent(lastIndex)
}

/** Presenter view (single-window version, entry aligned with the show) */
export function startPresenterView(ctx: ActionCtx, fromStart: boolean): void {
  if (ctx.slides.length === 0 || ctx.slideShow || ctx.presenter) return
  dropShowCurtain()
  const first = ctx.slides.findIndex((s) => !s.hidden)
  ctx.setPresenter({ startAt: fromStart ? Math.max(0, first) : ctx.current })
}

export function exitPresenterView(ctx: ActionCtx, lastIndex: number): void {
  ctx.setPresenter(null)
  ctx.setCurrent(lastIndex)
}

/** "Switch to normal show" inside presenter view: seamlessly turns into a full-screen show in this window */
export function switchPresenterToShow(ctx: ActionCtx, lastIndex: number): void {
  ctx.setPresenter(null)
  ctx.setCurrent(lastIndex)
  ctx.setSlideShow({ startAt: lastIndex })
}
