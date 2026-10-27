/**
 * slides main-process <-> renderer IPC contract (Phase 3: open/save/edit, AI not included yet).
 *
 * Architecture: pptx parsing/saving needs node:crypto/Buffer and can only run in the main
 * process (Node). The main process holds the parsed deck (with originalXml/archive) and sends
 * the renderer only plain-data RenderSlide (built by pptx-render, no Node dependency); the
 * renderer sends edit intents (text/geometry changes) back to the main process, which applies
 * them to the model and rebuilds the RenderSlide.
 */
import type { RenderSlide } from '@genoffice/pptx-render'
// the font catalog's shape is the store's to declare, not a copy kept in step by hand
import type { CatalogEntry as FontCatalogEntry } from '@genoffice/electron-utils/font-store'
import type { LinkTargetOp } from '@genoffice/pptx-ops'

export type { LinkTargetOp }

export type UiTheme = 'light' | 'dark' | 'system'

export interface OpenResult {
  path: string
  slides: RenderSlide[]
  /** Slide page size in EMU */
  size: { cx: number; cy: number }
  /** Theme body default font (fallback shown in the font box when the selection has no text element) */
  defaultFont?: string
}

export type TransitionKind =
  | 'none'
  | 'morph'
  | 'fade'
  | 'push'
  | 'wipe'
  | 'split'
  | 'circle'
  | 'cover'
  | 'pull'
  | 'dissolve'
  | 'zoom'
  | 'random'

// ── Shape animations (the "Animations" tab) ──────────────────────────

/** Reveal direction for the directional effects. 'bottom' reveals upward (default). */
export type AnimDirection = 'top' | 'bottom' | 'left' | 'right'

export type AnimEffectKind =
  | 'appear'
  | 'fade'
  | 'flyIn'
  | 'wipe'
  /** Write-compatible alias for wipe with direction 'top' (the same preset subtype). */
  | 'wipeDown'
  | 'splitIn'
  | 'bounce'
  | 'flipIn'
  | 'zoom' // entrance
  | 'pulse'
  | 'spin'
  | 'grow'
  | 'teeter' // emphasis
  | 'disappear'
  | 'fadeOut'
  | 'flyOut'
  | 'wipeOut'
  | 'shrink'
  | 'zoomOut' // exit
  | 'motionPath' // motion path (move along a path)
  | 'mediaPlay'
  | 'mediaPause'
  | 'mediaStop' // media commands on a video/audio shape

export type AnimTrigger = 'onClick' | 'withPrev' | 'afterPrev'

/** One animation (list order = play order); sourceId locates the target element. */
export interface AnimationItem {
  sourceId: string
  /** Target element display name (shown in the animation pane, filled by the main process) */
  targetName: string
  effect: AnimEffectKind
  trigger: AnimTrigger
  durationMs: number
  delayMs: number
  /** Path when effect='motionPath' (SVG subset M/L/C/Z, coordinates 0..1 relative to slide width/height) */
  motionPath?: string
  /** Wipe/fly direction (enters only when the engine models one); the player reads it
   *  so a top wipe does not play bottom-up. Defaults per effect when absent. */
  direction?: AnimDirection
  /** Per-paragraph animation: 0-based paragraph number; default = the whole shape */
  paragraph?: number
}

/** Element pairing key for Morph transitions: cNvPr id / name are stable across pages; sourceId is per-page temporary. */
export interface ShapeKey {
  sourceId: string
  spid: number | null
  name: string
}

/** What the show needs to play a slide by itself: its saved advance time and its narration. */
export interface ShowTiming {
  advanceMs: number | null
  /** Element ids of the slide's recorded narration pictures, hidden during the show */
  narrationIds: string[]
}

/** Show sync state from presenter view -> audience window (absolute state mirror; audience seek is idempotent) */
export interface ShowSyncState {
  /** Original index of the current page */
  idx: number
  /** Number of in-page animation steps already played */
  played: number
  /** Whether the current step is playing (audience plays that step from the start) */
  playing: boolean
  /** Whether this page change is forward (audience plays the transition effect) */
  fresh: boolean
  /** Reached the "end of show" black screen */
  ended: boolean
  /** Presenter toggled black screen (B key/toolbar button) */
  black: boolean
  /** Presenter toggled white screen (W key); absent from older presenters */
  white?: boolean
}

/** Presenter ink/laser event; coordinates are 0..1 normalized relative to the slide area (laser x<0 = off) */
export type ShowInkEvent =
  | { type: 'laser'; x: number; y: number }
  | { type: 'stroke-start'; x: number; y: number; color: string }
  | { type: 'stroke-move'; x: number; y: number }
  | { type: 'clear' }

/** Navigation actions sent back from the audience window (click/keypress) */
export type AudienceNavAction = 'next' | 'prev' | 'exit'

export interface SlidesApi {
  /** current UI language (persisted by the shell in app-settings.json) */
  getLanguage: () => Promise<
    'zh' | 'en' | 'ja' | 'ko' | 'fr' | 'de' | 'es' | 'th' | 'id' | 'ru' | 'ar'
  >
  /** language switched from the shell home page */
  onLanguageChanged: (
    handler: (
      lang: 'zh' | 'en' | 'ja' | 'ko' | 'fr' | 'de' | 'es' | 'th' | 'id' | 'ru' | 'ar',
    ) => void,
  ) => () => void
  /** current UI theme preference (persisted by the shell in app-settings.json) */
  getTheme: () => Promise<UiTheme>
  /** theme switched from the shell home page */
  onThemeChanged: (handler: (theme: UiTheme) => void) => () => void
  /** snap the host window in/out of instant fullscreen for the slideshow
   *  (macOS simpleFullScreen — skips the animated Space transition) */
  setShowFullScreen: (on: boolean) => Promise<void>
  /** Office-private faces referenced by layouts so far; the renderer registers them as FontFaces
   *  (files invisible to Chromium — DFonts/cloud fonts), so drawing uses the measuring font. */
  privateFontFaces: () => Promise<
    Array<{ id: string; family: string; bold: boolean; italic: boolean }>
  >
  /** Single-face sfnt bytes for one private face (null = gone/unreadable) */
  privateFontData: (id: string) => Promise<ArrayBuffer | null>
  /**
   * Curated downloadable (OFL) font catalog with per-family install state and
   * what each family would cost. Typed from the store that produces it rather
   * than restated here: `ipcMain.handle` is untyped, so a drifted copy of this
   * shape passes typecheck and fails at the point of use.
   */
  fontCatalog: () => Promise<FontCatalogEntry[]>
  /** Download a catalog family into the user font store; layouts refresh via deck-changed */
  fontDownload: (family: string) => Promise<{ ok: boolean; error?: string }>
  /** File picker → install local font files into the user font store */
  fontInstallLocal: () => Promise<{ families: string[] }>
  /** Families this deck references that are missing locally but downloadable */
  fontMissing: () => Promise<string[]>
  /** The user font store changed (download/local install): re-sync private FontFaces */
  onFontsChanged: (handler: () => void) => () => void
  openPptx: (fitWidthPx: number) => Promise<OpenResult | null>
  consumePendingOpen: (fitWidthPx: number) => Promise<OpenResult | null>
  /** New blank presentation (single blank 16:9 page, untitled) */
  newBlank: (fitWidthPx: number) => Promise<OpenResult>
  /** Read-only: RenderSlide for every page of the current session (E2E driver/debug use) */
  getRenderSlides: () => Promise<RenderSlide[] | null>
  /** Read an audio/video element's media data (double-click playback); embedded media converts to dataUrl, external links return as-is */
  getMediaData: (
    slideIndex: number,
    sourceId: string,
  ) => Promise<{ kind: 'video' | 'audio'; dataUrl: string } | null>
  /** All element hyperlinks on a slide (groups included) — slideshow click hit-testing */
  getSlideLinks: (slideIndex: number) => Promise<Array<{ sourceId: string; target: LinkTargetOp }>>
  /** Run-level hyperlinks on a slide (resolved live); keyed by element + paragraph + run indexes */
  getRunLinks: (
    slideIndex: number,
  ) => Promise<
    Array<{ sourceId: string; paraIndex: number; runIndex: number; target: LinkTargetOp }>
  >
  /** The current page's transition effect (echoed on page switch) */
  getTransition: (slideIndex: number) => Promise<TransitionKind>
  getShowTiming: (slideIndex: number) => Promise<ShowTiming>
  /** Asks the OS for the microphone where it gates it (macOS); false when refused. */
  microphoneAccess: () => Promise<boolean>
  /** The current page's animation list (read by the Animations tab / during shows) */
  getAnimations: (slideIndex: number) => Promise<AnimationItem[]>
  /** Morph pairing keys of the current page's elements (matched against same-name/same-id elements on the previous page during morph tweening) */
  getShapeKeys: (slideIndex: number) => Promise<ShapeKey[]>
  /** Plain text of the current page's speaker notes ('' when there are none) */
  getNotes: (slideIndex: number) => Promise<string>
  /** Another window attached to the same file changed the deck (shared session): fresh render state to apply */
  onDeckChanged: (
    handler: (state: { slides: RenderSlide[]; size: { cx: number; cy: number } }) => void,
  ) => () => void
  onOpened: (handler: (result: OpenResult) => void) => () => void
  /** The file was renamed externally (shell Home list rename) — pushes the new path, the renderer updates the title bar */
  onRenamed: (handler: (newPath: string) => void) => () => void
  // ── Presenter-view multi-screen show ────────────────────────────────
  /** Enter presenter view: detects multiple displays and opens a fullscreen audience show window on the external screen (sharing this session's document) */
  presenterStart: () => Promise<{ audience: boolean }>
  /** Broadcast show state to the audience window (fire-and-forget) */
  presenterSync: (state: ShowSyncState) => void
  /** Broadcast ink/laser events to the audience window */
  presenterInk: (ev: ShowInkEvent) => void
  /** Swap the displays of the presenter/audience windows; returns false with no audience window or same screen */
  presenterSwap: () => Promise<boolean>
  /** Exit presenter view: close the audience window */
  presenterEnd: () => Promise<void>
  /** Audience window: fetch the presenter's most recent sync state (re-sent when mounting after the broadcast) */
  audienceReady: () => Promise<ShowSyncState | null>
  /** Audience window: send navigation actions back to the presenter */
  audienceNav: (action: AudienceNavAction) => void
  /** Audience window: subscribe to presenter sync state */
  onShowSync: (handler: (state: ShowSyncState) => void) => () => void
  /** Audience window: subscribe to ink/laser events */
  onShowInk: (handler: (ev: ShowInkEvent) => void) => () => void
  /** Presenter: subscribe to navigation actions sent back by the audience window */
  onAudienceNav: (handler: (action: AudienceNavAction) => void) => () => void
}

declare global {
  interface Window {
    slidesApi: SlidesApi
  }
}
