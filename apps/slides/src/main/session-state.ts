/**
 * Shared main-process state for GenOffice Slides, extracted from slides-main.ts so
 * the IPC modules (slides-main, ai-ipc, presenter-show) can share it:
 * per-renderer sessions, snapshot undo/redo history, runtime paths, window
 * references, and RenderSlide rebuild helpers.
 */
import { BrowserWindow } from 'electron'
import type { WebContents } from 'electron'
import { join } from 'node:path'
import {
  materializeSlide,
  parseClrMap,
  parseTheme,
  resolveSchemeColor,
  type OpenedPptx,
} from '@genoffice/pptx-engine'
import {
  buildRenderSlide,
  type FontMetricsProvider,
  type RenderSlide,
} from '@genoffice/pptx-render'
import { createSystemFontMetrics, resetFontRegistry } from './fonts'
import { tiffToPng } from './tiff-decode'
import { neutralizeJpegOrientation } from './jpeg-orientation'
import { displayMime } from './media-mime'

export interface RuntimePaths {
  preloadPath: string
  rendererDevUrl?: string | undefined
  rendererFilePath?: string | undefined
  /** Shell router used to open exported PDFs in a new GenOffice tab. */
  openGeneratedPath?: (path: string) => boolean
}

export const runtime: RuntimePaths = {
  preloadPath: join(__dirname, '../preload/index.js'),
  rendererDevUrl: process.env.ELECTRON_RENDERER_URL,
  rendererFilePath: join(__dirname, '../renderer/index.html'),
}

// One session per renderer process (standalone window or shell tab), keyed by webContents.id
export interface Session {
  path: string
  opened: OpenedPptx
  fitWidthPx: number
}
export const sessions = new Map<number, Session>()

/** webContents ids currently mapped to this session (aliased entries included). */
export function attachedIds(session: Session): number[] {
  const ids: number[] = []
  for (const [id, s] of sessions) if (s === session) ids.push(id)
  return ids
}

// ── Window references (shell tab mode + active renderer tracking) ──────
export const windowRefs = {
  /** Parent window for dialogs in tab mode (the shell's single BrowserWindow) */
  shellWindow: null as BrowserWindow | null,
  /** Currently active slides renderer (window or tab view) — target of menu commands; the shell updates it on tab switch */
  activeWebContents: null as WebContents | null,
}

/** Shell-registered hook (aggregate/tab mode only): cover the tab strip with a tab's
 *  view during a slideshow without going through HTML fullscreen. Standalone slides
 *  windows have no tab strip and leave this null. */
export const showChrome = {
  setBleed: null as ((wc: WebContents, on: boolean) => void) | null,
  /** the window hosting a tab's WebContentsView when BrowserWindow.fromWebContents
   *  cannot tell (a detached "Open in New Window" / torn-off editor) */
  hostWindow: null as ((wc: WebContents) => BrowserWindow | undefined) | null,
}

/** the window a tab view lives in: its own BrowserWindow, the shell-registered
 *  detached window, else the shell window */
export function hostWindowFor(wc: WebContents): BrowserWindow | undefined {
  const own = BrowserWindow.fromWebContents(wc) ?? showChrome.hostWindow?.(wc)
  if (own && !own.isDestroyed()) return own
  return windowRefs.shellWindow && !windowRefs.shellWindow.isDestroyed()
    ? windowRefs.shellWindow
    : undefined
}

export function dialogParent(): BrowserWindow | undefined {
  // Focused window first: a detached editor window must parent its own dialogs
  // (the shell's tab views live inside the shell window, so tab mode is unchanged)
  return BrowserWindow.getFocusedWindow() ?? windowRefs.shellWindow ?? undefined
}

// ── RenderSlide rebuild helpers ─────────────────────────────────────────

/** Precise system-font metrics (lazily built, shared process-wide; unmatched fonts fall back to heuristics per run). */
let fontMetrics: FontMetricsProvider | null = null
export function getFontMetrics(): FontMetricsProvider {
  if (!fontMetrics) fontMetrics = createSystemFontMetrics()
  return fontMetrics
}

/** Drop the cached metrics (and its font registry) after the user font store changes. */
export function resetFontMetrics(): void {
  resetFontRegistry()
  fontMetrics = null
}

export function buildAllRenderSlides(opened: OpenedPptx, fitWidthPx: number): RenderSlide[] {
  return opened.deck.slides.map((s, i) =>
    buildRenderSlide(s, opened.deck.size, {
      fitWidthPx,
      media: makeMediaResolver(opened, s.path),
      metrics: getFontMetrics(),
      slideNo: i + 1,
    }),
  )
}

/** Office theme-class slot (MsftOfcThm_<slot>_Fill/_Stroke) → schemeClr name. */
const SVG_THEME_SLOTS: Record<string, string> = {
  background1: 'bg1',
  text1: 'tx1',
  background2: 'bg2',
  text2: 'tx2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hyperlink: 'hlink',
  followedhyperlink: 'folHlink',
}

/**
 * Office-exported SVGs carry `.MsftOfcThm_<slot>_Fill/_Stroke` CSS classes whose baked
 * values PowerPoint rewrites to the CURRENT theme's colors at render time (the static
 * value is just the export-time snapshot — probe deck: a bg1-classed blob draws white
 * on a white theme, not its baked blue). Mirror that rewrite before serving the SVG.
 */
export function retintThemedSvg(svg: string, opened: OpenedPptx, slidePath?: string): string {
  const path = slidePath ?? opened.deck.slides[0]?.path
  if (!path) return svg
  let theme
  try {
    const chain = opened.archive.resolveSlideChain(path)
    const themeXml = chain.themePath ? opened.archive.readText(chain.themePath) : null
    if (!themeXml) return svg
    theme = parseTheme(themeXml)
    const masterXml = chain.masterPath ? opened.archive.readText(chain.masterPath) : undefined
    const layoutXml = chain.layoutPath ? opened.archive.readText(chain.layoutPath) : undefined
    theme.clrMap = parseClrMap(
      masterXml ?? undefined,
      layoutXml ?? undefined,
      opened.archive.readText(path) ?? undefined,
    )
  } catch {
    return svg
  }
  return svg.replace(
    /\.MsftOfcThm_(\w+?)_(Fill|Stroke)\w*\s*\{[^}]*\}/g,
    (rule, slot: string, kind: string) => {
      const scheme = SVG_THEME_SLOTS[slot.toLowerCase()]
      const color = scheme ? resolveSchemeColor(scheme, theme) : undefined
      if (!color) return rule
      const prop = kind === 'Stroke' ? 'stroke' : 'fill'
      return rule.replace(new RegExp(`${prop}\\s*:\\s*[^;}]+`, 'g'), `${prop}:${color}`)
    },
  )
}

/** Image mediaRef -> dataUrl (lazily decoded). TIFF is transcoded to PNG for display
    (Chromium can't decode it); the archive keeps the original bytes for save fidelity.
    The mime comes from magic-byte sniffing first (legacy decks mislabel media — a PNG
    stored as .emf must not enter the EMF parser), extension second. */
export function makeMediaResolver(opened: OpenedPptx, slidePath?: string) {
  const cache = new Map<string, string | undefined>()
  return (mediaRef: string): string | undefined => {
    if (cache.has(mediaRef)) return cache.get(mediaRef)
    const bytes = opened.archive.readBytes(mediaRef)
    let url: string | undefined
    if (bytes) {
      const mime = displayMime(mediaRef, bytes)
      if (mime === 'image/tiff') {
        const decoded = tiffToPng(bytes)
        if (decoded) url = `data:image/png;base64,${Buffer.from(decoded.png).toString('base64')}`
      } else if (mime === 'image/svg+xml') {
        let text = Buffer.from(bytes).toString('utf8')
        if (text.includes('MsftOfcThm_')) text = retintThemedSvg(text, opened, slidePath)
        url = `data:${mime};base64,${Buffer.from(text, 'utf8').toString('base64')}`
      } else {
        // PowerPoint ignores EXIF orientation; Chromium applies it on decode — neutralize
        // the flag so rotated-pixel JPEGs with a shape-level rot don't double-rotate
        const served = mime === 'image/jpeg' ? neutralizeJpegOrientation(bytes) : bytes
        url = `data:${mime};base64,${Buffer.from(served).toString('base64')}`
      }
    }
    cache.set(mediaRef, url)
    return url
  }
}

/** Rebuild a single page's RenderSlide (sent back to the renderer after an edit). */
export function rebuildSlide(session: Session, slideIndex: number): RenderSlide | null {
  const slide = session.opened.deck.slides[slideIndex]
  if (!slide) return null
  return buildRenderSlide(slide, session.opened.deck.size, {
    fitWidthPx: session.fitWidthPx,
    media: makeMediaResolver(session.opened, slide.path),
    metrics: getFontMetrics(),
    slideNo: slideIndex + 1,
  })
}

/**
 * Rebuild the RenderSlide after re-parsing the whole page (the model is stale after a chart
 * edit and needs a reparse). Equivalent to materializeSlide then rebuildSlide, but without the
 * structureDirty save logic.
 */
export function rebuildSlideWithReparse(session: Session, slideIndex: number): RenderSlide | null {
  const fresh = materializeSlide(session.opened, slideIndex)
  if (!fresh) return null
  return buildRenderSlide(fresh, session.opened.deck.size, {
    fitWidthPx: session.fitWidthPx,
    media: makeMediaResolver(session.opened, fresh.path),
    metrics: getFontMetrics(),
    slideNo: slideIndex + 1,
  })
}
