/**
 * GenOffice Slides main process — pptx parsing/render-tree building/edit application/saving all live
 * here (Node side). The renderer only gets plain-data RenderSlide; edit intents are sent back
 * here to apply. Structure mirrors apps/docs: exports embeddable configure/register/start for
 * future shell reuse.
 */
import {
  app,
  BrowserWindow,
  dialog,
  ipcMain,
  systemPreferences,
  webContents,
} from 'electron'
import type { WebContents } from 'electron'
import { readFile, open } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { basename, dirname, join, resolve } from 'node:path'
import {
  showOpenDialogWithMemory,
  installRendererProtocol,
  registerRendererScheme,
  rendererUrl,
} from '@genoffice/electron-utils'
import { convertOdpToPptx } from '@genoffice/odf2ooxml'
import { getUiLang, normalizeLang, setUiLang } from '@genoffice/i18n'
import {
  getElementLink,
  getSlideLinks,
  getRunLinks,
  createBlankPptx,
  getSlideNotes,
  getSlideTransition,
  elementSpid,
  narrationSpids,
  getSlideAdvanceTime,
  getSlideAnimations,
  listEmbeddedFonts,
  openPptx,
  parseTheme,
  type OpenedPptx,
} from '@genoffice/pptx-engine'
// From its own electron-utils subpath rather than the package barrel: the
// barrel drags Electron-only modules into headless bundles.
import { atomicWriteFile } from '@genoffice/electron-utils/atomic-write'
import { refineComplexWidths, shapedMetricsReady } from './shaped-metrics'
import { cfbKind, isCfbHeader } from './cfb-sniff'
import { OPENABLE_DECK_EXTENSIONS, deckExtension } from '../shared/deck-open-formats'
import type {
  OpenResult,
  ShowTiming,
  AnimationItem,
  ShapeKey,
} from '../shared/ipc'

import { tm } from './i18n-main'
import {
  buildAllRenderSlides,
  dialogParent,
  getFontMetrics,
  hostWindowFor,
  resetFontMetrics,
  rebuildSlide,
  runtime,
  sessions,
  showChrome,
  attachedIds,
  Session,
} from './session-state'
import { listPrivateFontFaces, getPrivateFontData, registerEmbeddedFonts } from './fonts'
import { listMetafileFonts } from './metafile-fonts'
import {
  downloadFontFamily,
  initFontStore,
  installLocalFontFiles,
  listFontCatalog,
  missingCatalogFonts,
} from './font-store'

import { registerPresenterIpc } from './presenter-show'

/** standalone: path queued before window creation (argv/open-file) */
let pendingOpenPath: string | null = null

/**
 * An .odp opens only as a converted .pptx deck: confirm the lossy import,
 * pick the destination (the sibling .pptx by default), convert, write the
 * file and open it — the explicit-destination flow the PDF "Export as
 * PowerPoint" path already uses, so every later Save lands in the copy the
 * user chose. Anything that is not an .odp passes through unchanged; null
 * cancels the open.
 */
async function preparePresentationForOpen(filePath: string): Promise<string | null> {
  const extension = deckExtension(filePath)
  if (extension !== 'odp' && extension !== 'otp') return filePath
  const convertedFilePath = join(dirname(filePath), basename(filePath).replace(/\.(odp|otp)$/i, '.pptx'))
  try {
    const result = await convertOdpToPptx(new Uint8Array(await readFile(filePath)))
    // atomic like the Docs and shell copies: a plain write leaves a truncated
    // .pptx where the user expects a deck if the process dies mid-write
    await atomicWriteFile(convertedFilePath, Buffer.from(result.pptx))
  } catch (e) {
    console.error(e)
    return null
  }
  return convertedFilePath
}

/**
 * .ppt (97-2003 binary compound document) and encrypted OOXML are unsupported: show an actionable message instead of a parse error.
 * Detection uses the magic number rather than the extension -- a binary ppt with a renamed suffix is caught too. A CFB containing an
 * EncryptedPackage stream is a password-protected pptx and gets dedicated copy (instead of being mislabeled as the legacy format).
 */
async function rejectLegacyPpt(path: string): Promise<boolean> {
  let head: Buffer
  try {
    const fh = await open(path, 'r')
    try {
      head = Buffer.alloc(8)
      await fh.read(head, 0, 8, 0)
    } finally {
      await fh.close()
    }
  } catch {
    return false
  }
  if (!isCfbHeader(head)) return false
  let kind: 'legacy' | 'encrypted' = 'legacy'
  try {
    kind = cfbKind(await readFile(path)) ?? 'legacy'
  } catch {
    // on read failure, fall back to the legacy-format message
  }
  console.warn(kind)
  return true
}

/** Register the deck's usable embedded fonts before layout; new faces invalidate cached metrics. */
function adoptEmbeddedFonts(opened: OpenedPptx): void {
  try {
    if (registerEmbeddedFonts(listEmbeddedFonts(opened.archive))) resetFontMetrics()
  } catch {
    // Embedded fonts are best-effort: a malformed fntdata must never block opening
  }
  // Metafile pictures draw text through canvas fonts: resolving their facenames here puts the
  // Office-private faces (Yu Gothic UI, MS PGothic…) on the renderer's private-font list before
  // the EMF/WMF previews rasterize.
  try {
    const metrics = getFontMetrics()
    for (const f of listMetafileFonts(opened.archive))
      metrics.displayFamily?.({
        fontFamily: f.family,
        fontSizePx: 100,
        bold: f.bold,
        italic: f.italic,
      })
  } catch {
    // best-effort as well
  }
}

async function openAndBuild(
  wc: WebContents,
  path: string,
  fitWidthPx: number,
): Promise<OpenResult> {
  // Same file already open in another window: attach to that session instead of
  // opening an independent copy (which would silently lose the loser's edits on
  // save). Both windows then see the same live deck; edits propagate via
  // scheduleDeckBroadcast. Untitled sessions have path '' and never alias.
  const wanted = resolve(path)
  for (const [id, existing] of sessions) {
    if (id === wc.id || !existing.path || resolve(existing.path) !== wanted) continue
    sessions.set(wc.id, existing)
    return {
      path: existing.path,
      slides: buildAllRenderSlides(existing.opened, fitWidthPx),
      size: { cx: existing.opened.deck.size.cx, cy: existing.opened.deck.size.cy },
      defaultFont: deckDefaultFont(existing.opened),
    }
  }
  const raw = await readFile(path)
  // a 0-byte .pptx is an empty deck, not a corrupt one: open the blank template
  // under the file's own path so Save writes back to it
  const bytes = raw.length === 0 ? await createBlankPptx() : new Uint8Array(raw)
  await shapedMetricsReady() // Lay out only after complex-script shaped metrics are ready, avoiding an init race falling back to estimation
  const opened = await openPptx(bytes)
  adoptEmbeddedFonts(opened)
  sessions.set(wc.id, {
    path,
    opened,
    fitWidthPx,
    undoStack: [],
    redoStack: [],
  })
  let slides = buildAllRenderSlides(opened, fitWidthPx)
  // If the first layout pass had complex-script misses (Arabic/Thai etc.), re-lay out once with renderer-measured widths
  if (await refineComplexWidths(wc)) slides = buildAllRenderSlides(opened, fitWidthPx)
  return {
    path,
    slides,
    size: { cx: opened.deck.size.cx, cy: opened.deck.size.cy },
    defaultFont: deckDefaultFont(opened),
  }
}

/** Theme body (minor) Latin font: fallback shown in the ribbon font box when the selection has no text element. */
function deckDefaultFont(opened: OpenedPptx): string | undefined {
  try {
    const slidePath = opened.archive.readPresentation().slidePaths[0]
    if (!slidePath) return undefined
    const themePath = opened.archive.resolveSlideChain(slidePath).themePath
    const xml = themePath ? opened.archive.readText(themePath) : undefined
    return xml ? parseTheme(xml).minorFont : undefined
  } catch {
    return undefined
  }
}

let ipcRegistered = false

export function registerSlidesIpc(): void {
  if (ipcRegistered) return
  ipcRegistered = true

  // shared with the other editor modules — last (identical) registration wins
  ipcMain.removeHandler('app:get-language')
  ipcMain.handle('app:get-language', () => getUiLang())

  ipcMain.handle('slides:private-font-faces', () => listPrivateFontFaces())
  ipcMain.handle('slides:private-font-data', (_e, id: string) => getPrivateFontData(id))

  initFontStore()
  // Fonts changed (download or local install): rebuild every open session with a fresh
  // registry and push the relaid-out slides + a re-sync ping to all attached windows.
  const afterFontsChanged = (): void => {
    resetFontMetrics()
    const seen = new Set<Session>()
    for (const [wcId, session] of sessions) {
      if (seen.has(session)) continue
      seen.add(session)
      const payload = {
        slides: buildAllRenderSlides(session.opened, session.fitWidthPx),
        size: { cx: session.opened.deck.size.cx, cy: session.opened.deck.size.cy },
      }
      for (const id of attachedIds(session))
        webContents.fromId(id)?.send('slides:deck-changed', payload)
      void wcId
    }
    for (const wc of webContents.getAllWebContents()) wc.send('slides:fonts-changed')
  }
  ipcMain.handle('slides:font-catalog', () => listFontCatalog())
  ipcMain.handle('slides:font-download', async (_e, family: string) => {
    try {
      await downloadFontFamily(family)
      afterFontsChanged()
      return { ok: true }
    } catch (err) {
      return { ok: false, error: err instanceof Error ? err.message : String(err) }
    }
  })
  ipcMain.handle('slides:font-install-local', async () => {
    const r = await showOpenDialogWithMemory(dialog, dialogParent(), {
      properties: ['openFile', 'multiSelections'],
      filters: [{ name: 'Fonts', extensions: ['ttf', 'otf', 'ttc', 'otc'] }],
    })
    if (r.canceled || !r.filePaths.length) return { families: [] }
    const families = installLocalFontFiles(r.filePaths)
    if (families.length) afterFontsChanged()
    return { families }
  })
  ipcMain.handle('slides:font-missing', (e) => {
    const session = sessions.get(e.sender.id)
    return session ? missingCatalogFonts(session.opened) : []
  })

  ipcMain.handle('slides:open', async (e, fitWidthPx: number) => {
    const parent = dialogParent()
    const options = {
      properties: ['openFile' as const],
      filters: [{ name: 'PowerPoint', extensions: [...OPENABLE_DECK_EXTENSIONS] }],
    }
    const r = await showOpenDialogWithMemory(dialog, parent, options)
    if (r.canceled || !r.filePaths[0]) return null
    const prepared = await preparePresentationForOpen(r.filePaths[0])
    if (!prepared) return null
    if (await rejectLegacyPpt(prepared)) return null
    return openAndBuild(e.sender, prepared, fitWidthPx)
  })

  ipcMain.handle('slides:open-path', async (e, path: string, fitWidthPx: number) => {
    if (!path || !existsSync(path)) return null
    const prepared = await preparePresentationForOpen(path)
    if (!prepared) return null
    if (await rejectLegacyPpt(prepared)) return null
    return openAndBuild(e.sender, prepared, fitWidthPx)
  })

  ipcMain.handle('slides:consume-pending-open', async (e, fitWidthPx: number) => {
    if (pendingOpenPath && existsSync(pendingOpenPath)) {
      const dropQueued = () => {
        pendingOpenPath = null
      }
      // A CFB file (legacy .ppt / encrypted, possibly misnamed .pptx) can never
      // parse: tell the user and drop it, or every relaunch restores a blank tab
      const prepared = await preparePresentationForOpen(pendingOpenPath)
      if (!prepared || (await rejectLegacyPpt(prepared))) {
        dropQueued()
        return null
      }
      // Clear the queue only after a successful open: keep it on parse failure or a mid-flight renderer reload, so a remount can retry
      const result = await openAndBuild(e.sender, prepared, fitWidthPx)
      dropQueued()
      return result
    }
    // No queued path but the main process already has a session (remount after an HMR full
    // reload/crash recovery) -> restore from the session; otherwise the document is lost leaving
    // only the start screen, and reopening the same file just activates this empty tab with no
    // way to self-heal
    const session = sessions.get(e.sender.id)
    if (session) {
      session.fitWidthPx = fitWidthPx
      return {
        path: session.path,
        slides: buildAllRenderSlides(session.opened, fitWidthPx),
        size: { cx: session.opened.deck.size.cx, cy: session.opened.deck.size.cy },
        defaultFont: deckDefaultFont(session.opened),
      } satisfies OpenResult
    }
    return null
  })

  // Read-only: RenderSlide for every page of the current session (E2E driver/debug use, no state change)
  ipcMain.handle('slides:get-render-slides', (e) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    return session.opened.deck.slides.map((_, i) => rebuildSlide(session, i))
  })

  ipcMain.handle('slides:new-blank', async (e, fitWidthPx: number): Promise<OpenResult> => {
    const opened = await openPptx(await createBlankPptx())
    sessions.set(e.sender.id, { path: '', opened, fitWidthPx, undoStack: [], redoStack: [] })
    return {
      path: '',
      slides: buildAllRenderSlides(opened, fitWidthPx),
      size: { cx: opened.deck.size.cx, cy: opened.deck.size.cy },
      defaultFont: deckDefaultFont(opened),
    }
  })

  // Double-click playback: read the media bytes of an audio/video element (embedded converts to dataUrl, external links return as-is)
  const AV_MIME: Record<string, string> = {
    mp4: 'video/mp4',
    m4v: 'video/mp4',
    // Chromium refuses to even load video/quicktime, but demuxes QuickTime bytes
    // fine through the ISO-BMFF path when served as video/mp4
    mov: 'video/mp4',
    webm: 'video/webm',
    avi: 'video/x-msvideo',
    mp3: 'audio/mpeg',
    wav: 'audio/wav',
    m4a: 'audio/mp4',
    aac: 'audio/aac',
    ogg: 'audio/ogg',
  }
  ipcMain.handle('slides:media-data', (e, slideIndex: number, sourceId: string) => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[slideIndex]
    if (!session || !slide) return null
    const el = slide.elements.find((x) => x.id === sourceId)
    if (!el || el.type !== 'picture') return null
    const media = (
      el as { media?: { kind: 'video' | 'audio'; target?: string; external?: boolean } }
    ).media
    if (!media?.target) return null
    if (media.external) return { kind: media.kind, dataUrl: media.target }
    const bytes = session.opened.archive.readBytes(media.target)
    if (!bytes) return null
    const ext = media.target.split('.').pop()?.toLowerCase() ?? ''
    const mime = AV_MIME[ext] ?? (media.kind === 'video' ? 'video/mp4' : 'audio/mpeg')
    return {
      kind: media.kind,
      dataUrl: `data:${mime};base64,${Buffer.from(bytes).toString('base64')}`,
    }
  })

  ipcMain.handle('slides:get-link', (e, slideIndex: number, sourceId: string) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    return getElementLink(session.opened, slideIndex, sourceId)
  })

  ipcMain.handle('slides:get-slide-links', (e, slideIndex: number) => {
    const session = sessions.get(e.sender.id)
    if (!session) return []
    return getSlideLinks(session.opened, slideIndex).map(({ elementId, target }) => ({
      sourceId: elementId,
      target,
    }))
  })

  ipcMain.handle('slides:get-run-links', (e, slideIndex: number) => {
    const session = sessions.get(e.sender.id)
    if (!session) return []
    return getRunLinks(session.opened, slideIndex).map(({ elementId, ...rest }) => ({
      sourceId: elementId,
      ...rest,
    }))
  })

  ipcMain.handle('slides:get-transition', (e, slideIndex: number) => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[slideIndex]
    return slide ? getSlideTransition(slide) : 'none'
  })

  ipcMain.handle('slides:get-show-timing', (e, slideIndex: number): ShowTiming => {
    const slide = sessions.get(e.sender.id)?.opened.deck.slides[slideIndex]
    if (!slide) return { advanceMs: null, narrationIds: [] }
    const narration = new Set(narrationSpids(slide))
    return {
      advanceMs: getSlideAdvanceTime(slide),
      narrationIds: slide.elements
        .filter((el) => narration.has(elementSpid(el) ?? -1))
        .map((el) => el.id),
    }
  })

  ipcMain.handle('slides:microphone-access', async () => {
    if (process.platform !== 'darwin') return true
    if (systemPreferences.getMediaAccessStatus('microphone') === 'granted') return true
    return systemPreferences.askForMediaAccess('microphone')
  })

  // ── Shape animations (<p:timing>; the spid <-> temporary element id mapping happens here) ──
  ipcMain.handle('slides:get-animations', (e, slideIndex: number): AnimationItem[] => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[slideIndex]
    if (!slide) return []
    const bySpid = new Map<number, (typeof slide.elements)[number]>()
    for (const el of slide.elements) {
      const spid = elementSpid(el)
      if (spid != null && !bySpid.has(spid)) bySpid.set(spid, el)
    }
    const typeLabel: Record<string, string> = {
      text: tm('labelTextBox'),
      shape: tm('labelShape'),
      picture: tm('labelPicture'),
      group: tm('labelGroup'),
      table: tm('labelTable'),
      chart: tm('labelChart'),
      passthrough: tm('labelObject'),
    }
    const out: AnimationItem[] = []
    for (const a of getSlideAnimations(slide)) {
      const el = bySpid.get(a.spid)
      if (!el) continue // Leftover animations whose target shape was deleted are not echoed back
      out.push({
        sourceId: el.id,
        targetName: el.name || typeLabel[el.type] || tm('labelObject'),
        effect: a.effect,
        trigger: a.trigger,
        durationMs: a.durationMs,
        delayMs: a.delayMs,
        // a modelled directional effect carries its own direction; a top wipe must not
        // come back to the player as a bare 'wipe' and play bottom-up
        ...(a.direction != null ? { direction: a.direction } : {}),
        ...(a.motionPath != null ? { motionPath: a.motionPath } : {}),
        ...(a.paragraph != null ? { paragraph: a.paragraph } : {}),
      })
    }
    return out
  })

  // Pairing keys for Morph transitions: sourceId changes on every reparse, so match across pages by cNvPr id/name
  ipcMain.handle('slides:get-shape-keys', (e, slideIndex: number): ShapeKey[] => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[slideIndex]
    if (!slide) return []
    return slide.elements.map((el) => ({
      sourceId: el.id,
      spid: elementSpid(el),
      name: el.name ?? '',
    }))
  })

  // ── Speaker notes / comments (archive surgery, riding on snapshot undo and savePptx) ────
  ipcMain.handle('slides:get-notes', (e, slideIndex: number) => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[slideIndex]
    return session && slide ? getSlideNotes(session.opened.archive, slide.path) : ''
  })

  // ── Show fullscreen: macOS native fullscreen is an animated Space transition, so
  // the slideshow would render windowed for ~1s mid-flight. Instead one call covers
  // everything while the show's black root hides the relayout: the tab view bleeds
  // over the tab strip (shell hook) and the window snaps via simpleFullScreen (same
  // trick as the audience window in presenter-show.ts). The renderer skips HTML
  // fullscreen on macOS entirely. Snap is skipped when the user already fullscreened
  // the window into its own Space; Windows/Linux keep HTML fullscreen (instant there)
  // and only need the bleed. Release is debounced: React strict-mode remounts and
  // presenter→show handoffs flip off→on within a tick, and honoring the off
  // immediately makes the window visibly bounce. ──
  let showFsRelease: ReturnType<typeof setTimeout> | null = null
  ipcMain.handle('slides:show-fullscreen', (e, on: boolean) => {
    // a detached editor window fullscreens itself, never the shell behind it
    const win = hostWindowFor(e.sender)
    if (!win || win.isDestroyed()) return
    const wc = e.sender
    if (showFsRelease) {
      clearTimeout(showFsRelease)
      showFsRelease = null
    }
    if (on) {
      showChrome.setBleed?.(wc, true)
      if (process.platform === 'darwin' && !win.isFullScreen()) {
        win.setFullScreenable(false)
        if (!win.isSimpleFullScreen()) win.setSimpleFullScreen(true)
      }
      // The snap can leave the window's first responder on the shell chrome view, so
      // keys land in the tab-strip renderer (Esc dead until a click on the show).
      // win.focus() must NOT be used here — it focuses the shell renderer itself.
      // Focus the tab's webContents now and once more on the next tick (the snap's
      // responder change lands async). HTML fullscreen used to do this implicitly.
      wc.focus()
      setTimeout(() => {
        if (!wc.isDestroyed()) wc.focus()
      }, 50)
    } else {
      showFsRelease = setTimeout(() => {
        showFsRelease = null
        if (!wc.isDestroyed()) showChrome.setBleed?.(wc, false)
        if (win.isDestroyed()) return
        if (process.platform === 'darwin') {
          if (win.isSimpleFullScreen()) win.setSimpleFullScreen(false)
          win.setFullScreenable(true)
        }
      }, 150)
    }
  })

  // ── Presenter-view multi-screen show (registered inside registerSlidesIpc: shell
  // aggregate mode only calls this function) ──
  registerPresenterIpc()
}

export function createSlidesWindow(openPath?: string | null): BrowserWindow {
  const win = new BrowserWindow({
    width: 1280,
    height: 840,
    title: 'GenOffice Slides',
    ...(process.platform === 'darwin'
      ? { titleBarStyle: 'hiddenInset' as const }
      : {
          titleBarStyle: 'hidden' as const,
          titleBarOverlay: { color: '#ffffff', symbolColor: '#444444', height: 40 },
        }),
    webPreferences: {
      preload: runtime.preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  })
  // Titles are owned by the main process (initial file name, Save As updates);
  // the renderer's static <title> must not overwrite them
  win.on('page-title-updated', (event) => event.preventDefault())

  void win.loadURL(rendererUrl(runtime.rendererDevUrl, 'slides'))

  if (openPath) {
    win.setTitle(basename(openPath))
    win.webContents.once('did-finish-load', async () => {
      try {
        const result = await openAndBuild(win.webContents, openPath, 1280)
        win.webContents.send('slides:opened', result)
      } catch {
        /* ignore */
      }
    })
  }
  return win
}

export function startSlidesStandalone(): void {
  registerRendererScheme()

  if (!app.requestSingleInstanceLock()) {
    app.quit()
    return
  }

  app.on('open-file', (event, path) => {
    event.preventDefault()
    if (app.isReady()) {
      const win = BrowserWindow.getAllWindows()[0]
      if (win) {
        openAndBuild(win.webContents, path, 1280).then((r) =>
          win.webContents.send('slides:opened', r),
        )
        win.focus()
      } else createSlidesWindow(path)
    } else pendingOpenPath = path
  })

  const argPath = process.argv.find((a) => a.toLowerCase().endsWith('.pptx'))
  if (argPath && existsSync(argPath)) pendingOpenPath = argPath

  app.whenReady().then(async () => {
    installRendererProtocol({ slides: join(__dirname, '../renderer') })
    setUiLang(normalizeLang(process.env.GENOFFICE_LANG ?? app.getLocale()))
    registerSlidesIpc()
    createSlidesWindow(pendingOpenPath)
    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) createSlidesWindow()
    })
  })

  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit()
  })
}
