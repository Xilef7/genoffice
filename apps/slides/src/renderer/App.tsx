import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import type {
  RenderFill,
  RenderNode,
  RenderSlide,
} from '@genoffice/pptx-render'
import { baseName } from '../shared/base-name'
import { createImageLoader } from './image-loader'
import { syncPrivateFonts } from './doc-fonts'
import { SlideShowView } from './components/SlideShowView'
import { PresenterView } from './components/PresenterView'
import { showToast } from './components/toast-bus'
import { t } from './i18n/locale'
import type {
  ActionCtx,
  SlideShowState,
} from './action-context'
import { FIT_WIDTH } from './app-constants'
import * as showActions from './show-actions'
import { normalizeSelection } from '../shared/slide-selection'

/** Resizable thumbnail sidebar: drag the right edge; width persisted, clamped to a sane range */
const THUMBS_W_KEY = 'ai-slides-thumbs-width'
const THUMBS_W_DEFAULT = 120
const THUMBS_W_MIN = 110

function clampThumbsW(w: number): number {
  return Math.min(Math.max(w, THUMBS_W_MIN), Math.min(400, Math.round(window.innerWidth * 0.4)))
}

function loadThumbsW(): number {
  const saved = Number(localStorage.getItem(THUMBS_W_KEY))
  return Number.isFinite(saved) && saved > 0 ? clampThumbsW(saved) : THUMBS_W_DEFAULT
}

export function App() {
  const [slides, setSlides] = useState<RenderSlide[]>([])
  // Layouts may reference Office-private fonts (resolved in main); register them as FontFaces
  useEffect(() => {
    if (slides.length) void syncPrivateFonts()
  }, [slides])
  const [path, setPath] = useState<string | null>(null)
  // Deck references catalog fonts that are missing locally → one-click download banner
  const [missingFonts, setMissingFonts] = useState<string[]>([])
  const [fontBannerBusy, setFontBannerBusy] = useState(false)
  const missingFontsDismissed = useRef(false)
  const refreshMissingFonts = useCallback(() => {
    if (missingFontsDismissed.current) return
    window.slidesApi
      .fontMissing?.()
      .then((m) => setMissingFonts(m ?? []))
      .catch(() => {})
  }, [])
  const hasSlides = slides.length > 0
  useEffect(() => {
    missingFontsDismissed.current = false
    if (hasSlides) refreshMissingFonts()
  }, [path, hasSlides, refreshMissingFonts])
  // Font store changed (download / local install): re-register FontFaces, refresh the banner
  useEffect(
    () =>
      window.slidesApi.onFontsChanged?.(() => {
        void syncPrivateFonts()
        refreshMissingFonts()
      }),
    [refreshMissingFonts],
  )
  const downloadMissingFonts = useCallback(async () => {
    setFontBannerBusy(true)
    try {
      for (const f of missingFonts) await window.slidesApi.fontDownload?.(f)
    } finally {
      setFontBannerBusy(false)
      refreshMissingFonts()
    }
  }, [missingFonts, refreshMissingFonts])
  const [current, setCurrent] = useState(0)
  const [slideSelRaw, setSelectedSlides] = useState<number[]>([0])
  /** Latest-state bundle for the extracted action modules; refreshed every render (see action-context.ts). */
  const ctxRef = useRef<ActionCtx>(null as unknown as ActionCtx)
  const [zoom, setZoom] = useState(1)
  const [status, setStatus] = useState('')
  // Status messages auto-dismiss after 4s: operation feedback needs only a brief showing; persistent info (page number/file name) lives on the left and in the title bar
  useEffect(() => {
    if (!status) return
    const t = window.setTimeout(() => setStatus(''), 4000)
    return () => window.clearTimeout(t)
  }, [status])
  // ── Thumbnail sidebar width (drag the divider to resize; persisted) ─────────
  const [thumbsW, setThumbsW] = useState(loadThumbsW)
  const thumbsListRef = useRef<HTMLDivElement | null>(null)
  // Re-clamp when the window shrinks (max is 40% of the window), like the AI panel
  useEffect(() => {
    const onResize = () => setThumbsW((w) => clampThumbsW(w))
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [])
  const consumePendingRef = useRef<ReturnType<typeof window.slidesApi.consumePendingOpen> | null>(
    null,
  )
  const bootHandledRef = useRef(false)
  const [images, setImages] = useState<Map<string, HTMLImageElement>>(new Map())
  const imageLoaderRef = useRef<ReturnType<typeof createImageLoader> | null>(null)
  const stageWrapRef = useRef<HTMLDivElement | null>(null)
  const [stageViewportSize, setStageViewportSize] = useState({ w: 0, h: 0 })
  // Follow slide changes made outside the list (arrow keys, canvas paging); viewMode/showThumbs
  // re-run it because the list remounts at scroll 0 when the normal view returns
  useEffect(() => {
    thumbsListRef.current?.querySelector('.thumb.active')?.scrollIntoView({ block: 'nearest' })
  }, [current])
  // ── Slide show: when non-null, covers the whole window (startAt is the start page index;
  //    customOrder = custom show playback sequence; rehearse = rehearsal timing mode) ────────
  const [slideShow, setSlideShow] = useState<SlideShowState | null>(null)
  /** Presenter view (single-window version; mutually exclusive with slideShow) */
  const [presenter, setPresenter] = useState<{ startAt: number } | null>(null)

  const slide = slides[current]
  // Rail multi-selection: stale entries (page deleted, anchor moved elsewhere) collapse to the anchor
  const selectedSlides = useMemo(
    () => normalizeSelection(slideSelRaw, current, slides.length),
    [slideSelRaw, current, slides.length],
  )
  useEffect(() => {
    // Compare by content: with no deck the normalizer yields a fresh [0] every render
    const same =
      selectedSlides.length === slideSelRaw.length &&
      selectedSlides.every((v, k) => v === slideSelRaw[k])
    if (!same) setSelectedSlides(selectedSlides)
  }, [selectedSlides, slideSelRaw])
  const hasDoc = !!slide

  // Uncapped proportional fit ratio from the stage container's measured size
  // (fall back to a window estimate before mount)
  const rawFit = useCallback(
    (s?: RenderSlide) => {
      if (!s) return 1
      const el = stageWrapRef.current
      const viewportW =
        el?.clientWidth ||
        stageViewportSize.w ||
        window.innerWidth - thumbsW - 34
      const viewportH = el?.clientHeight || stageViewportSize.h || window.innerHeight - 150
      const availW = viewportW - 56
      // -72: vertical padding is 48 (AI-bar headroom) + 32, minus the same 8px slack as width
      const availH = viewportH - 72
      return Math.min(availW / s.widthPx, availH / s.heightPx)
    },
    [stageViewportSize.h, stageViewportSize.w, thumbsW],
  )
  /** Fit zoom for display: rawFit clamped to the 0.1–1.5 auto-fit range */
  const fitZoom = useCallback(
    (s?: RenderSlide) => Math.min(1.5, Math.max(0.1, rawFit(s))),
    [rawFit],
  )
  // The Konva stage includes a transparent bleed around the slide for editing
  // off-page objects. That bleed must not create scrollbars while the slide
  // itself still fits; once the nominal slide overflows, normal panning resumes.
  const stageFitsViewport = !slide || zoom <= rawFit(slide) + 0.001
  /** Fit-pending flag after open: consumed when the editor's first frame mounts (stage container measurable) */
  const needsFitRef = useRef(false)
  /** Last auto-fit value: if current zoom still equals it → treated as "fit mode", re-fit on size changes */
  const lastFitRef = useRef<number | null>(null)
  const zoomLiveRef = useRef(1)
  // Layout effect, not passive: committing a zoom step past fit makes scrollbars
  // appear, and the fit-keeper ResizeObserver fires BEFORE passive effects run.
  // With a stale ref it still reads the old fit value, decides "fit mode", and
  // snaps the fresh zoom straight back — a single +/− step from fit never sticks.
  useLayoutEffect(() => {
    zoomLiveRef.current = zoom
  }, [zoom])
  const slideLiveRef = useRef<RenderSlide | undefined>(undefined)
  useEffect(() => {
    slideLiveRef.current = slide
  }, [slide])
  // Live rawFit for the zoom-gesture code (previewZoom is dependency-free)
  const rawFitRef = useRef(rawFit)
  useEffect(() => {
    rawFitRef.current = rawFit
  }, [rawFit])

  /** While the slide fits the viewport, the centered position is the only valid one —
   * overflow is hidden, so the user cannot correct a stray scroll offset, and the
   * Konva bleed keeps scrollHeight above clientHeight so the browser never clamps
   * it back to 0 on its own. A zoom gesture's anchoring math can leave such an
   * offset behind (large Ctrl+wheel steps on Windows made the slide sit clipped
   * under the ribbon); zero it whenever the target zoom is at/below fit. */
  const resetFitScroll = useCallback((z: number) => {
    if (z > rawFitRef.current(slideLiveRef.current) + 0.001) return
    const wrap = stageWrapRef.current
    if (wrap && (wrap.scrollLeft !== 0 || wrap.scrollTop !== 0)) {
      wrap.scrollLeft = 0
      wrap.scrollTop = 0
    }
  }, [])

  // Non-gesture zoom paths (zoom-to-fit button, open-time fit, resize re-fit) land here
  useLayoutEffect(() => {
    if (stageFitsViewport) resetFitScroll(zoom)
  }, [stageFitsViewport, zoom, resetFitScroll])

  useLayoutEffect(() => {
    if (!slide || !needsFitRef.current) return
    needsFitRef.current = false
    const z = fitZoom(slide)
    lastFitRef.current = z
    setZoom(z)
  }, [slide, fitZoom])

  // On container size changes (window/sidebar/thumbnail toggles): follow with a
  // re-fit while in fit mode, and clamp any manual zoom back down to fit whenever
  // the container gets too small — the canvas must never overflow the pane. A
  // manual zoom smaller than fit is left alone.
  useEffect(() => {
    const el = stageWrapRef.current
    if (!hasDoc || !el) return
    // Border-box size at the previous callback. Zooming past fit makes classic
    // scrollbars appear, which shrinks the observed content box without touching
    // the border box; treating that as "container got smaller" snapped every
    // zoom-in above fit straight back to fit. Only a border-box change is a real
    // window/pane resize allowed to clamp a manual zoom.
    let outerBox: { w: number; h: number } | null = null
    const ro = new ResizeObserver(() => {
      const nextSize = { w: el.clientWidth, h: el.clientHeight }
      setStageViewportSize((previous) =>
        previous.w === nextSize.w && previous.h === nextSize.h ? previous : nextSize,
      )
      const prevOuter = outerBox
      outerBox = { w: el.offsetWidth, h: el.offsetHeight }
      const z = fitZoom(slideLiveRef.current)
      const lf = lastFitRef.current
      const inFitMode = lf != null && Math.abs(zoomLiveRef.current - lf) <= 0.001
      if (!inFitMode) {
        // The overflow test uses the uncapped ratio: on large windows a manual
        // zoom above the 1.5 fit cap can still fit and must not be wiped
        if (zoomLiveRef.current <= rawFit(slideLiveRef.current) + 0.001) return
        const outerResized =
          prevOuter != null && (outerBox.w !== prevOuter.w || outerBox.h !== prevOuter.h)
        if (!outerResized) return
      }
      lastFitRef.current = z
      setZoom(z)
    })
    ro.observe(el)
    return () => ro.disconnect()
    // viewMode: reading/sorter unmount the editor, so the observer must re-bind
    // to the remounted .stage-wrap or window resizes stop re-fitting the canvas
  }, [hasDoc, fitZoom, rawFit])

  const applyOpen = useCallback(
    (result: { path: string; slides: RenderSlide[]; defaultFont?: string } | null) => {
      if (!result) return
      setSlides(result.slides)
      setPath(result.path)
      setCurrent(0)
      setSelectedSlides([0])
      needsFitRef.current = true
      setStatus(
        result.path
          ? t('appStatusOpened', {
              name: baseName(result.path),
              count: result.slides.length,
            })
          : t('appStatusNewBlank'),
      )
    },
    [],
  )

  // Shared session (a second window on the same file): apply the other window's
  // edits without touching local selection or an open editor — the ids are stable,
  // so the selection stays valid; a same-element conflict resolves last-write-wins.
  useEffect(
    () =>
      window.slidesApi.onDeckChanged?.(({ slides: all }) => {
        setSlides(all)
        setCurrent((c) => Math.min(c, Math.max(0, all.length - 1)))
      }),
    [],
  )

  const newBlank = useCallback(async () => {
    const r = await window.slidesApi.newBlank(FIT_WIDTH)
    applyOpen(r)
    return r
  }, [applyOpen])

  // Files are pushed open by the main process (double-click/command line); with no pending file the
  // window lands directly in the editor on a fresh blank deck (the AI panel carries the generate-from-prompt flow).
  // StrictMode runs the mount effect twice, but the pending queue can only be consumed once, so the
  // consume Promise is stored in a shared ref and its result is processed only once.
  useEffect(() => {
    // newBlank itself can fail (IPC/main-process error); one retry for transient
    // hiccups, then surface the error — otherwise the tab silently sticks on
    // "Opening…" forever and the click that created it looks like a no-op
    const bootBlank = () =>
      newBlank().catch((err: unknown) => {
        void newBlank().catch(() => {
          showToast(err instanceof Error ? err.message : String(err), 'error')
        })
      })
    const off = window.slidesApi.onOpened((r) => applyOpen(r))
    consumePendingRef.current ??= window.slidesApi.consumePendingOpen(FIT_WIDTH)
    void consumePendingRef.current
      .then((r) => {
        if (bootHandledRef.current) return
        bootHandledRef.current = true
        if (r) applyOpen(r)
        else void bootBlank()
      })
      // Open failures (corrupt file etc.) also land on a blank deck, or it stays at "Opening…" forever
      .catch(() => {
        if (bootHandledRef.current) return
        bootHandledRef.current = true
        void bootBlank()
      })
    return off
  }, [applyOpen, newBlank])

  // Path changed outside this renderer (shell Home list rename, or an MCP save that
  // wrote the session to disk) → sync the title-bar path and ask the session
  // whether it is still dirty rather than assuming
  useEffect(
    () =>
      window.slidesApi.onRenamed((p) => {
        setPath(p)
      }),
    [],
  )

  const openDialog = useCallback(async () => {
    const r = await window.slidesApi.openPptx(FIT_WIDTH)
    applyOpen(r)
  }, [applyOpen])

  // ── Slide show tab (show-actions.ts): start show / presenter view / hide slide ──
  const startSlideShow = useCallback(
    (fromStart: boolean) => showActions.startSlideShow(ctxRef.current, fromStart),
    [],
  )
  const exitSlideShow = useCallback(
    (lastIndex: number) => showActions.exitSlideShow(ctxRef.current, lastIndex),
    [],
  )

  const startPresenterView = useCallback(
    (fromStart: boolean) => showActions.startPresenterView(ctxRef.current, fromStart),
    [],
  )
  const exitPresenterView = useCallback(
    (lastIndex: number) => showActions.exitPresenterView(ctxRef.current, lastIndex),
    [],
  )
  const switchPresenterToShow = useCallback(
    (lastIndex: number) => showActions.switchPresenterToShow(ctxRef.current, lastIndex),
    [],
  )

  // Load images (dataUrl → HTMLImageElement): picture elements + picture fills + background images
  // (recursing into group children and decoration nodes). Walk all pages so thumbnails get
  // background images/pictures of non-current pages (otherwise unvisited pages' thumbnails are blank).
  useEffect(() => {
    if (slides.length === 0) return
    const urls = new Set<string>()
    const addFillUrl = (fill: RenderFill | undefined) => {
      if (fill && fill.kind === 'image' && fill.dataUrl) urls.add(fill.dataUrl)
    }
    const addBulletUrls = (
      text: { lines: Array<{ runs: Array<{ image?: string }> }> } | undefined,
    ) => {
      for (const l of text?.lines ?? []) for (const r of l.runs) if (r.image) urls.add(r.image)
    }
    const walk = (nodes: readonly RenderNode[]) => {
      for (const n of nodes) {
        if (n.type === 'picture' && n.dataUrl) urls.add(n.dataUrl)
        if (n.type === 'shape' || n.type === 'text') {
          if (n.fill) addFillUrl(n.fill)
          addBulletUrls(n.text)
        }
        if (n.type === 'chart') {
          addFillUrl((n as { bgFill?: RenderFill }).bgFill)
          for (const b of (n as { bars: Array<{ fill?: RenderFill }> }).bars)
            if (b.fill) addFillUrl(b.fill)
        }
        if (n.type === 'group' && Array.isArray(n.children)) walk(n.children)
        if (n.type === 'table' && Array.isArray(n.cells)) {
          for (const c of n.cells) {
            if (c.fill) addFillUrl(c.fill)
            addBulletUrls(c.text)
          }
        }
      }
    }
    for (const s of slides) {
      addFillUrl(s.background)
      walk(s.nodes)
    }
    if (!imageLoaderRef.current) {
      imageLoaderRef.current = createImageLoader((entries) => {
        setImages((prev) => {
          const m = new Map(prev)
          for (const [k, v] of entries) m.set(k, v)
          return m
        })
      })
    }
    imageLoaderRef.current.load(urls)
  }, [slides])
  // Dispose on unmount and clear the ref so a remount (e.g. React Strict Mode's
  // dev double-mount) lazily recreates a fresh loader instead of reusing a disposed one.
  useEffect(
    () => () => {
      imageLoaderRef.current?.dispose()
      imageLoaderRef.current = null
    },
    [],
  )

  // Refresh the action-module context every render so extracted actions never see stale state
  ctxRef.current = {
    slides,
    setSlides,
    current,
    setCurrent,
    slide,
    path,
    setPath,
    setStatus,
    slideShow,
    setSlideShow,
    presenter,
    setPresenter,
  }

  return (
    <div className="app">
      <button onClick={() => openDialog()}>Open</button>
      <button onClick={() => startSlideShow(true)}>Slide Show</button>
      <button onClick={() => startPresenterView(true)}>Presenter View</button>
      {missingFonts.length > 0 && (
        <div className="font-missing-banner">
          <span className="fmb-text">
            {t('fontMissingBanner')}
            <span className="fmb-fonts">{missingFonts.join(', ')}</span>
          </span>
          <button
            className="fmb-download"
            disabled={fontBannerBusy}
            onClick={() => void downloadMissingFonts()}
          >
            {fontBannerBusy ? t('ribbonFontDownloading') : t('fontMissingDownloadAll')}
          </button>
          <button
            className="fmb-dismiss"
            onClick={() => {
              missingFontsDismissed.current = true
              setMissingFonts([])
            }}
          >
            {t('fontMissingDismiss')}
          </button>
        </div>
      )}
      {slideShow && slides.length > 0 && (
        <SlideShowView
          slides={slides}
          images={images}
          startAt={slideShow.startAt}
          customOrder={slideShow.customOrder}
          rehearseMode={slideShow.rehearse}
          recordMode={slideShow.record}
          onExit={exitSlideShow}
        />
      )}
      {presenter && slides.length > 0 && (
        <PresenterView
          slides={slides}
          images={images}
          startAt={presenter.startAt}
          onExit={exitPresenterView}
          onUseSlideShow={switchPresenterToShow}
        />
      )}
    </div>
  )
}

/** Clamp a zoom-anchor coordinate into the viewport span; undefined passes through (→ center). */
function anchorClamp(v: number | undefined, min: number, max: number): number | undefined {
  return v == null ? undefined : Math.min(Math.max(v, min), max)
}
