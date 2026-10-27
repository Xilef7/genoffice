import { contextBridge, ipcRenderer } from 'electron'
import type { IpcRendererEvent } from 'electron'
import type { RenderSlide } from '@genoffice/pptx-render'
import { installDropOpenBridge } from '@genoffice/electron-utils/drop-open'
import type {
  AudienceNavAction,
  ShowInkEvent,
  ShowSyncState,
  OpenResult,
  SlidesApi,
  UiTheme,
} from '../shared/ipc'

const api: SlidesApi = {
  getLanguage: () => ipcRenderer.invoke('app:get-language'),
  onLanguageChanged: (handler) => {
    const listener = (
      _event: IpcRendererEvent,
      lang: 'zh' | 'en' | 'ja' | 'ko' | 'fr' | 'de' | 'es' | 'th' | 'id' | 'ru' | 'ar',
    ) => handler(lang)
    ipcRenderer.on('app:language-changed', listener)
    return () => ipcRenderer.removeListener('app:language-changed', listener)
  },
  getTheme: () => ipcRenderer.invoke('app:get-theme'),
  onThemeChanged: (handler) => {
    const listener = (_event: IpcRendererEvent, theme: UiTheme) => handler(theme)
    ipcRenderer.on('app:theme-changed', listener)
    return () => ipcRenderer.removeListener('app:theme-changed', listener)
  },
  setShowFullScreen: (on) => ipcRenderer.invoke('slides:show-fullscreen', on),
  privateFontFaces: () => ipcRenderer.invoke('slides:private-font-faces'),
  privateFontData: (id) => ipcRenderer.invoke('slides:private-font-data', id),
  fontCatalog: () => ipcRenderer.invoke('slides:font-catalog'),
  fontDownload: (family) => ipcRenderer.invoke('slides:font-download', family),
  fontInstallLocal: () => ipcRenderer.invoke('slides:font-install-local'),
  fontMissing: () => ipcRenderer.invoke('slides:font-missing'),
  onFontsChanged: (handler) => {
    const listener = () => handler()
    ipcRenderer.on('slides:fonts-changed', listener)
    return () => ipcRenderer.removeListener('slides:fonts-changed', listener)
  },
  openPptx: (fitWidthPx) => ipcRenderer.invoke('slides:open', fitWidthPx),
  consumePendingOpen: (fitWidthPx) => ipcRenderer.invoke('slides:consume-pending-open', fitWidthPx),
  newBlank: (fitWidthPx) => ipcRenderer.invoke('slides:new-blank', fitWidthPx),
  getRenderSlides: () => ipcRenderer.invoke('slides:get-render-slides'),
  getMediaData: (slideIndex: number, sourceId: string) =>
    ipcRenderer.invoke('slides:media-data', slideIndex, sourceId),
  getSlideLinks: (slideIndex: number) => ipcRenderer.invoke('slides:get-slide-links', slideIndex),
  getRunLinks: (slideIndex: number) => ipcRenderer.invoke('slides:get-run-links', slideIndex),
  getTransition: (slideIndex: number) => ipcRenderer.invoke('slides:get-transition', slideIndex),
  getShowTiming: (slideIndex: number) => ipcRenderer.invoke('slides:get-show-timing', slideIndex),
  microphoneAccess: () => ipcRenderer.invoke('slides:microphone-access'),
  getAnimations: (slideIndex: number) => ipcRenderer.invoke('slides:get-animations', slideIndex),
  getShapeKeys: (slideIndex: number) => ipcRenderer.invoke('slides:get-shape-keys', slideIndex),
  getNotes: (slideIndex: number) => ipcRenderer.invoke('slides:get-notes', slideIndex),
  onDeckChanged: (
    handler: (state: { slides: RenderSlide[]; size: { cx: number; cy: number } }) => void,
  ) => {
    const listener = (
      _e: IpcRendererEvent,
      state: { slides: RenderSlide[]; size: { cx: number; cy: number } },
    ) => handler(state)
    ipcRenderer.on('slides:deck-changed', listener)
    return () => ipcRenderer.removeListener('slides:deck-changed', listener)
  },
  onOpened: (handler: (result: OpenResult) => void) => {
    const listener = (_e: IpcRendererEvent, result: OpenResult) => handler(result)
    ipcRenderer.on('slides:opened', listener)
    return () => ipcRenderer.removeListener('slides:opened', listener)
  },
  onRenamed: (handler: (newPath: string) => void) => {
    const listener = (_e: IpcRendererEvent, newPath: string) => handler(newPath)
    ipcRenderer.on('slides:renamed', listener)
    return () => ipcRenderer.removeListener('slides:renamed', listener)
  },
  presenterStart: () => ipcRenderer.invoke('slides:presenter-start'),
  presenterSync: (state: ShowSyncState) => ipcRenderer.send('slides:presenter-sync', state),
  presenterInk: (ev: ShowInkEvent) => ipcRenderer.send('slides:presenter-ink', ev),
  presenterSwap: () => ipcRenderer.invoke('slides:presenter-swap'),
  presenterEnd: () => ipcRenderer.invoke('slides:presenter-end'),
  audienceReady: () => ipcRenderer.invoke('slides:audience-ready'),
  audienceNav: (action: AudienceNavAction) => ipcRenderer.send('slides:audience-nav', action),
  onShowSync: (handler: (state: ShowSyncState) => void) => {
    const listener = (_e: IpcRendererEvent, state: ShowSyncState) => handler(state)
    ipcRenderer.on('slides:show-sync', listener)
    return () => ipcRenderer.removeListener('slides:show-sync', listener)
  },
  onShowInk: (handler: (ev: ShowInkEvent) => void) => {
    const listener = (_e: IpcRendererEvent, ev: ShowInkEvent) => handler(ev)
    ipcRenderer.on('slides:show-ink', listener)
    return () => ipcRenderer.removeListener('slides:show-ink', listener)
  },
  onAudienceNav: (handler: (action: AudienceNavAction) => void) => {
    const listener = (_e: IpcRendererEvent, action: AudienceNavAction) => handler(action)
    ipcRenderer.on('slides:audience-nav', listener)
    return () => ipcRenderer.removeListener('slides:audience-nav', listener)
  },
}

contextBridge.exposeInMainWorld('slidesApi', api)

// open documents dragged from the OS onto this tab as a new shell tab
installDropOpenBridge()
