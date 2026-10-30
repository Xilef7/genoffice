/**
 * Ground-truth metrics for complex scripts (Arabic/Hebrew/Thai/Devanagari): HarfBuzz (wasm)
 * shapes against the actual system fonts to compute advance widths. These scripts have ligatures
 * and contextual forms, so per-glyph accumulation (the opentype path) or heuristic estimation
 * both drift from Chromium's actual rendered width; HarfBuzz shares its shaping engine lineage
 * with Chromium, and with displayFamily forcing the draw side to use the same font family,
 * measurement and drawing match pixel-for-pixel.
 *
 * Initialization is async (wasm compilation); when not ready or fonts are missing this returns
 * null and callers fall back to the existing path.
 */

/**
 * Renderer-process ground-truth widths (@100px, scaled linearly). HarfBuzz estimates still
 * deviate from Chromium's actual rendering (AAT/morx fonts, font matching cannot be replicated),
 * so the page's measureText is the final authority: the first layout pass records misses ->
 * refineComplexWidths measures them in batch -> the second layout pass hits the cache entirely.
 */
const gtCache = new Map<string, number>()
const gtPending = new Set<string>()
const gtKey = (text: string, family: string, bold: boolean, italic: boolean) =>
  `${bold ? 'b' : ''}${italic ? 'i' : ''}|${family}|${text}`

/**
 * Send complex-script strings that missed during layout to the renderer's measureText (same
 * engine measures and draws) and backfill the cache. Returns true when new widths were stored,
 * meaning the caller should rebuild the layout. If the page is not ready or the call fails,
 * keep them pending and retry next time.
 */
export async function refineComplexWidths(wc: {
  executeJavaScript(code: string): Promise<unknown>
}): Promise<boolean> {
  if (!gtPending.size) return false
  const items = [...gtPending]
  gtPending.clear()
  const payload = items.map((key) => {
    const [flags = '', family = '', ...rest] = key.split('|')
    return {
      key,
      family,
      text: rest.join('|'),
      bold: flags.includes('b'),
      italic: flags.includes('i'),
    }
  })
  try {
    const widths = (await wc.executeJavaScript(`(async () => {
      const items = ${JSON.stringify(payload)}
      // Ground truth requires the REAL faces: layout may resolve to Office-private
      // fonts Chromium cannot see by name, and this refinement can run before
      // doc-fonts syncs them — measuring the fallback would bake wrong widths in
      const need = new Set(items.map((it) => it.family))
      let privateFams = new Set()
      try {
        const faces = await window.slidesApi.privateFontFaces()
        privateFams = new Set((Array.isArray(faces) ? faces : []).map((f) => f.family))
        for (const f of Array.isArray(faces) ? faces : []) {
          if (!need.has(f.family)) continue
          // Each face registers independently: one bad face must not skip the rest
          try {
            const weight = f.bold ? '700' : '400'
            const style = f.italic ? 'italic' : 'normal'
            let have = false
            document.fonts.forEach((x) => {
              if (x.family === f.family && x.weight === weight && x.style === style) have = true
            })
            if (have) continue
            const data = await window.slidesApi.privateFontData(f.id)
            if (!data) continue
            const face = new FontFace(f.family, data, { weight, style })
            document.fonts.add(face)
            await face.load()
          } catch { /* this face keeps the fallback; check() below skips it */ }
        }
      } catch { /* no private-face listing: system fonts still pass check() */ }
      // fonts.check() returns true for unknown families on modern Chromium
      // (fallback counts as renderable), so it cannot gate this. A family is
      // measurable when it is NOT a private face (main resolved it to a system
      // font Chromium sees by name) or any face of it is registered — missing
      // styles are synthesized identically by canvas drawing and measureText.
      const registered = new Set()
      document.fonts.forEach((x) => registered.add(x.family))
      const ctx = document.createElement('canvas').getContext('2d')
      const out = {}
      for (const it of items) {
        if (privateFams.has(it.family) && !registered.has(it.family)) {
          out[it.key] = null // family not available yet — stays pending
          continue
        }
        // Same font stack as the Konva renderer (generic fallback of konva-adapter displayFontFamily)
        ctx.font = (it.italic ? 'italic ' : '') + (it.bold ? 'bold ' : '') +
          '100px "' + it.family + '", \\'PingFang SC\\', \\'Microsoft YaHei\\', \\'Yu Gothic\\', \\'Malgun Gothic\\', sans-serif'
        out[it.key] = ctx.measureText(it.text).width
      }
      return out
    })()`)) as Record<string, number>
    let added = false
    for (const [k, w] of Object.entries(widths)) {
      if (typeof w === 'number' && Number.isFinite(w)) {
        gtCache.set(k, w)
        added = true
      } else {
        gtPending.add(k) // family didn't resolve yet — retry next refine
      }
    }
    return added
  } catch {
    for (const it of items) gtPending.add(it)
    return false
  }
}

/** Draw font family for complex-script text (same font file as shaping); null if not applicable. */
/**
 * Renderer ground-truth width for non-complex text (same gtCache/refine batch as the
 * complex-script channel). Used for faces whose opentype-side advances drift from
 * Chromium's rendering (variable fonts: HVAR/instancing interpretations differ by a
 * few percent, which visually swallows word spaces). Returns null on a cache miss and
 * queues the string for refineComplexWidths; the second layout pass hits the cache.
 */
export function gtMeasure(
  text: string,
  family: string,
  fontSizePx: number,
  bold = false,
  italic = false,
): number | null {
  const key = gtKey(text, family, bold, italic)
  const gt = gtCache.get(key)
  if (gt != null) return (gt / 100) * fontSizePx
  gtPending.add(key)
  return null
}
