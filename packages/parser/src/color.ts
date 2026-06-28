import type { Color, SchemeColor, SolidColor, ThemeColors } from './types.js'
import { attr, attrNum, get, toArray } from './xml.js'

// ─── Raw color parsing from XML nodes ────────────────────────────────────────

function readColorMods(node: unknown): Pick<SchemeColor, 'lumMod' | 'lumOff' | 'shade' | 'tint' | 'alpha'> {
  const mods: Pick<SchemeColor, 'lumMod' | 'lumOff' | 'shade' | 'tint' | 'alpha'> = {}

  const lumMod = attrNum(get(node, 'a:lumMod'), 'val')
  if (lumMod !== undefined) mods.lumMod = lumMod / 100000

  const lumOff = attrNum(get(node, 'a:lumOff'), 'val')
  if (lumOff !== undefined) mods.lumOff = lumOff / 100000

  const shade = attrNum(get(node, 'a:shade'), 'val')
  if (shade !== undefined) mods.shade = shade / 100000

  const tint = attrNum(get(node, 'a:tint'), 'val')
  if (tint !== undefined) mods.tint = tint / 100000

  const alpha = attrNum(get(node, 'a:alpha'), 'val')
  if (alpha !== undefined) mods.alpha = alpha / 100000

  return mods
}

/**
 * Parse a color from any OOXML color container node (e.g. a:solidFill child,
 * a:srgbClr, a:schemeClr, a:sysClr, a:prstClr, etc.)
 */
export function parseColor(node: unknown): Color | undefined {
  if (node === null || typeof node !== 'object') return undefined

  const n = node as Record<string, unknown>

  // <a:srgbClr val="FF0000"/>
  if ('a:srgbClr' in n) {
    const hex = attr(n['a:srgbClr'], 'val') ?? '000000'
    const mods = readColorMods(n['a:srgbClr'])
    const raw: SolidColor = { type: 'solid', hex: hex.toLowerCase(), alpha: 100 }
    if (mods.alpha !== undefined) raw.alpha = mods.alpha * 100
    // Apply shade/tint/lum directly to solid colors at parse time
    return applyModsToSolid(raw, mods)
  }

  // <a:sysClr lastClr="FFFFFF" val="windowText"/>
  if ('a:sysClr' in n) {
    const lastClr = attr(n['a:sysClr'], 'lastClr') ?? '000000'
    const mods = readColorMods(n['a:sysClr'])
    const raw: SolidColor = { type: 'solid', hex: lastClr.toLowerCase(), alpha: 100 }
    return applyModsToSolid(raw, mods)
  }

  // <a:prstClr val="black"/>
  if ('a:prstClr' in n) {
    const name = attr(n['a:prstClr'], 'val') ?? 'black'
    const hex = PRESET_COLORS[name] ?? '000000'
    const mods = readColorMods(n['a:prstClr'])
    const raw: SolidColor = { type: 'solid', hex, alpha: 100 }
    return applyModsToSolid(raw, mods)
  }

  // <a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr>
  if ('a:schemeClr' in n) {
    const token = attr(n['a:schemeClr'], 'val') ?? 'dk1'
    const mods = readColorMods(n['a:schemeClr'])
    return { type: 'scheme', token, ...mods }
  }

  // Direct scheme color node (already unwrapped)
  const token = attr(node, 'val')
  if (token) {
    const mods = readColorMods(node)
    return { type: 'scheme', token, ...mods }
  }

  return undefined
}

/** Parse <a:solidFill> child and return the Color inside. */
export function parseSolidFillColor(solidFillNode: unknown): Color | undefined {
  return parseColor(solidFillNode)
}

// ─── Color resolution against a theme ────────────────────────────────────────

const SCHEME_TOKEN_MAP: Record<string, keyof ThemeColors> = {
  dk1: 'dk1',
  dk2: 'dk2',
  lt1: 'lt1',
  lt2: 'lt2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hlink: 'hlink',
  folHlink: 'folHlink',
  // Aliases
  tx1: 'dk1',
  tx2: 'dk2',
  bg1: 'lt1',
  bg2: 'lt2',
}

/**
 * Fully resolve a Color to a hex string using the theme.
 * Returns '#000000' for unresolvable colors.
 */
export function resolveColor(color: Color | undefined, themeColors: ThemeColors): string {
  if (!color) return '#000000'

  if (color.type === 'solid') {
    const alpha = color.alpha < 100 ? Math.round((color.alpha / 100) * 255) : undefined
    const hex = `#${color.hex}`
    if (alpha !== undefined && alpha < 255) {
      return `${hex}${alpha.toString(16).padStart(2, '0')}`
    }
    return hex
  }

  // SchemeColor — resolve via theme
  const themeKey = SCHEME_TOKEN_MAP[color.token]
  let hex = themeKey ? (themeColors[themeKey] ?? '000000') : '000000'

  // Apply modifications
  if (color.lumMod !== undefined || color.lumOff !== undefined || color.shade !== undefined || color.tint !== undefined) {
    hex = applyModsToHex(hex, {
      lumMod: color.lumMod,
      lumOff: color.lumOff,
      shade: color.shade,
      tint: color.tint,
    })
  }

  if (color.alpha !== undefined && color.alpha < 1) {
    const alpha = Math.round(color.alpha * 255)
    return `#${hex}${alpha.toString(16).padStart(2, '0')}`
  }

  return `#${hex}`
}

// ─── Color math helpers ───────────────────────────────────────────────────────

function applyModsToSolid(
  color: SolidColor,
  mods: Pick<SchemeColor, 'lumMod' | 'lumOff' | 'shade' | 'tint' | 'alpha'>,
): SolidColor {
  let hex = applyModsToHex(color.hex, mods)
  const alpha = mods.alpha !== undefined ? mods.alpha * 100 : color.alpha
  return { type: 'solid', hex, alpha }
}

function applyModsToHex(
  hex: string,
  mods: Pick<SchemeColor, 'lumMod' | 'lumOff' | 'shade' | 'tint'>,
): string {
  if (!mods.lumMod && !mods.lumOff && !mods.shade && !mods.tint) return hex

  let [h, s, l] = hexToHsl(hex)

  if (mods.lumMod !== undefined) l = l * mods.lumMod
  if (mods.lumOff !== undefined) l = l + mods.lumOff
  if (mods.shade !== undefined) l = l * mods.shade
  if (mods.tint !== undefined) l = l + (1 - l) * (1 - mods.tint)

  l = Math.max(0, Math.min(1, l))
  return hslToHex(h, s, l)
}

function hexToHsl(hex: string): [number, number, number] {
  const r = Number.parseInt(hex.slice(0, 2), 16) / 255
  const g = Number.parseInt(hex.slice(2, 4), 16) / 255
  const b = Number.parseInt(hex.slice(4, 6), 16) / 255

  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  let h = 0
  let s = 0
  const l = (max + min) / 2

  if (max !== min) {
    const d = max - min
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
    switch (max) {
      case r:
        h = ((g - b) / d + (g < b ? 6 : 0)) / 6
        break
      case g:
        h = ((b - r) / d + 2) / 6
        break
      case b:
        h = ((r - g) / d + 4) / 6
        break
    }
  }

  return [h, s, l]
}

function hslToHex(h: number, s: number, l: number): string {
  let r: number
  let g: number
  let b: number

  if (s === 0) {
    r = g = b = l
  } else {
    const hue2rgb = (p: number, q: number, t: number) => {
      const tt = ((t % 1) + 1) % 1
      if (tt < 1 / 6) return p + (q - p) * 6 * tt
      if (tt < 1 / 2) return q
      if (tt < 2 / 3) return p + (q - p) * (2 / 3 - tt) * 6
      return p
    }
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s
    const p = 2 * l - q
    r = hue2rgb(p, q, h + 1 / 3)
    g = hue2rgb(p, q, h)
    b = hue2rgb(p, q, h - 1 / 3)
  }

  const toHex = (x: number) =>
    Math.round(x * 255)
      .toString(16)
      .padStart(2, '0')

  return `${toHex(r)}${toHex(g)}${toHex(b)}`
}

// ─── Preset color names ───────────────────────────────────────────────────────

const PRESET_COLORS: Record<string, string> = {
  black: '000000',
  white: 'ffffff',
  red: 'ff0000',
  green: '008000',
  blue: '0000ff',
  yellow: 'ffff00',
  cyan: '00ffff',
  magenta: 'ff00ff',
  gray: '808080',
  grey: '808080',
  silver: 'c0c0c0',
  maroon: '800000',
  olive: '808000',
  navy: '000080',
  purple: '800080',
  teal: '008080',
  orange: 'ffa500',
  pink: 'ffc0cb',
  brown: 'a52a2a',
  lime: '00ff00',
  indigo: '4b0082',
  violet: 'ee82ee',
  gold: 'ffd700',
  coral: 'ff7f50',
  salmon: 'fa8072',
  khaki: 'f0e68c',
  lavender: 'e6e6fa',
  chocolate: 'd2691e',
  tomato: 'ff6347',
  crimson: 'dc143c',
  darkBlue: '00008b',
  darkCyan: '008b8b',
  darkGray: 'a9a9a9',
  darkGrey: 'a9a9a9',
  darkGreen: '006400',
  darkOrange: 'ff8c00',
  darkRed: '8b0000',
  darkViolet: '9400d3',
  deeppink: 'ff1493',
  dimgray: '696969',
  dimgrey: '696969',
  forestgreen: '228b22',
  hotpink: 'ff69b4',
  limegreen: '32cd32',
  mediumblue: '0000cd',
  mediumvioletred: 'c71585',
  midnightblue: '191970',
  orangered: 'ff4500',
  royalblue: '4169e1',
  seagreen: '2e8b57',
  sienna: 'a0522d',
  skyblue: '87ceeb',
  slateblue: '6a5acd',
  slategray: '708090',
  steelblue: '4682b4',
}

// ─── Gradient stop parsing ────────────────────────────────────────────────────

export function parseGradientStops(gsLst: unknown) {
  return toArray(gsLst as unknown[]).map((gs) => {
    const pos = (attrNum(gs, 'pos') ?? 0) / 100000
    const color = parseColor(gs) ?? { type: 'solid' as const, hex: '000000', alpha: 100 }
    return { position: pos, color }
  })
}
