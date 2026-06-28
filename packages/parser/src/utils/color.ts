export function hexToRgb(hex: string): { r: number; g: number; b: number } {
  const cleaned = hex.replace(/^#/, '');
  if (cleaned.length !== 6 && cleaned.length !== 3) return { r: 0, g: 0, b: 0 };
  const full = cleaned.length === 3
    ? cleaned[0]! + cleaned[0]! + cleaned[1]! + cleaned[1]! + cleaned[2]! + cleaned[2]!
    : cleaned;
  const num = parseInt(full, 16);
  return { r: (num >> 16) & 0xff, g: (num >> 8) & 0xff, b: num & 0xff };
}

export function rgbToHex(r: number, g: number, b: number): string {
  const clamp = (v: number) => Math.max(0, Math.min(255, Math.round(v)));
  return '#' + [clamp(r), clamp(g), clamp(b)].map((c) => c.toString(16).padStart(2, '0')).join('');
}

export function rgbToHsl(r: number, g: number, b: number): { h: number; s: number; l: number } {
  const rn = r / 255, gn = g / 255, bn = b / 255;
  const max = Math.max(rn, gn, bn), min = Math.min(rn, gn, bn);
  const l = (max + min) / 2;
  let h = 0, s = 0;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    switch (max) {
      case rn: h = ((gn - bn) / d + (gn < bn ? 6 : 0)) * 60; break;
      case gn: h = ((bn - rn) / d + 2) * 60; break;
      case bn: h = ((rn - gn) / d + 4) * 60; break;
    }
  }
  return { h, s, l };
}

export function hslToRgb(h: number, s: number, l: number): { r: number; g: number; b: number } {
  h = ((h % 360) + 360) % 360;
  s = Math.max(0, Math.min(1, s));
  l = Math.max(0, Math.min(1, l));
  if (s === 0) { const v = Math.round(l * 255); return { r: v, g: v, b: v }; }
  const hueToRgb = (p: number, q: number, t: number) => {
    if (t < 0) t += 1; if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const hNorm = h / 360;
  return {
    r: Math.round(hueToRgb(p, q, hNorm + 1 / 3) * 255),
    g: Math.round(hueToRgb(p, q, hNorm) * 255),
    b: Math.round(hueToRgb(p, q, hNorm - 1 / 3) * 255),
  };
}

function srgbToLinear(c: number): number {
  const s = c / 255;
  return s <= 0.04045 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4;
}

function linearToSrgb(c: number): number {
  const s = c <= 0.0031308 ? c * 12.92 : 1.055 * c ** (1 / 2.4) - 0.055;
  return Math.max(0, Math.min(255, Math.round(s * 255)));
}

export function applyTint(hex: string, tint: number): string {
  const { r, g, b } = hexToRgb(hex);
  const t = tint / 100000;
  return rgbToHex(
    linearToSrgb(srgbToLinear(r) * t + 1.0 * (1 - t)),
    linearToSrgb(srgbToLinear(g) * t + 1.0 * (1 - t)),
    linearToSrgb(srgbToLinear(b) * t + 1.0 * (1 - t)),
  );
}

export function applyShade(hex: string, shade: number): string {
  const { r, g, b } = hexToRgb(hex);
  const s = shade / 100000;
  return rgbToHex(linearToSrgb(srgbToLinear(r) * s), linearToSrgb(srgbToLinear(g) * s), linearToSrgb(srgbToLinear(b) * s));
}

export interface ColorModifier { name: string; val: number; }

export function applyColorModifiers(hex: string, modifiers: ColorModifier[]): { color: string; alpha: number } {
  let color = hex;
  let alpha = 1;
  for (const mod of modifiers) {
    const name = mod.name.startsWith('a:') ? mod.name.slice(2) : mod.name;
    switch (name) {
      case 'tint': color = applyTint(color, mod.val); break;
      case 'shade': color = applyShade(color, mod.val); break;
      case 'lumMod': {
        const { r, g, b } = hexToRgb(color);
        const { h, s, l } = rgbToHsl(r, g, b);
        const rgb = hslToRgb(h, s, Math.max(0, Math.min(1, l * (mod.val / 100000))));
        color = rgbToHex(rgb.r, rgb.g, rgb.b); break;
      }
      case 'lumOff': {
        const { r, g, b } = hexToRgb(color);
        const { h, s, l } = rgbToHsl(r, g, b);
        const rgb = hslToRgb(h, s, Math.max(0, Math.min(1, l + mod.val / 100000)));
        color = rgbToHex(rgb.r, rgb.g, rgb.b); break;
      }
      case 'satMod': {
        const { r, g, b } = hexToRgb(color);
        const { h, s, l } = rgbToHsl(r, g, b);
        const rgb = hslToRgb(h, Math.max(0, Math.min(1, s * (mod.val / 100000))), l);
        color = rgbToHex(rgb.r, rgb.g, rgb.b); break;
      }
      case 'alpha': alpha = Math.max(0, Math.min(1, mod.val / 100000)); break;
      case 'alphaMod': alpha = Math.max(0, Math.min(1, alpha * (mod.val / 100000))); break;
      case 'alphaOff': alpha = Math.max(0, Math.min(1, alpha + mod.val / 100000)); break;
      default: break;
    }
  }
  return { color, alpha };
}

export function presetColorToHex(name: string): string | undefined {
  return PRESET_COLORS[name] ?? PRESET_COLORS[name.toLowerCase()];
}

const PRESET_COLORS: Record<string, string> = {
  black: '#000000', white: '#FFFFFF', red: '#FF0000', green: '#008000',
  blue: '#0000FF', yellow: '#FFFF00', cyan: '#00FFFF', magenta: '#FF00FF',
  orange: '#FFA500', purple: '#800080', brown: '#A52A2A', pink: '#FFC0CB',
  gray: '#808080', grey: '#808080', lime: '#00FF00', navy: '#000080',
  teal: '#008080', maroon: '#800000', olive: '#808000', silver: '#C0C0C0',
  aqua: '#00FFFF', fuchsia: '#FF00FF', darkBlue: '#00008B', darkCyan: '#008B8B',
  darkGray: '#A9A9A9', darkGreen: '#006400', darkRed: '#8B0000', darkViolet: '#9400D3',
  lightBlue: '#ADD8E6', lightGray: '#D3D3D3', lightGreen: '#90EE90',
};
