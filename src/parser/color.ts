/**
 * DrawingML colour resolution.
 *
 * Colours in a presentation can be literal sRGB values, entries from the theme
 * palette (`schemeClr`), system colours, or named preset colours. Scheme colours
 * are resolved through the master's colour map (`p:clrMap`) because a value such
 * as `tx1` names a semantic slot, not a palette entry.
 */

import type { ColorMap, RGBA, ThemeData } from "../model/types";
import { attr, children, type XmlNode } from "./xml";
import { clamp } from "./units";

export const BLACK: RGBA = { r: 0, g: 0, b: 0, a: 1 };
export const WHITE: RGBA = { r: 255, g: 255, b: 255, a: 1 };

/** `val` is `RRGGBB` (or `#RRGGBB`); alpha is not part of the attribute. */
export function parseHexColor(value: string): RGBA | null {
  let hex = value.trim();
  if (hex.startsWith("#")) hex = hex.slice(1);
  if (hex.length === 3) {
    const r = hex[0] as string;
    const g = hex[1] as string;
    const b = hex[2] as string;
    hex = r + r + g + g + b + b;
  }
  if (hex.length !== 6 || !/^[0-9a-fA-F]{6}$/.test(hex)) return null;
  return {
    r: parseInt(hex.slice(0, 2), 16),
    g: parseInt(hex.slice(2, 4), 16),
    b: parseInt(hex.slice(4, 6), 16),
    a: 1,
  };
}

export function rgbaToCss(color: RGBA): string {
  const r = Math.round(clamp(color.r, 0, 255));
  const g = Math.round(clamp(color.g, 0, 255));
  const b = Math.round(clamp(color.b, 0, 255));
  const a = clamp(color.a, 0, 1);
  if (a >= 1) {
    return `rgb(${r}, ${g}, ${b})`;
  }
  return `rgba(${r}, ${g}, ${b}, ${Math.round(a * 1000) / 1000})`;
}

export function rgbaToHex(color: RGBA): string {
  const part = (n: number) =>
    Math.round(clamp(n, 0, 255)).toString(16).padStart(2, "0").toUpperCase();
  return `#${part(color.r)}${part(color.g)}${part(color.b)}`;
}

export function withAlpha(color: RGBA, alpha: number): RGBA {
  return { r: color.r, g: color.g, b: color.b, a: clamp(color.a * alpha, 0, 1) };
}

/** Rec. 601 luminance, the basis of DrawingML's `lumMod`/`lumOff`. */
function luminance(color: RGBA): number {
  return (0.2126 * color.r + 0.7152 * color.g + 0.0722 * color.b) / 255;
}

/**
 * Apply the luminance of `color` to `target` following DrawingML's modifier
 * semantics: `lumMod` scales luminance, `lumOff` shifts it, and both preserve
 * the hue/saturation of the target channel ratios.
 */
function applyLuminance(target: RGBA, lum: number): RGBA {
  const current = luminance(target);
  if (current <= 0.00001) {
    // Near-black: rebuild from the target luminance directly.
    return { ...target, r: lum * 255, g: lum * 255, b: lum * 255, a: target.a };
  }
  const factor = lum / current;
  return {
    r: clamp(target.r * factor, 0, 255),
    g: clamp(target.g * factor, 0, 255),
    b: clamp(target.b * factor, 0, 255),
    a: target.a,
  };
}

function rgbToHsl(color: RGBA): { h: number; s: number; l: number } {
  const r = color.r / 255;
  const g = color.g / 255;
  const b = color.b / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return { h: 0, s: 0, l };
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
  else if (max === g) h = ((b - r) / d + 2) / 6;
  else h = ((r - g) / d + 4) / 6;
  return { h, s, l };
}

function hueToChannel(p: number, q: number, t: number): number {
  let tt = t;
  if (tt < 0) tt += 1;
  if (tt > 1) tt -= 1;
  if (tt < 1 / 6) return p + (q - p) * 6 * tt;
  if (tt < 1 / 2) return q;
  if (tt < 2 / 3) return p + (q - p) * (2 / 3 - tt) * 6;
  return p;
}

function hslToRgb(h: number, s: number, l: number, a: number): RGBA {
  if (s === 0) {
    const v = l * 255;
    return { r: v, g: v, b: v, a };
  }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return {
    r: hueToChannel(p, q, h + 1 / 3) * 255,
    g: hueToChannel(p, q, h) * 255,
    b: hueToChannel(p, q, h - 1 / 3) * 255,
    a,
  };
}

/**
 * Apply the colour modifiers that can appear as children of a colour element.
 *
 * Modifiers are applied in document order, the same way PowerPoint does.
 */
export function applyColorModifiers(base: RGBA, modifiers: XmlNode[]): RGBA {
  let color: RGBA = { ...base };
  for (const modifier of modifiers) {
    const value = Number(attr(modifier, "val") ?? "");
    if (!Number.isFinite(value)) continue;
    const ratio = value / 100_000;
    switch (modifier.local) {
      case "alpha":
        color = { ...color, a: clamp(color.a * ratio, 0, 1) };
        break;
      case "alphaMod":
      case "alphaModulate": {
        const hsl = rgbToHsl(color);
        color = hslToRgb(hsl.h, hsl.s, hsl.l, clamp(color.a * ratio, 0, 1));
        break;
      }
      case "alphaOff":
        color = { ...color, a: clamp(color.a + ratio, 0, 1) };
        break;
      case "lumMod":
        color = applyLuminance(color, luminance(color) * ratio);
        break;
      case "lumOff":
        color = applyLuminance(color, luminance(color) + ratio);
        break;
      case "tint": {
        const t = clamp(ratio, 0, 1);
        color = {
          r: color.r * (1 - t) + 255 * t,
          g: color.g * (1 - t) + 255 * t,
          b: color.b * (1 - t) + 255 * t,
          a: color.a,
        };
        break;
      }
      case "shade": {
        const s = 1 - clamp(ratio, 0, 1);
        color = { r: color.r * s, g: color.g * s, b: color.b * s, a: color.a };
        break;
      }
      case "satMod": {
        const hsl = rgbToHsl(color);
        color = hslToRgb(hsl.h, clamp(hsl.s * ratio, 0, 1), hsl.l, color.a);
        break;
      }
      case "satOff": {
        const hsl = rgbToHsl(color);
        color = hslToRgb(hsl.h, clamp(hsl.s + ratio, 0, 1), hsl.l, color.a);
        break;
      }
      case "hueMod": {
        const hsl = rgbToHsl(color);
        color = hslToRgb(clamp(hsl.h * ratio, 0, 1), hsl.s, hsl.l, color.a);
        break;
      }
      case "hueOff": {
        const hsl = rgbToHsl(color);
        const h = ((hsl.h + ratio) % 1 + 1) % 1;
        color = hslToRgb(h, hsl.s, hsl.l, color.a);
        break;
      }
      case "shadeMod":
      case "gray":
      case "comp":
      case "inv":
      case "gamma":
      case "invGamma": {
        // Rarely used in presentations; approximating would be misleading, so we
        // intentionally leave the base colour untouched.
        break;
      }
      default:
        break;
    }
  }
  return color;
}

/** Colour elements that can contain modifiers directly. */
export const COLOR_ELEMENT_NAMES = new Set([
  "srgbClr",
  "schemeClr",
  "sysClr",
  "prstClr",
  "scrgbClr",
  "hslClr",
  "styleClr",
]);

export function isColorElement(node: XmlNode): boolean {
  return COLOR_ELEMENT_NAMES.has(node.local);
}

/** Named colours accepted by `a:prstClr`. A pragmatic subset of the X11 names. */
const PRESET_COLORS: Record<string, string> = {
  aliceblue: "F0F8FF",
  antiquewhite: "FAEBD7",
  aqua: "00FFFF",
  aquamarine: "7FFFD4",
  azure: "F0FFFF",
  beige: "F5F5DC",
  bisque: "FFE4C4",
  black: "000000",
  blue: "0000FF",
  blueviolet: "8A2BE2",
  brown: "A52A2A",
  burlywood: "DEB887",
  cadetblue: "5F9EA0",
  chartreuse: "7FFF00",
  chocolate: "D2691E",
  coral: "FF7F50",
  cornflowerblue: "6495ED",
  cornsilk: "FFF8DC",
  crimson: "DC143C",
  cyan: "00FFFF",
  darkblue: "00008B",
  darkcyan: "008B8B",
  darkgoldenrod: "B8860B",
  darkgray: "A9A9A9",
  darkgreen: "006400",
  darkgrey: "A9A9A9",
  darkkhaki: "BDB76B",
  darkmagenta: "8B008B",
  darkolivegreen: "556B2F",
  darkorange: "FF8C00",
  darkorchid: "9932CC",
  darkred: "8B0000",
  darksalmon: "E9967A",
  darkseagreen: "8FBC8F",
  darkslateblue: "483D8B",
  darkslategray: "2F4F4F",
  darkturquoise: "00CED1",
  darkviolet: "9400D3",
  deeppink: "FF1493",
  deepskyblue: "00BFFF",
  dimgray: "696969",
  dodgerblue: "1E90FF",
  firebrick: "B22222",
  floralwhite: "FFFAF0",
  forestgreen: "228B22",
  fuchsia: "FF00FF",
  gainsboro: "DCDCDC",
  ghostwhite: "F8F8FF",
  gold: "FFD700",
  goldenrod: "DAA520",
  gray: "808080",
  green: "008000",
  greenyellow: "ADFF2F",
  honeydew: "F0FFF0",
  hotpink: "FF69B4",
  indianred: "CD5C5C",
  indigo: "4B0082",
  ivory: "FFFFF0",
  khaki: "F0E68C",
  lavender: "E6E6FA",
  lawngreen: "7CFC00",
  lightblue: "ADD8E6",
  lightcoral: "F08080",
  lightcyan: "E0FFFF",
  lightgray: "D3D3D3",
  lightgreen: "90EE90",
  lightgrey: "D3D3D3",
  lightpink: "FFB6C1",
  lightsalmon: "FFA07A",
  lightseagreen: "20B2AA",
  lightskyblue: "87CEFA",
  lightslategray: "778899",
  lightsteelblue: "B0C4DE",
  lightyellow: "FFFFE0",
  lime: "00FF00",
  limegreen: "32CD32",
  linen: "FAF0E6",
  magenta: "FF00FF",
  maroon: "800000",
  mediumaquamarine: "66CDAA",
  mediumblue: "0000CD",
  mediumorchid: "BA55D3",
  mediumpurple: "9370DB",
  mediumseagreen: "3CB371",
  mediumslateblue: "7B68EE",
  mediumspringgreen: "00FA9A",
  mediumturquoise: "48D1CC",
  mediumvioletred: "C71585",
  midnightblue: "191970",
  mintcream: "F5FFFA",
  mistyrose: "FFE4E1",
  moccasin: "FFE4B5",
  navajowhite: "FFDEAD",
  navy: "000080",
  oldlace: "FDF5E6",
  olive: "808000",
  olivedrab: "6B8E23",
  orange: "FFA500",
  orangered: "FF4500",
  orchid: "DA70D6",
  palegoldenrod: "EEE8AA",
  palegreen: "98FB98",
  paleturquoise: "AFEEEE",
  palevioletred: "DB7093",
  papayawhip: "FFEFD5",
  peachpuff: "FFDAB9",
  peru: "CD853F",
  pink: "FFC0CB",
  plum: "DDA0DD",
  powderblue: "B0E0E6",
  purple: "800080",
  red: "FF0000",
  rosybrown: "BC8F8F",
  royalblue: "4169E1",
  saddlebrown: "8B4513",
  salmon: "FA8072",
  sandybrown: "F4A460",
  seagreen: "2E8B57",
  seashell: "FFF5EE",
  sienna: "A0522D",
  silver: "C0C0C0",
  skyblue: "87CEEB",
  slateblue: "6A5ACD",
  slategray: "708090",
  snow: "FFFAFA",
  springgreen: "00FF7F",
  steelblue: "4682B4",
  tan: "D2B48C",
  teal: "008080",
  thistle: "D8BFD8",
  tomato: "FF6347",
  turquoise: "40E0D0",
  violet: "EE82EE",
  wheat: "F5DEB3",
  white: "FFFFFF",
  whitesmoke: "F5F5F5",
  yellow: "FFFF00",
  yellowgreen: "9ACD32",
};

const SYSTEM_COLORS: Record<string, string> = {
  windowtext: "000000",
  window: "FFFFFF",
  buttonface: "F0F0F0",
  buttontext: "000000",
  highlight: "3399FF",
  highlighttext: "FFFFFF",
  graytext: "808080",
  captiontext: "000000",
  inactivecaptiontext: "808080",
  activeborder: "FFFFFF",
  inactiveborder: "FFFFFF",
  menu: "F0F0F0",
  menutext: "000000",
  scrollbar: "D4D0C8",
  threeddarkshadow: "000000",
  threedface: "F0F0F0",
  threedhighlight: "FFFFFF",
  threedlightshadow: "D4D0C8",
  threedshadow: "808080",
};

export interface ColorContext {
  theme: ThemeData;
  /** absent means "use the default Office colour map" */
  colorMap?: ColorMap;
  /** base colour used for `schemeClr` values not present in the theme */
  fallback?: RGBA;
}

/** Resolve a single `<a:srgbClr>`-style element to RGBA. */
export function resolveColorElement(node: XmlNode, ctx: ColorContext): RGBA {
  const value = attr(node, "val") ?? "";
  const modifiers = children(node, "").length ? node.children : node.children;
  let base: RGBA = ctx.fallback ?? BLACK;

  switch (node.local) {
    case "srgbClr": {
      base = parseHexColor(value) ?? ctx.fallback ?? BLACK;
      break;
    }
    case "scrgbClr": {
      const toByte = (name: string) => {
        const raw = Number(attr(node, name) ?? "0");
        return Number.isFinite(raw) ? (raw / 100_000) * 255 : 0;
      };
      base = { r: toByte("r"), g: toByte("g"), b: toByte("b"), a: 1 };
      break;
    }
    case "schemeClr": {
      const mapped = (ctx.colorMap ?? {})[value] ?? value;
      const fromMap = ctx.theme.colors[mapped];
      const direct = ctx.theme.colors[value];
      base = fromMap ?? direct ?? ctx.fallback ?? BLACK;
      break;
    }
    case "sysClr": {
      const last = attr(node, "lastClr");
      const parsed = last ? parseHexColor(last) : null;
      base = parsed ?? parseHexColor(SYSTEM_COLORS[value.toLowerCase()] ?? "") ?? ctx.fallback ?? BLACK;
      break;
    }
    case "prstClr": {
      const hex = PRESET_COLORS[value.toLowerCase()];
      base = (hex ? parseHexColor(hex) : null) ?? ctx.fallback ?? BLACK;
      break;
    }
    case "hslClr": {
      const num = (name: string, fallback: number) => {
        const raw = Number(attr(node, name) ?? "");
        return Number.isFinite(raw) ? raw : fallback;
      };
      const hue = num("hue", 0) / 60_000;
      const sat = num("sat", 100_000) / 100_000;
      const lum = num("lum", 50_000) / 100_000;
      base = hslToRgb(((hue % 360) + 360) % 360 / 360, clamp(sat, 0, 1), clamp(lum, 0, 1), 1);
      break;
    }
    default:
      base = ctx.fallback ?? BLACK;
      break;
  }

  // Modifiers are the element's own children; re-reading them directly is
  // cheaper than filtering and keeps ordering intact.
  return applyColorModifiers(base, modifiers);
}

export function themeColor(theme: ThemeData, name: string): RGBA | null {
  return theme.colors[name] ?? null;
}
