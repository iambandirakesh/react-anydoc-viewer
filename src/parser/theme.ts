/** Parsing of `ppt/theme/themeN.xml` and the master colour map (`p:clrMap`). */

import type { ColorMap, Fill, LineStyle, RGBA, Shadow, ThemeData } from "../model/types";
import { attr, child, children, type XmlNode } from "./xml";
import { parseFill, parseLine, parseShadow } from "./drawing";

export const DEFAULT_THEME_COLORS: Record<string, RGBA> = {
  dk1: { r: 0, g: 0, b: 0, a: 1 },
  lt1: { r: 255, g: 255, b: 255, a: 1 },
  dk2: { r: 0x44, g: 0x54, b: 0x6a, a: 1 },
  lt2: { r: 0xee, g: 0xec, b: 0xe1, a: 1 },
  accent1: { r: 0x4f, g: 0x81, b: 0xbd, a: 1 },
  accent2: { r: 0xc0, g: 0x50, b: 0x4d, a: 1 },
  accent3: { r: 0x9b, g: 0xbb, b: 0x59, a: 1 },
  accent4: { r: 0x80, g: 0x64, b: 0xa2, a: 1 },
  accent5: { r: 0x4b, g: 0xac, b: 0xc6, a: 1 },
  accent6: { r: 0xf7, g: 0x96, b: 0x46, a: 1 },
  hlink: { r: 0x00, g: 0x00, b: 0xff, a: 1 },
  folHlink: { r: 0x80, g: 0x00, b: 0x80, a: 1 },
};

/** The colour map used when a master does not declare one. */
export const DEFAULT_COLOR_MAP: ColorMap = {
  bg1: "lt1",
  tx1: "dk1",
  bg2: "lt2",
  tx2: "dk2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hlink",
  folHlink: "folHlink",
};

export function createEmptyTheme(): ThemeData {
  return {
    colors: { ...DEFAULT_THEME_COLORS },
    majorFont: null,
    minorFont: null,
    fillStyles: [],
    bgFillStyles: [],
    lineStyles: [],
    effectStyles: [],
  };
}

function schemeColor(node: XmlNode | null): RGBA | null {
  if (!node) return null;
  const srgb = child(node, "srgbClr");
  if (srgb) {
    const value = attr(srgb, "val") ?? "";
    if (/^[0-9a-fA-F]{6}$/.test(value)) {
      return {
        r: parseInt(value.slice(0, 2), 16),
        g: parseInt(value.slice(2, 4), 16),
        b: parseInt(value.slice(4, 6), 16),
        a: 1,
      };
    }
  }
  const sys = child(node, "sysClr");
  if (sys) {
    const last = attr(sys, "lastClr") ?? "";
    if (/^[0-9a-fA-F]{6}$/.test(last)) {
      return {
        r: parseInt(last.slice(0, 2), 16),
        g: parseInt(last.slice(2, 4), 16),
        b: parseInt(last.slice(4, 6), 16),
        a: 1,
      };
    }
  }
  return null;
}

const COLOR_SLOTS = [
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
] as const;

export function parseTheme(root: XmlNode | null): ThemeData {
  const theme = createEmptyTheme();
  if (!root) return theme;

  const themeElements = child(root, "themeElements");
  const clrScheme = child(themeElements ?? root, "clrScheme");
  if (clrScheme) {
    for (const slot of COLOR_SLOTS) {
      const entry = child(clrScheme, slot);
      const color = schemeColor(entry);
      if (color) theme.colors[slot] = color;
    }
  }

  const fontScheme = child(themeElements ?? root, "fontScheme");
  if (fontScheme) {
    const major = child(fontScheme, "majorFont");
    const minor = child(fontScheme, "minorFont");
    theme.majorFont = textOfLatinTypeface(major);
    theme.minorFont = textOfLatinTypeface(minor);
  }

  const fmtScheme = child(themeElements ?? root, "fmtScheme");
  if (fmtScheme) {
    const fillStyleLst = child(fmtScheme, "fillStyleLst");
    if (fillStyleLst) {
      for (const entry of fillStyleLst.children) {
        const fill = parseFill(entry, { theme });
        theme.fillStyles.push(fill ?? { type: "none" });
      }
    }
    const bgFillStyleLst = child(fmtScheme, "bgFillStyleLst");
    if (bgFillStyleLst) {
      for (const entry of bgFillStyleLst.children) {
        const fill = parseFill(entry, { theme });
        theme.bgFillStyles.push(fill ?? { type: "none" });
      }
    }
    const lnStyleLst = child(fmtScheme, "lnStyleLst");
    if (lnStyleLst) {
      for (const entry of lnStyleLst.children) {
        theme.lineStyles.push(parseLine(entry, { theme }) ?? { width: null, fill: null, dash: null, cap: null, compound: null, head: null, tail: null });
      }
    }
    const effectStyleLst = child(fmtScheme, "effectStyleLst");
    if (effectStyleLst) {
      for (const entry of effectStyleLst.children) {
        const effectLst = child(entry, "effectLst") ?? child(entry, "effectDag");
        theme.effectStyles.push(parseShadow(effectLst));
      }
    }
  }

  return theme;
}

function textOfLatinTypeface(fontGroup: XmlNode | null): string | null {
  const latin = child(fontGroup, "latin");
  const typeface = attr(latin, "typeface");
  if (!typeface) return null;
  // Office uses empty typefaces to mean "theme font"; treat those as unresolved.
  return typeface.length > 0 ? typeface : null;
}

/**
 * Parse `<p:clrMap>` from a slide master. Missing attributes fall back to the
 * default Office mapping rather than to "unknown", because a partial map is
 * common in hand-written decks.
 */
export function parseColorMap(node: XmlNode | null): ColorMap {
  if (!node) return { ...DEFAULT_COLOR_MAP };
  const map: ColorMap = { ...DEFAULT_COLOR_MAP };
  for (const slot of Object.keys(DEFAULT_COLOR_MAP)) {
    const value = attr(node, slot);
    if (value) map[slot] = value;
  }
  return map;
}

/** Resolve `<p:bgRef idx="…">` against the theme's style lists. */
export function resolveBackgroundRef(
  idx: number,
  theme: ThemeData,
  colorContext: { theme: ThemeData },
): Fill | null {
  const list: Fill[] = idx >= 1001 ? theme.bgFillStyles : theme.fillStyles;
  const index = idx >= 1001 ? idx - 1001 : idx - 1;
  const fill = list[index];
  if (!fill) return null;
  if (fill.type === "solid") return fill;
  // For gradient/image style references PowerPoint tints the style; we keep the
  // style as-is, which is visually close for the overwhelmingly common case of
  // a solid `bgFillStyleLst` entry.
  void colorContext;
  return fill;
}

export function cloneFill(fill: Fill): Fill {
  switch (fill.type) {
    case "gradient":
      return { ...fill, stops: fill.stops.map((s) => ({ ...s, color: { ...s.color } })) };
    case "solid":
      return { ...fill, color: { ...fill.color } };
    case "image":
      return { ...fill, crop: fill.crop ? { ...fill.crop } : null };
    default:
      return { ...fill };
  }
}

export function cloneShadow(shadow: Shadow | null): Shadow | null {
  return shadow ? { ...shadow, color: { ...shadow.color } } : null;
}

export function themeLineStyles(theme: ThemeData): LineStyle[] {
  return theme.lineStyles;
}

export function allThemeColors(theme: ThemeData): string[] {
  return Object.keys(theme.colors);
}
