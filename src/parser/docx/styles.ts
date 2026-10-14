/**
 * WordprocessingML styles, numbering and default formatting.
 *
 * Word's cascade is long; the parts that change what a document looks like are
 * implemented here:
 *
 *   docDefaults -> paragraph style chain (`w:basedOn`) -> numbering level
 *     -> direct `w:pPr` -> paragraph-mark `w:rPr` -> run's own `w:rPr`
 *       -> character style (`w:rStyle`)
 *
 * Documents that carry no styles at all (hand-written or generated files) still
 * get Word's own defaults, including the built-in heading ladder, which is what
 * makes an unstyled document legible instead of a wall of 11pt body text.
 */

import type { Bullet, ParaProps, RGBA, RunProps } from "../../model/types";
import { child, children, type XmlNode } from "../xml";
import { clamp, halfPtToPx, ptToPx, twipsToPx } from "../units";
import { wAttr, wNum, wToggle, wVal } from "./xmlw";

/* ------------------------------------------------------------------ colours */

/** `w:highlight` / `w:shd` named colours, as HTML-style values. */
const NAMED_HIGHLIGHTS: Record<string, string> = {
  black: "000000",
  blue: "0000FF",
  cyan: "00FFFF",
  darkBlue: "000080",
  darkCyan: "008080",
  darkGray: "808080",
  darkGreen: "008000",
  darkMagenta: "800080",
  darkRed: "800000",
  darkYellow: "808000",
  green: "00FF00",
  lightGray: "C0C0C0",
  magenta: "FF00FF",
  red: "FF0000",
  white: "FFFFFF",
  yellow: "FFFF00",
};

/** The default Office theme accents, used when `w:themeColor` is set. */
const THEME_COLORS: Record<string, string> = {
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "44546A",
  lt2: "E7E6E6",
  accent1: "4472C4",
  accent2: "ED7D31",
  accent3: "A5A5A5",
  accent4: "FFC000",
  accent5: "5B9BD5",
  accent6: "70AD47",
  hlink: "0563C1",
  folHlink: "954F72",
  background1: "FFFFFF",
  text1: "000000",
  background2: "E7E6E6",
  text2: "44546A",
};

export function parseHexColor(value: string | null): RGBA | undefined {
  if (!value) return undefined;
  const hex = value.trim().replace(/^#/, "");
  if (!/^[0-9a-fA-F]{6}$/.test(hex)) return undefined;
  return {
    r: parseInt(hex.slice(0, 2), 16),
    g: parseInt(hex.slice(2, 4), 16),
    b: parseInt(hex.slice(4, 6), 16),
    a: 1,
  };
}

function colorFromNode(node: XmlNode | null): RGBA | undefined {
  if (!node) return undefined;
  const raw = wVal(node);
  if (raw && raw.toLowerCase() !== "auto") {
    const parsed = parseHexColor(raw);
    if (parsed) return parsed;
  }
  const theme = wAttr(node, "themeColor");
  if (theme) return parseHexColor(THEME_COLORS[theme] ?? null);
  return undefined;
}

/* ------------------------------------------------------------------- runs */

/** Word's default body run: 11pt Calibri, black. */
export const DEFAULT_RUN_PROPS: RunProps = {
  size: halfPtToPx(22),
  font: "Calibri",
  color: { r: 0, g: 0, b: 0, a: 1 },
};

export const DEFAULT_PARA_PROPS: ParaProps = {};

/** Map `w:u` values onto the model's underline tokens. */
function underlineFrom(value: string | null): string | false | undefined {
  if (value === null) return undefined;
  const normalized = value.trim().toLowerCase();
  if (normalized === "none") return false;
  if (normalized === "double" || normalized === "thick" || normalized === "wave") {
    return normalized === "double" ? "dbl" : "sng";
  }
  return "sng";
}

/** Read `w:rPr` (run properties) into the model's run shape. */
export function parseRunProperties(node: XmlNode | null, base?: RunProps): RunProps {
  const props: RunProps = { ...(base ?? {}) };
  if (!node) return props;
  for (const item of node.children) {
    switch (item.local) {
      case "rFonts": {
        const font = wAttr(item, "ascii") ?? wAttr(item, "hAnsi") ?? wAttr(item, "cs");
        if (font) props.font = font;
        break;
      }
      case "b": {
        const value = wToggle(item);
        if (value !== null) props.bold = value;
        break;
      }
      case "i": {
        const value = wToggle(item);
        if (value !== null) props.italic = value;
        break;
      }
      case "caps": {
        const value = wToggle(item);
        if (value !== null) props.caps = value ? "all" : "none";
        break;
      }
      case "smallCaps": {
        const value = wToggle(item);
        if (value !== null) props.caps = value ? "small" : "none";
        break;
      }
      case "strike":
      case "dstrike": {
        const value = wToggle(item);
        if (value === true) props.strike = "sngStrike";
        else if (value === false) props.strike = false;
        break;
      }
      case "u": {
        const underline = underlineFrom(wVal(item) ?? "single");
        if (underline !== undefined) props.underline = underline;
        break;
      }
      case "color": {
        const color = colorFromNode(item);
        if (color) props.color = color;
        break;
      }
      case "sz": {
        const size = wNum(item, "val");
        if (size !== null && size > 0) props.size = clamp(halfPtToPx(size), 4, 400);
        break;
      }
      case "spacing": {
        // Character spacing is measured in twips in WordprocessingML.
        const spacing = wNum(item, "val");
        if (spacing !== null) props.spacing = twipsToPx(spacing);
        break;
      }
      case "highlight": {
        const named = NAMED_HIGHLIGHTS[(wVal(item) ?? "").trim()] ?? null;
        const color = parseHexColor(named);
        if (color) props.highlight = color;
        break;
      }
      case "shd": {
        const fill = wAttr(item, "fill");
        if (fill && fill.toLowerCase() !== "auto") {
          const color = parseHexColor(fill);
          if (color) props.highlight = color;
        }
        break;
      }
      case "vertAlign": {
        const value = wVal(item);
        if (value === "superscript") props.baseline = 0.3;
        else if (value === "subscript") props.baseline = -0.25;
        break;
      }
      default:
        break;
    }
  }
  return props;
}

/* ------------------------------------------------------------- paragraphs */

export type WordAlign = "l" | "ctr" | "r" | "just" | "dist";

function alignFrom(value: string | null): WordAlign | undefined {
  switch (value) {
    case "center":
      return "ctr";
    case "right":
    case "end":
      return "r";
    case "both":
      return "just";
    case "distribute":
      return "dist";
    case "left":
    case "start":
      return "l";
    default:
      return undefined;
  }
}

export interface WordParaStyle {
  props: ParaProps;
  pageBreakBefore: boolean;
  keepNext: boolean;
  numId: number | null;
  ilvl: number;
  hasBorders: boolean;
  hasShading: boolean;
  /** true when the paragraph itself set `w:ind` (it then beats the list indent) */
  directIndent: boolean;
  /** run properties inherited by the paragraph's runs (`w:pPr/w:rPr`) */
  runProps: RunProps;
  styleId: string | null;
}

/** Read `w:pPr` into paragraph properties plus the flags pagination needs. */
export function parseParagraphProperties(node: XmlNode | null, base?: ParaProps): WordParaStyle {
  const props: ParaProps = { ...(base ?? {}) };
  const result: WordParaStyle = {
    props,
    pageBreakBefore: false,
    keepNext: false,
    numId: null,
    ilvl: 0,
    hasBorders: false,
    hasShading: false,
    directIndent: false,
    runProps: {},
    styleId: null,
  };
  if (!node) return result;

  for (const item of node.children) {
    switch (item.local) {
      case "pStyle":
        result.styleId = wVal(item);
        break;
      case "jc": {
        const align = alignFrom(wVal(item));
        if (align) props.align = align;
        break;
      }
      case "ind": {
        const left = wNum(item, "left") ?? wNum(item, "start");
        const right = wNum(item, "right") ?? wNum(item, "end");
        const hanging = wNum(item, "hanging");
        const firstLine = wNum(item, "firstLine");
        if (left !== null || right !== null || hanging !== null || firstLine !== null) {
          result.directIndent = true;
          props.indent = {
            marL: twipsToPx(left ?? 0),
            marR: twipsToPx(right ?? 0),
            indent: hanging !== null ? -twipsToPx(hanging) : twipsToPx(firstLine ?? 0),
          };
        }
        break;
      }
      case "spacing": {
        const before = wNum(item, "before");
        const after = wNum(item, "after");
        if (before !== null) props.spaceBefore = twipsToPx(before);
        if (after !== null) props.spaceAfter = twipsToPx(after);
        const line = wNum(item, "line");
        if (line !== null && line > 0) {
          const rule = (wAttr(item, "lineRule") ?? "auto").toLowerCase();
          if (rule === "auto") props.lineSpacing = { type: "pct", value: clamp(line / 240, 0.5, 4) };
          else props.lineSpacing = { type: "pt", value: twipsToPx(line) };
        }
        break;
      }
      case "numPr": {
        const numId = wNum(child(item, "numId"), "val");
        if (numId !== null) result.numId = numId;
        const ilvl = wNum(child(item, "ilvl"), "val");
        if (ilvl !== null) result.ilvl = clamp(Math.round(ilvl), 0, 8);
        break;
      }
      case "pageBreakBefore": {
        const value = wToggle(item);
        if (value !== null) result.pageBreakBefore = value;
        break;
      }
      case "keepNext": {
        const value = wToggle(item);
        if (value !== null) result.keepNext = value;
        break;
      }
      case "pBdr":
        result.hasBorders = children(item, "top").length > 0 || item.children.length > 0;
        break;
      case "shd": {
        const fill = wAttr(item, "fill");
        if (fill && fill.toLowerCase() !== "auto") result.hasShading = true;
        break;
      }
      case "rPr":
        result.runProps = parseRunProperties(item);
        break;
      default:
        break;
    }
  }
  result.props = props;
  return result;
}

/* ---------------------------------------------------------------- numbering */

export interface WordNumberingLevel {
  bullet: Bullet | null;
  indent: { marL: number; marR: number; indent: number } | null;
}

interface RawLevel {
  start: number;
  numFmt: string;
  lvlText: string;
  runProps: RunProps;
  indent: { marL: number; marR: number; indent: number } | null;
}

interface RawAbstractNumbering {
  levels: Map<number, RawLevel>;
}

const SYMBOL_BULLETS: Record<string, string> = {
  "\uF0B7": "•",
  "\uF0A7": "▪",
  "\uF075": "◆",
  "\uF0D8": "➢",
  "\uF0FC": "✓",
  "\uF06C": "•",
  "\uF0A8": "▪",
  "o": "○",
};

/** Private-use-area bullet glyphs come from Symbol/Wingdings; map the common ones. */
export function bulletGlyph(raw: string): string {
  const trimmed = raw.trim();
  if (trimmed.length === 0) return "•";
  const mapped = SYMBOL_BULLETS[trimmed];
  if (mapped) return mapped;
  const code = trimmed.charCodeAt(0);
  if (code >= 0xe000 && code <= 0xf8ff) return "•";
  return trimmed;
}

const NUMBER_FORMATS: Record<string, string> = {
  decimal: "arabic",
  decimalZero: "arabic",
  lowerLetter: "alphaLc",
  upperLetter: "alphaUc",
  lowerRoman: "romanLc",
  upperRoman: "romanUc",
  ordinal: "arabic",
  cardinalText: "arabic",
  ordinalText: "arabic",
};

function autoNumberScheme(numFmt: string, lvlText: string): string {
  const base = NUMBER_FORMATS[numFmt] ?? "arabic";
  const suffix = lvlText.trim().endsWith(")")
    ? "ParenR"
    : lvlText.trim().endsWith(".")
      ? "Period"
      : "Plain";
  return `${base}${suffix}`;
}

/* ------------------------------------------------------------------ styles */

/** Built-in formatting for the styles Word always defines. */
function builtInParagraphStyle(name: string): { run: RunProps; para: ParaProps } | null {
  const heading = /^heading\s*([1-6])$/i.exec(name);
  if (heading) {
    const level = Number(heading[1]);
    const sizes = [16, 13, 12, 11, 11, 11];
    const colors = ["2F5496", "2F5496", "1F3763", "2F5496", "2F5496", "2F5496"];
    const before = [12, 10, 8, 6, 6, 6];
    const after = [4, 4, 4, 2, 2, 2];
    return {
      run: {
        size: ptToPx(sizes[level - 1] ?? 11),
        bold: level <= 4,
        italic: level >= 5,
        color: parseHexColor(colors[level - 1] ?? "2F5496"),
        font: "Calibri Light",
      },
      para: { spaceBefore: ptToPx(before[level - 1] ?? 6), spaceAfter: ptToPx(after[level - 1] ?? 2) },
    };
  }
  const lower = name.toLowerCase();
  if (lower === "title") {
    return {
      run: { size: ptToPx(28), color: parseHexColor("2F5496"), font: "Calibri Light" },
      para: { spaceAfter: ptToPx(8) },
    };
  }
  if (lower === "subtitle") {
    return {
      run: { size: ptToPx(15), color: parseHexColor("5A5A5A") },
      para: { spaceAfter: ptToPx(8) },
    };
  }
  if (lower.startsWith("quote")) {
    return { run: { italic: true, color: parseHexColor("404040") }, para: { indent: { marL: 48, marR: 0, indent: 0 } } };
  }
  return null;
}

export interface ResolvedParagraphStyle {
  run: RunProps;
  para: ParaProps;
  /** numbering declared by the style itself (`w:numPr` inside `w:pPr`) */
  numId: number | null;
  ilvl: number;
}

export interface DocxStyles {
  /** run properties from `w:docDefaults` */
  defaultRun: RunProps;
  /** paragraph properties from `w:docDefaults` */
  defaultPara: ParaProps;
  /** resolve a paragraph style id (plus its `w:basedOn` chain, plus built-ins) */
  paragraphStyle(styleId: string | null): ResolvedParagraphStyle | null;
  /** resolve a character style id */
  characterStyle(styleId: string | null): RunProps | null;
  /** resolve `w:numPr` to a bullet and indent */
  numbering(numId: number | null, ilvl: number): WordNumberingLevel | null;
  /** style name for a paragraph style id, when known */
  styleName(styleId: string | null): string | null;
  /** true when the style id exists in the file or is one of Word's built-ins */
  isKnownStyle(styleId: string | null): boolean;
}

export function buildStyles(stylesRoot: XmlNode | null, numberingRoot: XmlNode | null): DocxStyles {
  let defaultRun: RunProps = { ...DEFAULT_RUN_PROPS };
  let defaultPara: ParaProps = {};
  interface StyleNode {
    id: string;
    type: string;
    name: string | null;
    basedOn: string | null;
    run: RunProps;
    paraProps: ParaProps;
    numId: number | null;
    ilvl: number;
  }
  const styleNodes = new Map<string, StyleNode>();

  if (stylesRoot) {
    const docDefaults = child(stylesRoot, "docDefaults");
    const runDefaults = child(child(docDefaults, "rPrDefault"), "rPr");
    if (runDefaults) defaultRun = { ...DEFAULT_RUN_PROPS, ...parseRunProperties(runDefaults) };
    const paraDefaults = child(child(docDefaults, "pPrDefault"), "pPr");
    if (paraDefaults) defaultPara = parseParagraphProperties(paraDefaults).props;

    for (const style of children(stylesRoot, "style")) {
      const id = wAttr(style, "styleId");
      if (!id) continue;
      const pPr = child(style, "pPr");
      const flags = pPr ? parseParagraphProperties(pPr) : null;
      styleNodes.set(id, {
        id,
        type: wAttr(style, "type") ?? "paragraph",
        name: wVal(child(style, "name")),
        basedOn: wVal(child(style, "basedOn")),
        run: parseRunProperties(child(style, "rPr")),
        paraProps: flags?.props ?? {},
        numId: flags?.numId ?? null,
        ilvl: flags?.ilvl ?? 0,
      });
    }
  }

  const resolvedParagraph = new Map<string, ResolvedParagraphStyle>();
  const resolvedCharacter = new Map<string, RunProps>();

  function builtInFor(styleId: string): { run: RunProps; para: ParaProps } | null {
    const style = styleNodes.get(styleId);
    const name = style?.name ?? styleId.replace(/([a-z])([0-9])/g, "$1 $2");
    return builtInParagraphStyle(name);
  }

  function resolveParagraphStyle(styleId: string, seen: Set<string>): ResolvedParagraphStyle | null {
    const key = styleId.toLowerCase();
    const cached = resolvedParagraph.get(key);
    if (cached) return cached;
    if (seen.has(key)) return null; // `basedOn` cycle
    seen.add(key);

    const style = styleNodes.get(styleId);
    const builtIn = builtInFor(styleId);
    let run: RunProps = { ...(builtIn?.run ?? {}) };
    let para: ParaProps = { ...(builtIn?.para ?? {}) };
    let numId: number | null = null;
    let ilvl = 0;

    if (style?.basedOn) {
      const parent = resolveParagraphStyle(style.basedOn, seen);
      if (parent) {
        run = { ...run, ...parent.run };
        para = { ...para, ...parent.para };
        numId = parent.numId;
        ilvl = parent.ilvl;
      }
    }
    // The style's own properties win over the built-in defaults it inherits.
    if (style) {
      run = { ...run, ...style.run };
      para = { ...para, ...style.paraProps };
      if (style.numId !== null) numId = style.numId;
      if (style.numId !== null) ilvl = style.ilvl;
    }
    const merged: ResolvedParagraphStyle = { run, para, numId, ilvl };
    resolvedParagraph.set(key, merged);
    return merged;
  }

  function resolveCharacterStyle(styleId: string, seen: Set<string>): RunProps | null {
    const key = styleId.toLowerCase();
    const cached = resolvedCharacter.get(key);
    if (cached) return cached;
    if (seen.has(key)) return null;
    seen.add(key);
    const style = styleNodes.get(styleId);
    if (!style) return null;
    const parent = style.basedOn ? resolveCharacterStyle(style.basedOn, seen) : null;
    const merged: RunProps = { ...(parent ?? {}), ...style.run };
    resolvedCharacter.set(key, merged);
    return merged;
  }

  /* ---------------------------------------------------------- numbering */

  const abstractLevels = new Map<string, RawAbstractNumbering>();
  const numToAbstract = new Map<string, string>();

  if (numberingRoot) {
    for (const abstract of children(numberingRoot, "abstractNum")) {
      const id = wAttr(abstract, "abstractNumId");
      if (!id) continue;
      const levels = new Map<number, RawLevel>();
      for (const lvl of children(abstract, "lvl")) {
        const ilvl = wNum(lvl, "ilvl") ?? 0;
        const pPr = child(lvl, "pPr");
        const ind = pPr ? child(pPr, "ind") : null;
        const left = ind ? (wNum(ind, "left") ?? wNum(ind, "start")) : null;
        const hanging = ind ? wNum(ind, "hanging") : null;
        const firstLine = ind ? wNum(ind, "firstLine") : null;
        levels.set(clamp(Math.round(ilvl), 0, 8), {
          start: wNum(child(lvl, "start"), "val") ?? 1,
          numFmt: (wVal(child(lvl, "numFmt")) ?? "decimal").toLowerCase(),
          lvlText: wVal(child(lvl, "lvlText")) ?? "",
          runProps: parseRunProperties(child(lvl, "rPr")),
          indent:
            left !== null || hanging !== null || firstLine !== null
              ? {
                  marL: twipsToPx(left ?? (ilvl + 1) * 720),
                  marR: 0,
                  indent: hanging !== null ? -twipsToPx(hanging) : twipsToPx(firstLine ?? 0),
                }
              : null,
        });
      }
      abstractLevels.set(id, { levels });
    }
    for (const num of children(numberingRoot, "num")) {
      const id = wAttr(num, "numId");
      const abstractId = wVal(child(num, "abstractNumId"));
      if (id && abstractId) numToAbstract.set(id, abstractId);
    }
  }

  return {
    defaultRun,
    defaultPara,

    paragraphStyle(styleId) {
      if (!styleId) return null;
      return resolveParagraphStyle(styleId, new Set());
    },

    isKnownStyle(styleId) {
      if (!styleId) return false;
      return styleNodes.has(styleId) || builtInParagraphStyle(styleId) !== null;
    },

    characterStyle(styleId) {
      if (!styleId) return null;
      return resolveCharacterStyle(styleId, new Set());
    },

    styleName(styleId) {
      if (!styleId) return null;
      return styleNodes.get(styleId)?.name ?? null;
    },

    numbering(numId, ilvl) {
      let level: RawLevel | undefined;
      if (numId !== null) {
        const abstractId = numToAbstract.get(String(numId));
        const abstract = abstractId ? abstractLevels.get(abstractId) : null;
        level = abstract?.levels.get(ilvl);
      }
      if (!level) {
        // Fall back to any abstract numbering so a broken reference still indents.
        for (const abstract of abstractLevels.values()) {
          const candidate = abstract.levels.get(ilvl);
          if (candidate) {
            level = candidate;
            break;
          }
        }
      }
      if (!level) return null;

      const defaultIndent = { marL: twipsToPx((ilvl + 1) * 720), marR: 0, indent: -twipsToPx(360) };
      if (level.numFmt === "none") {
        return { bullet: { type: "none" }, indent: level.indent ?? defaultIndent };
      }
      if (level.numFmt === "bullet") {
        return {
          bullet: {
            type: "char",
            text: {
              char: bulletGlyph(level.lvlText),
              color: level.runProps.color,
              font: level.runProps.font,
            },
          },
          indent: level.indent ?? defaultIndent,
        };
      }
      return {
        bullet: {
          type: "autoNum",
          scheme: autoNumberScheme(level.numFmt, level.lvlText),
          startAt: level.start,
          text: { char: "", color: level.runProps.color, font: level.runProps.font },
        },
        indent: level.indent ?? defaultIndent,
      };
    },
  };
}
