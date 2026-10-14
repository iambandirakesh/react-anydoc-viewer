/**
 * Parsing of the DrawingML property constructs shared by every shape type:
 * fills, outlines, effects, transforms and geometry.
 */

import type {
  ColorMap,
  Crop,
  Fill,
  Geometry,
  LineEnd,
  LineStyle,
  PathCommand,
  RGBA,
  Shadow,
  ThemeData,
  Transform,
} from "../model/types";
import { attr, attrBool, attrNum, child, children, relAttr, type XmlNode } from "./xml";
import { angleUnitsToDegrees, clamp, percentageToRatio } from "./units";
import { isColorElement, resolveColorElement, themeColor } from "./color";
import { isSupportedPreset } from "./geometry";
import { DEFAULT_COLOR_MAP } from "./theme";

export interface RelTarget {
  /** archive path of the target part, or the raw URL for external targets */
  target: string;
  external: boolean;
}

export interface DrawingContext {
  theme: ThemeData;
  colorMap?: ColorMap;
  /** resolves an `r:embed` / `r:link` id against the *current part's* rels */
  resolveRel?: (id: string) => RelTarget | null;
  /** reports skipped constructs without aborting the parse */
  warn?: (code: string, message: string) => void;
}

const DEFAULT_LINE_WIDTH = 9_525; // 0.75pt in EMU

function ctxColorMap(ctx: DrawingContext): ColorMap {
  return ctx.colorMap ?? DEFAULT_COLOR_MAP;
}

/* ------------------------------------------------------------- transforms */

export function parseTransform(node: XmlNode | null): Transform | null {
  const xfrm = child(node, "xfrm") ?? (node?.local === "xfrm" ? node : null);
  if (!xfrm) return null;
  const off = child(xfrm, "off");
  const ext = child(xfrm, "ext");
  return {
    x: attrNum(off, "x") ?? 0,
    y: attrNum(off, "y") ?? 0,
    cx: attrNum(ext, "cx") ?? 0,
    cy: attrNum(ext, "cy") ?? 0,
    rot: angleUnitsToDegrees(attrNum(xfrm, "rot")),
    flipH: attrBool(xfrm, "flipH") === true,
    flipV: attrBool(xfrm, "flipV") === true,
  };
}

export function parseGroupChildTransform(node: XmlNode | null): {
  off: { x: number; y: number };
  ext: { cx: number; cy: number };
} | null {
  const xfrm = child(node, "xfrm");
  if (!xfrm) return null;
  const chOff = child(xfrm, "chOff");
  const chExt = child(xfrm, "chExt");
  if (!chOff && !chExt) return null;
  return {
    off: { x: attrNum(chOff, "x") ?? 0, y: attrNum(chOff, "y") ?? 0 },
    ext: { cx: attrNum(chExt, "cx") ?? 0, cy: attrNum(chExt, "cy") ?? 0 },
  };
}

/* ------------------------------------------------------------------ fills */

function parseCrop(srcRect: XmlNode | null): Crop | null {
  if (!srcRect) return null;
  const read = (name: string) => {
    const raw = attrNum(srcRect, name);
    return raw === null ? 0 : clamp(percentageToRatio(raw), -1, 1);
  };
  const crop: Crop = { l: read("l"), t: read("t"), r: read("r"), b: read("b") };
  if (crop.l === 0 && crop.t === 0 && crop.r === 0 && crop.b === 0) return null;
  return crop;
}

function readColorChild(node: XmlNode, ctx: DrawingContext, fallback?: RGBA): RGBA | null {
  for (const candidate of node.children) {
    if (isColorElement(candidate)) {
      return resolveColorElement(candidate, {
        theme: ctx.theme,
        colorMap: ctxColorMap(ctx),
        fallback,
      });
    }
  }
  return null;
}

export function parseFill(node: XmlNode | null, ctx: DrawingContext): Fill | null {
  if (!node) return null;
  switch (node.local) {
    case "noFill":
      return { type: "none" };
    case "solidFill": {
      const color = readColorChild(node, ctx);
      return color ? { type: "solid", color } : { type: "none" };
    }
    case "gradFill": {
      const stops: { pos: number; color: RGBA }[] = [];
      const gsLst = child(node, "gsLst");
      for (const gs of children(gsLst, "gs")) {
        const pos = clamp(percentageToRatio(attrNum(gs, "pos")), 0, 1);
        const color = readColorChild(gs, ctx);
        if (color) stops.push({ pos, color });
      }
      if (stops.length === 0) return { type: "none" };
      if (stops.length === 1) return { type: "solid", color: stops[0]!.color };
      stops.sort((a, b) => a.pos - b.pos);
      const lin = child(node, "lin");
      const radial = child(node, "path") !== null;
      return {
        type: "gradient",
        stops,
        angle: (angleUnitsToDegrees(attrNum(lin, "ang")) + 90) % 360,
        scaled: attrBool(lin, "scaled") === true,
        radial,
      };
    }
    case "blipFill": {
      const blip = child(node, "blip");
      if (!blip) return { type: "none" };
      const embed = relAttr(blip, "embed") ?? relAttr(blip, "link");
      const resolved = embed && ctx.resolveRel ? ctx.resolveRel(embed) : null;
      if (!resolved) {
        ctx.warn?.("image.missing-relationship", "Picture fill references an unknown relationship");
        return { type: "none" };
      }
      return {
        type: "image",
        mediaPath: resolved.target,
        mode: child(node, "tile") ? "tile" : "stretch",
        crop: parseCrop(child(node, "srcRect")),
      };
    }
    case "pattFill": {
      // Patterns are rendered as their foreground colour: a solid block is much
      // closer to the original than dropping the fill entirely.
      const fg = child(node, "fgClr");
      const color = fg ? readColorChild(fg, ctx) : null;
      return color ? { type: "solid", color } : { type: "none" };
    }
    case "grpFill":
      ctx.warn?.("fill.group", "Group fill is not supported; falling back to no fill");
      return { type: "none" };
    case "solidFillLegacy":
      return { type: "none" };
    default:
      return null;
  }
}

/** First fill-bearing child of a shape property element. */
export function parseFillFromProps(
  props: XmlNode | null,
  ctx: DrawingContext,
  options: { lastOfType?: number } = {},
): Fill | null {
  if (!props) return null;
  const candidates = props.children.filter(
    (c) =>
      c.local === "solidFill" ||
      c.local === "noFill" ||
      c.local === "gradFill" ||
      c.local === "blipFill" ||
      c.local === "pattFill" ||
      c.local === "grpFill",
  );
  if (options.lastOfType !== undefined) {
    const target = candidates[options.lastOfType];
    return target ? parseFill(target, ctx) : null;
  }
  const first = candidates[0];
  return first ? parseFill(first, ctx) : null;
}

/* ------------------------------------------------------------------ lines */

function parseLineEnd(node: XmlNode | null): LineEnd | null {
  if (!node) return null;
  const type = attr(node, "type");
  if (!type || type === "none") return null;
  return {
    type,
    width: attr(node, "w") ?? undefined,
    length: attr(node, "len") ?? undefined,
  };
}

export function parseLine(node: XmlNode | null, ctx: DrawingContext): LineStyle | null {
  const ln = node?.local === "ln" ? node : child(node, "ln");
  if (!ln) return null;
  const width = attrNum(ln, "w");
  const fillChildren = ln.children.filter(
    (c) => c.local !== "prstDash" && c.local !== "headEnd" && c.local !== "tailEnd" && c.local !== "custDash",
  );
  let fill: Fill | null = null;
  for (const candidate of fillChildren) {
    const parsed = parseFill(candidate, ctx);
    if (parsed) {
      fill = parsed;
      break;
    }
  }
  const dash = attr(child(ln, "prstDash"), "val") ?? attr(child(ln, "custDash"), "val");
  return {
    width: width === null ? DEFAULT_LINE_WIDTH : width,
    fill,
    dash: dash ?? null,
    cap: attr(ln, "cap"),
    compound: attr(ln, "cmpd"),
    head: parseLineEnd(child(ln, "headEnd")),
    tail: parseLineEnd(child(ln, "tailEnd")),
  };
}

/* ---------------------------------------------------------------- effects */

/**
 * Effects we cannot express in CSS.
 *
 * They are reported rather than dropped so a slide that looks flatter than it
 * does in PowerPoint is explicable from the diagnostics panel.
 */
const UNAPPROXIMATED_EFFECTS = new Set([
  "glow",
  "softEdge",
  "reflection",
  "prstShdw",
  "fillOverlay",
  "effect3d",
  "blur",
]);

export function parseShadow(effectContainer: XmlNode | null, ctx?: DrawingContext): Shadow | null {
  const effectLst = effectContainer?.local === "effectLst" ? effectContainer : child(effectContainer, "effectLst");
  if (!effectLst) return null;

  for (const effect of effectLst.children) {
    if (UNAPPROXIMATED_EFFECTS.has(effect.local)) {
      ctx?.warn?.(
        "effect.unsupported",
        `Effect "${effect.local}" is not supported and was skipped.`,
      );
    }
  }

  const outer = child(effectLst, "outerShdw");
  const inner = child(effectLst, "innerShdw");
  const node = outer ?? inner;
  if (!node) return null;
  if (inner && !outer) {
    ctx?.warn?.(
      "effect.inner-shadow-approximated",
      "Inner shadows are drawn as an outer drop shadow.",
    );
  }
  const srgb = child(node, "srgbClr");
  const value = attr(srgb, "val") ?? "000000";
  const color: RGBA = /^[0-9a-fA-F]{6}$/.test(value)
    ? {
        r: parseInt(value.slice(0, 2), 16),
        g: parseInt(value.slice(2, 4), 16),
        b: parseInt(value.slice(4, 6), 16),
        a: 1,
      }
    : { r: 0, g: 0, b: 0, a: 1 };
  if (srgb) {
    const alpha = child(srgb, "alpha");
    const alphaValue = attrNum(alpha, "val");
    if (alphaValue !== null) color.a = clamp(percentageToRatio(alphaValue), 0, 1);
  } else {
    color.a = 0.5;
  }
  const scale = (name: string) => {
    const raw = attrNum(node, name);
    return raw === null ? null : percentageToRatio(raw);
  };
  return {
    kind: outer ? "outer" : "inner",
    blur: attrNum(node, "blurRad") ?? 0,
    dist: attrNum(node, "dist") ?? 0,
    dir: angleUnitsToDegrees(attrNum(node, "dir")),
    color,
    align: attr(node, "algn"),
    scaleX: scale("sx"),
    scaleY: scale("sy"),
  };
}

/* --------------------------------------------------------------- geometry */

/** Parse the adjustment-value list of a preset geometry into ratios. */
export function parseAdjustments(avLst: XmlNode | null): Record<string, number> {
  const out: Record<string, number> = {};
  if (!avLst) return out;
  let index = 0;
  for (const gd of children(avLst, "gd")) {
    const name = attr(gd, "name") ?? `adj${index + 1}`;
    const fmla = attr(gd, "fmla") ?? "";
    const match = /^\s*val\s+(-?\d+(?:\.\d+)?)\s*$/.exec(fmla);
    if (match) {
      out[name] = Number(match[1]) / 100_000;
    }
    index += 1;
  }
  return out;
}

function parseCustomGeometry(custGeom: XmlNode): Geometry {
  const pathLst = child(custGeom, "pathLst");
  const path = pathLst ? child(pathLst, "path") : null;
  if (!path) return { type: "none" };
  const w = attrNum(path, "w") ?? 0;
  const h = attrNum(path, "h") ?? 0;
  const commands: PathCommand[] = [];
  const point = (pt: XmlNode | null): number[] => [
    attrNum(pt, "x") ?? 0,
    attrNum(pt, "y") ?? 0,
  ];
  for (const command of path.children) {
    switch (command.local) {
      case "moveTo":
        commands.push({ c: "M", pts: point(child(command, "pt")) });
        break;
      case "lnTo":
        commands.push({ c: "L", pts: point(child(command, "pt")) });
        break;
      case "cubicBezTo": {
        const pts = children(command, "pt");
        const flat: number[] = [];
        for (const pt of pts.slice(0, 3)) flat.push(...point(pt));
        if (flat.length === 6) commands.push({ c: "C", pts: flat });
        break;
      }
      case "quadBezTo": {
        const pts = children(command, "pt");
        const flat: number[] = [];
        for (const pt of pts.slice(0, 2)) flat.push(...point(pt));
        if (flat.length === 4) commands.push({ c: "Q", pts: flat });
        break;
      }
      case "arcTo": {
        // Converted to a straight segment: arcs in custom geometry are rare and a
        // wrong bezier approximation is worse than an honest chord.
        commands.push({ c: "L", pts: [0, 0] });
        break;
      }
      case "close":
        commands.push({ c: "Z", pts: [] });
        break;
      default:
        break;
    }
  }
  // Drop the placeholder segment used for unsupported arc commands.
  const cleaned = commands.filter((c) => !(c.c === "L" && c.pts[0] === 0 && c.pts[1] === 0));
  if (cleaned.length === 0) return { type: "none" };
  return { type: "custom", commands: cleaned, w: w || 100_000, h: h || 100_000 };
}

export function parseGeometry(props: XmlNode | null, ctx: DrawingContext): Geometry {
  if (!props) return { type: "none" };
  const prstGeom = child(props, "prstGeom");
  if (prstGeom) {
    const preset = attr(prstGeom, "prst") ?? "rect";
    // Report unsupported presets at parse time so the renderer never needs to
    // warn during a render pass, and so diagnostics are available up front.
    if (!isSupportedPreset(preset)) {
      ctx.warn?.(
        "geometry.preset-unsupported",
        `Preset shape "${preset}" is not implemented; a rectangle is drawn in its place.`,
      );
    }
    return { type: "preset", preset, adjust: parseAdjustments(child(prstGeom, "avLst")) };
  }
  const custGeom = child(props, "custGeom");
  if (custGeom) {
    const geometry = parseCustomGeometry(custGeom);
    if (geometry.type === "none") {
      ctx.warn?.("geometry.custom-empty", "Custom geometry contained no usable path commands");
    }
    return geometry;
  }
  return { type: "none" };
}

/** Theme palette in the order PowerPoint uses for automatic chart/series colours. */
export function accentPalette(theme: ThemeData, colorMap: ColorMap | undefined): RGBA[] {
  const map = colorMap ?? DEFAULT_COLOR_MAP;
  const out: RGBA[] = [];
  for (const slot of ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]) {
    const color = themeColor(theme, map[slot] ?? slot);
    if (color) out.push(color);
  }
  return out;
}
