/**
 * Fill/outline/shadow conversion for the SVG + HTML renderer.
 *
 * Geometry is drawn as SVG (precise vector paths) while text is drawn as HTML
 * (predictable line breaking, selectable and accessible). These helpers produce
 * the styling for both halves.
 */

import type { Fill, LineStyle, RGBA, Shadow } from "../model/types";
import { emuToPx } from "../parser/units";
import { rgbaToCss, rgbaToHex } from "../parser/color";

export interface SvgPaint {
  fill: string;
  fillRule: "nonzero" | "evenodd";
  fillOpacity: number;
  stroke: string;
  strokeWidth: number;
  strokeOpacity: number;
  strokeDasharray: string | undefined;
  strokeLinecap: "butt" | "round" | "square";
  strokeLinejoin: "miter" | "round" | "bevel";
}

const DASH_PATTERNS: Record<string, number[]> = {
  solid: [],
  sysDot: [1, 1],
  sysDash: [3, 1],
  sysDashDot: [3, 1, 1, 1],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
  dot: [1, 3],
  dash: [4, 3],
  lgDash: [8, 3],
  dashDot: [4, 3, 1, 3],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
};

export function dashArrayFor(dash: string | null, strokeWidth: number): string | undefined {
  if (!dash) return undefined;
  const pattern = DASH_PATTERNS[dash];
  if (!pattern || pattern.length === 0) return undefined;
  const width = strokeWidth > 0 ? strokeWidth : 1;
  return pattern.map((value) => value * width).join(" ");
}

export interface PaintOptions {
  /** needed to emit gradient references */
  gradientId?: string;
  /** fallback fill when the shape has no fill at all */
  defaultFill?: Fill | null;
}

/**
 * Resolve a fill into either a concrete paint colour or a `url(#gradient)` /
 * pattern reference.
 */
export function paintForFill(
  fill: Fill | null,
  options: PaintOptions = {},
): { paint: string; opacity: number; gradient?: GradientDef } {
  const effective = fill ?? options.defaultFill ?? null;
  if (!effective || effective.type === "none") {
    return { paint: "none", opacity: 1 };
  }
  switch (effective.type) {
    case "solid":
      return { paint: rgbaToHex(effective.color), opacity: effective.color.a };
    case "gradient": {
      if (options.gradientId) {
        return { paint: `url(#${options.gradientId})`, opacity: 1, gradient: gradientDefFor(effective, options.gradientId) };
      }
      const first = effective.stops[0];
      return { paint: first ? rgbaToHex(first.color) : "none", opacity: first ? first.color.a : 1 };
    }
    case "image":
      return { paint: "none", opacity: 1 };
    default:
      return { paint: "none", opacity: 1 };
  }
}

export interface GradientDef {
  id: string;
  kind: "linear" | "radial";
  angle: number;
  stops: { offset: number; color: string }[];
}

export function gradientDefFor(fill: Extract<Fill, { type: "gradient" }>, id: string): GradientDef {
  return {
    id,
    kind: fill.radial ? "radial" : "linear",
    angle: fill.angle,
    stops: fill.stops.map((stop) => ({
      offset: Math.max(0, Math.min(1, stop.pos)),
      color: rgbaToCss(stop.color),
    })),
  };
}

export function svgPaint(fill: Fill | null, line: LineStyle | null, gradientId?: string): SvgPaint {
  const fillPaint = paintForFill(fill, { gradientId });
  const strokeWidth = line && line.width !== null ? emuToPx(line.width) : 0;
  const strokeFill = line?.fill ?? null;
  const strokePaint = strokeFill && strokeFill.type !== "none" ? paintForFill(strokeFill) : null;
  const stroke = strokePaint ? strokePaint.paint : "none";
  const usableStrokeWidth = stroke === "none" ? 0 : strokeWidth;
  return {
    fill: fillPaint.paint,
    fillRule: "nonzero",
    fillOpacity: fillPaint.opacity,
    stroke,
    strokeWidth: usableStrokeWidth,
    strokeOpacity: strokePaint ? strokePaint.opacity : 1,
    strokeDasharray: dashArrayFor(line?.dash ?? null, usableStrokeWidth),
    strokeLinecap:
      line?.cap === "rnd" ? "round" : line?.cap === "sq" ? "square" : "butt",
    strokeLinejoin: line?.compound === undefined ? "round" : "round",
  };
}

/** CSS `background` value for a fill, used for cells, text boxes and thumbnails. */
export function cssBackground(fill: Fill | null): string | undefined {
  if (!fill || fill.type === "none") return undefined;
  switch (fill.type) {
    case "solid":
      return rgbaToCss(fill.color);
    case "gradient": {
      const stops = fill.stops
        .map((stop) => `${rgbaToCss(stop.color)} ${Math.round(stop.pos * 100)}%`)
        .join(", ");
      return fill.radial
        ? `radial-gradient(circle at 50% 50%, ${stops})`
        : `linear-gradient(${fill.angle}deg, ${stops})`;
    }
    default:
      return undefined;
  }
}

export interface CssShadow {
  filter: string;
}

/** PowerPoint shadows anchored at the shape's top-left; CSS drop-shadow matches closely. */
export function cssShadowFilter(shadow: Shadow | null | undefined): CssShadow | null {
  if (!shadow) return null;
  const radians = (shadow.dir * Math.PI) / 180;
  const dx = (Math.cos(radians) * emuToPx(shadow.dist)).toFixed(2);
  const dy = (Math.sin(radians) * emuToPx(shadow.dist)).toFixed(2);
  const blur = (emuToPx(shadow.blur) / 2).toFixed(2);
  return { filter: `drop-shadow(${dx}px ${dy}px ${blur}px ${rgbaToCss(shadow.color)})` };
}

export function cssBorder(fill: Fill | null, widthEmu: number | null): string | undefined {
  if (!fill || fill.type !== "solid" || widthEmu === null || widthEmu <= 0) return undefined;
  const width = Math.max(emuToPx(widthEmu), 0.75);
  return `${width.toFixed(2)}px solid ${rgbaToCss(fill.color)}`;
}

export function cssColor(color: RGBA | undefined): string | undefined {
  return color ? rgbaToCss(color) : undefined;
}

/** Deterministic id generation: must match between server render and hydration. */
export function hashString(value: string): string {
  let hash = 2166136261;
  for (let i = 0; i < value.length; i += 1) {
    hash ^= value.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(36);
}

export function gradientIdFor(scope: string, nodeId: string, fill: Fill): string {
  if (fill.type !== "gradient") return "";
  const signature = fill.stops.map((stop) => `${stop.pos}:${rgbaToHex(stop.color)}`).join("|");
  return `radv-g-${hashString(`${scope}|${nodeId}|${signature}|${fill.angle}|${fill.radial}`)}`;
}
