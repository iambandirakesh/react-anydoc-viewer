/**
 * Preset geometry (`a:prstGeom`) rendering.
 *
 * Every generator returns an SVG path expressed in the shape's own pixel box
 * (`0 0 w h`). The renderer draws it with a matching `viewBox` and leaves
 * rotation and flipping to CSS, which keeps the maths in one place.
 *
 * Adjustment defaults mirror the Office preset definitions (`adj/50000`,
 * `ss = min(w, h)`, angles in 1/60000 degree units). Unknown presets return
 * `null` so the caller can emit a diagnostic instead of guessing.
 */

export interface GeometryPath {
  d: string;
  fillRule: "nonzero" | "evenodd";
  /** false for open paths such as `arc` and connector presets */
  closed: boolean;
}

interface Ctx {
  w: number;
  h: number;
  /** shortest side, used by most Office preset formulas */
  ss: number;
  adj: (name: string, fallback: number) => number;
  ratio: (name: string, fallback: number) => number;
}

const n = (value: number): number => Math.round(value * 100) / 100;

function poly(points: number[][], closed = true): string {
  const first = points[0];
  if (!first) return "";
  let d = `M${n(first[0] ?? 0)},${n(first[1] ?? 0)}`;
  for (let i = 1; i < points.length; i += 1) {
    const point = points[i]!;
    d += `L${n(point[0] ?? 0)},${n(point[1] ?? 0)}`;
  }
  if (closed) d += "Z";
  return d;
}

function ellipsePath(cx: number, cy: number, rx: number, ry: number): string {
  const r = (value: number) => Math.max(value, 0.01);
  return (
    `M${n(cx - r(rx))},${n(cy)}` +
    `A${n(r(rx))},${n(r(ry))} 0 1 0 ${n(cx + r(rx))},${n(cy)}` +
    `A${n(r(rx))},${n(r(ry))} 0 1 0 ${n(cx - r(rx))},${n(cy)}Z`
  );
}

function pointOnEllipse(cx: number, cy: number, rx: number, ry: number, degrees: number): [number, number] {
  const rad = (degrees * Math.PI) / 180;
  return [cx + rx * Math.cos(rad), cy + ry * Math.sin(rad)];
}

function arcFromTo(
  cx: number,
  cy: number,
  rx: number,
  ry: number,
  startDeg: number,
  endDeg: number,
): string {
  const [x1, y1] = pointOnEllipse(cx, cy, rx, ry, startDeg);
  const [x2, y2] = pointOnEllipse(cx, cy, rx, ry, endDeg);
  const sweep = endDeg - startDeg;
  const large = Math.abs(sweep) > 180 ? 1 : 0;
  const direction = sweep >= 0 ? 1 : 0;
  return `M${n(x1)},${n(y1)}A${n(Math.max(rx, 0.01))},${n(Math.max(ry, 0.01))} 0 ${large} ${direction} ${n(x2)},${n(y2)}`;
}

function regularPolygon(cx: number, cy: number, rx: number, ry: number, sides: number, rotation = -90): string {
  const points: number[][] = [];
  for (let i = 0; i < sides; i += 1) {
    points.push(pointOnEllipse(cx, cy, rx, ry, rotation + (360 / sides) * i));
  }
  return poly(points);
}

function starPath(ctx: Ctx, points: number, innerRatio: number): string {
  const cx = ctx.w / 2;
  const cy = ctx.h / 2;
  const rx = ctx.w / 2;
  const ry = ctx.h / 2;
  const coords: number[][] = [];
  for (let i = 0; i < points * 2; i += 1) {
    const outer = i % 2 === 0;
    const [x, y] = pointOnEllipse(cx, cy, rx, ry, -90 + (180 / points) * i);
    coords.push(outer ? [x, y] : [cx + (x - cx) * innerRatio, cy + (y - cy) * innerRatio]);
  }
  return poly(coords);
}

/** Straight arrows. `headW`/`headH` are DrawingML adjustment ratios (0..1). */
function arrowPath(ctx: Ctx, direction: "right" | "left" | "up" | "down", headW: number, headH: number): string {
  const { w, h, ss } = ctx;
  const hw = headW * ss * 0.5;
  const hh = headH * ss * 0.5;
  if (direction === "right" || direction === "left") {
    const top = h / 2 - hh;
    const bottom = h / 2 + hh;
    const head = direction === "right" ? w - hw * 2 : hw * 2;
    const points =
      direction === "right"
        ? [[0, top], [head, top], [head, 0], [w, h / 2], [head, h], [head, bottom], [0, bottom]]
        : [[w, top], [head, top], [head, 0], [0, h / 2], [head, h], [head, bottom], [w, bottom]];
    return poly(points);
  }
  const left = w / 2 - hh;
  const right = w / 2 + hh;
  const head = direction === "up" ? hw * 2 : h - hw * 2;
  const points =
    direction === "up"
      ? [[left, h], [left, head], [0, head], [w / 2, 0], [w, head], [right, head], [right, h]]
      : [[left, 0], [left, head], [0, head], [w / 2, h], [w, head], [right, head], [right, 0]];
  return poly(points);
}

type Generator = (ctx: Ctx) => GeometryPath;

const GENERATORS: Record<string, Generator> = {};

function define(presets: string | string[], generator: Generator): void {
  for (const preset of Array.isArray(presets) ? presets : [presets]) {
    GENERATORS[preset] = generator;
  }
}

const closedPath = (d: string, fillRule: "nonzero" | "evenodd" = "nonzero"): GeometryPath => ({
  d,
  fillRule,
  closed: true,
});

/* ----------------------------------------------------------------- basics */

define("rect", (ctx) => closedPath(poly([[0, 0], [ctx.w, 0], [ctx.w, ctx.h], [0, ctx.h]])));

const ROUND_RECT: Generator = (ctx) => {
  const r = Math.min(ctx.ratio("adj", 0.16667) * ctx.ss, ctx.w / 2, ctx.h / 2);
  const { w, h } = ctx;
  return closedPath(
    `M${n(r)},0L${n(w - r)},0A${n(r)},${n(r)} 0 0 1 ${n(w)},${n(r)}` +
      `L${n(w)},${n(h - r)}A${n(r)},${n(r)} 0 0 1 ${n(w - r)},${n(h)}` +
      `L${n(r)},${n(h)}A${n(r)},${n(r)} 0 0 1 0,${n(h - r)}` +
      `L0,${n(r)}A${n(r)},${n(r)} 0 0 1 ${n(r)},0Z`,
  );
};
define(["roundRect", "round1Rect"], ROUND_RECT);

define(["round2SameRect", "round2DiagRect"], (ctx) => {
  const r = Math.min(ctx.ratio("adj1", 0.16667) * ctx.ss, ctx.w / 2, ctx.h / 2);
  const { w, h } = ctx;
  return closedPath(
    `M${n(r)},0L${n(w)},0L${n(w)},${n(h - r)}A${n(r)},${n(r)} 0 0 1 ${n(w - r)},${n(h)}` +
      `L0,${n(h)}L0,${n(r)}A${n(r)},${n(r)} 0 0 1 ${n(r)},0Z`,
  );
});

function snippedRect(ctx: Ctx, corners: { tr?: boolean; bl?: boolean }, cut: number): string {
  const { w, h } = ctx;
  const c = Math.min(cut, w, h);
  const pts: number[][] = [[0, 0]];
  if (corners.tr) pts.push([w - c, 0], [w, c]);
  else pts.push([w, 0]);
  if (corners.bl) pts.push([c, h], [0, h - c]);
  else pts.push([0, h]);
  return poly(pts);
}

define(["snip1Rect", "snipRoundRect"], (ctx) => closedPath(snippedRect(ctx, { tr: true }, ctx.ratio("adj", 0.16667) * ctx.ss)));
define("snip2SameRect", (ctx) => closedPath(snippedRect(ctx, { tr: true, bl: true }, ctx.ratio("adj1", 0.16667) * ctx.ss)));
define("snip2DiagRect", (ctx) => closedPath(snippedRect(ctx, { tr: true, bl: true }, ctx.ratio("adj1", 0.16667) * ctx.ss)));

define(["ellipse", "circle"], (ctx) => closedPath(ellipsePath(ctx.w / 2, ctx.h / 2, ctx.w / 2, ctx.h / 2)));

define("triangle", (ctx) => closedPath(poly([[ctx.w / 2, 0], [ctx.w, ctx.h], [0, ctx.h]])));
define("rtTriangle", (ctx) => closedPath(poly([[0, 0], [0, ctx.h], [ctx.w, ctx.h]])));
define(["diamond", "rhombus"], (ctx) =>
  closedPath(poly([[ctx.w / 2, 0], [ctx.w, ctx.h / 2], [ctx.w / 2, ctx.h], [0, ctx.h / 2]])),
);

define("parallelogram", (ctx) => {
  const cut = (ctx.ratio("adj", 0.25) * ctx.ss) / 2;
  return closedPath(poly([[cut, 0], [ctx.w - cut, 0], [ctx.w, ctx.h], [0, ctx.h]]));
});

define("trapezoid", (ctx) => {
  const cut = (ctx.ratio("adj", 0.25) * ctx.ss) / 2;
  return closedPath(poly([[cut, 0], [ctx.w - cut, 0], [ctx.w, ctx.h], [0, ctx.h]]));
});

define("pentagon", (ctx) => closedPath(regularPolygon(ctx.w / 2, ctx.h / 2, ctx.w / 2, ctx.h / 2, 5)));
define("hexagon", (ctx) => {
  const cut = (ctx.ratio("adj", 0.25) * ctx.ss) / 2;
  return closedPath(
    poly([[cut, 0], [ctx.w - cut, 0], [ctx.w, ctx.h / 2], [ctx.w - cut, ctx.h], [cut, ctx.h], [0, ctx.h / 2]]),
  );
});
define("heptagon", (ctx) => closedPath(regularPolygon(ctx.w / 2, ctx.h / 2, ctx.w / 2, ctx.h / 2, 7)));
define("octagon", (ctx) => {
  const cut = (ctx.ratio("adj", 0.29289) * ctx.ss) / 2;
  const { w, h } = ctx;
  return closedPath(
    poly([
      [cut, 0], [w - cut, 0], [w, cut], [w, h - cut],
      [w - cut, h], [cut, h], [0, h - cut], [0, cut],
    ]),
  );
});
define("decagon", (ctx) => closedPath(regularPolygon(ctx.w / 2, ctx.h / 2, ctx.w / 2, ctx.h / 2, 10)));
define("dodecagon", (ctx) => closedPath(regularPolygon(ctx.w / 2, ctx.h / 2, ctx.w / 2, ctx.h / 2, 12)));

define("plus", (ctx) => {
  const a = (ctx.ratio("adj", 0.25) * ctx.ss) / 2;
  const { w, h } = ctx;
  return closedPath(
    poly([
      [a, 0], [w - a, 0], [w - a, a], [w, a], [w, h - a], [w - a, h - a],
      [w - a, h], [a, h], [a, h - a], [0, h - a], [0, a], [a, a],
    ]),
  );
});

define("frame", (ctx) => {
  const x = Math.min(ctx.ratio("adj1", 0.125) * ctx.ss, ctx.w / 2);
  const y = Math.min(ctx.ratio("adj2", 0.125) * ctx.ss, ctx.h / 2);
  const { w, h } = ctx;
  return closedPath(poly([[0, 0], [w, 0], [w, h], [0, h]]) + poly([[x, y], [x, h - y], [w - x, h - y], [w - x, y]]), "evenodd");
});

define("halfFrame", (ctx) => {
  const x = Math.min(ctx.ratio("adj1", 0.33333) * ctx.ss, ctx.w / 2);
  const y = Math.min(ctx.ratio("adj2", 0.33333) * ctx.h, ctx.h / 2);
  const { w, h } = ctx;
  // Upper-left "L" band, drawn as an outer rect minus an inner notch.
  return closedPath(
    poly([[0, 0], [w, 0], [w, y], [x, y], [x, h], [0, h]]) + poly([[x, y], [x, y], [x, y]]),
  );
});

define("corner", (ctx) => {
  const x = Math.min(ctx.ratio("adj1", 0.5) * ctx.ss, ctx.w);
  const y = Math.min(ctx.ratio("adj2", 0.5) * ctx.ss, ctx.h);
  return closedPath(poly([[0, 0], [x, 0], [0, y]]));
});

define("diagStripe", (ctx) => {
  const x = ctx.ratio("adj", 0.5) * ctx.ss;
  const { w, h } = ctx;
  return closedPath(poly([[0, h], [w - x, 0], [w, 0], [0, h - x]]));
});

define("donut", (ctx) => {
  const t = ctx.ratio("adj", 0.25) * ctx.ss;
  const rx = ctx.w / 2;
  const ry = ctx.h / 2;
  const d = ellipsePath(rx, ry, rx, ry) + ellipsePath(rx, ry, rx - t, ry - t);
  return closedPath(d, "evenodd");
});

define("noSmoking", (ctx) => {
  const t = ctx.ratio("adj", 0.1875) * ctx.ss;
  const rx = ctx.w / 2;
  const ry = ctx.h / 2;
  const d = ellipsePath(rx, ry, rx, ry) + ellipsePath(rx, ry, rx - t, ry - t);
  return closedPath(d, "evenodd");
});

define("blockArc", (ctx) => {
  const start = ctx.adj("adj1", 10_800_000) / 60_000;
  const sweep = (ctx.adj("adj2", 0) || 0) / 60_000;
  const end = start + (sweep === 0 ? 270 : sweep);
  const t = ctx.ratio("adj3", 0.25);
  const rx = ctx.w / 2;
  const ry = ctx.h / 2;
  const innerRx = Math.max(rx * (1 - t), 0.01);
  const innerRy = Math.max(ry * (1 - t), 0.01);
  const [sx, sy] = pointOnEllipse(rx, ry, rx, ry, start);
  const [ex, ey] = pointOnEllipse(rx, ry, rx, ry, end);
  const [isx, isy] = pointOnEllipse(rx, ry, innerRx, innerRy, end);
  const [iex, iey] = pointOnEllipse(rx, ry, innerRx, innerRy, start);
  const large = Math.abs(end - start) > 180 ? 1 : 0;
  return closedPath(
    `M${n(sx)},${n(sy)}A${n(rx)},${n(ry)} 0 ${large} 1 ${n(ex)},${n(ey)}` +
      `L${n(isx)},${n(isy)}A${n(innerRx)},${n(innerRy)} 0 ${large} 0 ${n(iex)},${n(iey)}Z`,
  );
});

define("pie", (ctx) => {
  const start = ctx.adj("adj1", 0) / 60_000;
  const end = ctx.adj("adj2", 16_200_000) / 60_000;
  return closedPath(`${arcFromTo(ctx.w / 2, ctx.h / 2, ctx.w / 2, ctx.h / 2, start, end)}L${n(ctx.w / 2)},${n(ctx.h / 2)}Z`);
});

define("chord", (ctx) =>
  closedPath(arcFromTo(ctx.w / 2, ctx.h / 2, ctx.w / 2, ctx.h / 2, ctx.adj("adj1", 2_700_000) / 60_000, ctx.adj("adj2", 16_200_000) / 60_000) + "Z"),
);

define("arc", (ctx) => ({
  d: arcFromTo(ctx.w / 2, ctx.h / 2, ctx.w / 2, ctx.h / 2, ctx.adj("adj1", 16_200_000) / 60_000, ctx.adj("adj2", 0) / 60_000),
  fillRule: "nonzero",
  closed: false,
}));

define("moon", (ctx) => {
  const t = 1 - ctx.ratio("adj", 0.5) * 0.7 - 0.3;
  const rx = ctx.w / 2;
  const ry = ctx.h / 2;
  return closedPath(
    `M0,0A${n(rx)},${n(ry)} 0 1 1 0,${n(ctx.h)}` +
      `A${n(Math.max(rx * t, 0.01))},${n(ry)} 0 1 0 0,0Z`,
  );
});

define("teardrop", (ctx) => {
  const rx = ctx.w / 2;
  const ry = ctx.h / 2;
  return closedPath(`${ellipsePath(rx, ry, rx, ry)}M0,0L${n(rx * 0.55)},0L${n(rx * 0.4)},${n(ry * 0.4)}Z`);
});

define("sun", (ctx) => {
  const inner = Math.max(Math.min(ctx.w, ctx.h) / 2 - ctx.ratio("adj", 0.25) * ctx.ss, 0.01);
  return closedPath(
    `${starPath(ctx, 8, inner / (Math.min(ctx.w, ctx.h) / 2 || 1))}${ellipsePath(ctx.w / 2, ctx.h / 2, inner, inner)}`,
  );
});

define("heart", (ctx) => {
  const { w, h } = ctx;
  return closedPath(
    `M${n(w / 2)},${n(h)}C${n(w / 8)},${n(h * 0.68)} 0,${n(h * 0.45)} 0,${n(h * 0.28)}` +
      `A${n(w / 4)},${n(h / 4)} 0 0 1 ${n(w / 2)},${n(h * 0.28)}` +
      `A${n(w / 4)},${n(h / 4)} 0 0 1 ${n(w)},${n(h * 0.28)}` +
      `C${n(w)},${n(h * 0.45)} ${n((w * 7) / 8)},${n(h * 0.68)} ${n(w / 2)},${n(h)}Z`,
  );
});

define("cloud", (ctx) => {
  const { w, h } = ctx;
  return closedPath(
    `M${n(w * 0.25)},${n(h)}A${n(w * 0.16)},${n(h * 0.2)} 0 0 1 ${n(w * 0.16)},${n(h * 0.66)}` +
      `A${n(w * 0.17)},${n(h * 0.22)} 0 0 1 ${n(w * 0.42)},${n(h * 0.28)}` +
      `A${n(w * 0.19)},${n(h * 0.25)} 0 0 1 ${n(w * 0.72)},${n(h * 0.3)}` +
      `A${n(w * 0.16)},${n(h * 0.2)} 0 0 1 ${n(w * 0.86)},${n(h * 0.62)}` +
      `A${n(w * 0.14)},${n(h * 0.19)} 0 0 1 ${n(w * 0.75)},${n(h)}Z`,
  );
});

define("smileyFace", (ctx) => {
  const { w, h } = ctx;
  const eyeR = Math.max(Math.min(w, h) * 0.06, 1);
  const mouth = `M${n(w * 0.28)},${n(h * 0.6)}A${n(w * 0.28)},${n(h * 0.25)} 0 0 0 ${n(w * 0.72)},${n(h * 0.6)}`;
  return closedPath(
    `${ellipsePath(w / 2, h / 2, w / 2, h / 2)}` +
      `${ellipsePath(w * 0.33, h * 0.36, eyeR, eyeR)}` +
      `${ellipsePath(w * 0.67, h * 0.36, eyeR, eyeR)}${mouth}`,
    "evenodd",
  );
});

define("lightningBolt", (ctx) => {
  const { w, h } = ctx;
  return closedPath(
    poly([
      [w * 0.36, 0], [w, 0], [w * 0.56, h * 0.42], [w * 0.86, h * 0.42],
      [w * 0.24, h], [w * 0.4, h * 0.55], [w * 0.1, h * 0.55],
    ]),
  );
});

/* ----------------------------------------------------------------- arrows */

define(["rightArrow", "notchedRightArrow", "stripedRightArrow"], (ctx) =>
  closedPath(arrowPath(ctx, "right", ctx.ratio("adj1", 0.5), ctx.ratio("adj2", 0.5))),
);
define("leftArrow", (ctx) => closedPath(arrowPath(ctx, "left", ctx.ratio("adj1", 0.5), ctx.ratio("adj2", 0.5))));
define("upArrow", (ctx) => closedPath(arrowPath(ctx, "up", ctx.ratio("adj1", 0.5), ctx.ratio("adj2", 0.5))));
define("downArrow", (ctx) => closedPath(arrowPath(ctx, "down", ctx.ratio("adj1", 0.5), ctx.ratio("adj2", 0.5))));

define("leftRightArrow", (ctx) => {
  const { w, h } = ctx;
  const head = Math.min(ctx.ratio("adj1", 0.5) * ctx.ss, w / 2);
  const half = (ctx.ratio("adj2", 0.5) * ctx.ss) / 2;
  const top = h / 2 - half;
  const bottom = h / 2 + half;
  return closedPath(
    poly([
      [0, h / 2], [head, 0], [head, top], [w - head, top], [w - head, 0],
      [w, h / 2], [w - head, h], [w - head, bottom], [head, bottom], [head, h],
    ]),
  );
});

define("upDownArrow", (ctx) => {
  const { w, h } = ctx;
  const head = Math.min(ctx.ratio("adj1", 0.5) * ctx.ss, h / 2);
  const half = (ctx.ratio("adj2", 0.5) * ctx.ss) / 2;
  const left = w / 2 - half;
  const right = w / 2 + half;
  return closedPath(
    poly([
      [w / 2, 0], [w, head], [right, head], [right, h - head], [w, h - head],
      [w / 2, h], [0, h - head], [left, h - head], [left, head], [0, head],
    ]),
  );
});

define("quadArrow", (ctx) => {
  const { w, h } = ctx;
  const cx = w / 2;
  const cy = h / 2;
  const a = Math.min(ctx.ratio("adj1", 0.225) * ctx.ss * 0.5, Math.min(w, h) / 4) || Math.min(w, h) / 8;
  return closedPath(
    poly([
      [cx - a, 0], [cx + a, 0], [cx + a, cy - a], [w, cy - a], [w, cy + a],
      [cx + a, cy + a], [cx + a, h], [cx - a, h], [cx - a, cy + a],
      [0, cy + a], [0, cy - a], [cx - a, cy - a],
    ]),
  );
});

define("homePlate", (ctx) => {
  const x = ctx.ratio("adj", 0.5) * ctx.ss;
  return closedPath(poly([[0, 0], [ctx.w - x, 0], [ctx.w, ctx.h / 2], [ctx.w - x, ctx.h], [0, ctx.h]]));
});

define("chevron", (ctx) => {
  const x = ctx.ratio("adj", 0.5) * ctx.ss;
  return closedPath(
    poly([[0, 0], [ctx.w - x, 0], [ctx.w, ctx.h / 2], [ctx.w - x, ctx.h], [0, ctx.h], [x, ctx.h / 2]]),
  );
});

define("rightArrowCallout", (ctx) => closedPath(arrowPath(ctx, "right", ctx.ratio("adj1", 0.5), ctx.ratio("adj2", 0.5))));
define("leftArrowCallout", (ctx) => closedPath(arrowPath(ctx, "left", ctx.ratio("adj1", 0.5), ctx.ratio("adj2", 0.5))));
define("upArrowCallout", (ctx) => closedPath(arrowPath(ctx, "up", ctx.ratio("adj1", 0.5), ctx.ratio("adj2", 0.5))));
define("downArrowCallout", (ctx) => closedPath(arrowPath(ctx, "down", ctx.ratio("adj1", 0.5), ctx.ratio("adj2", 0.5))));

function calloutBody(ctx: Ctx, round: boolean): string {
  const { w, h } = ctx;
  if (!round) return poly([[0, 0], [w, 0], [w, h], [0, h]]);
  const r = Math.min(ctx.ratio("adj3", 0.16667) * ctx.ss, w / 2, h / 2);
  return (
    `M${n(r)},0L${n(w - r)},0A${n(r)},${n(r)} 0 0 1 ${n(w)},${n(r)}` +
    `L${n(w)},${n(h - r)}A${n(r)},${n(r)} 0 0 1 ${n(w - r)},${n(h)}` +
    `L${n(r)},${n(h)}A${n(r)},${n(r)} 0 0 1 0,${n(h - r)}` +
    `L0,${n(r)}A${n(r)},${n(r)} 0 0 1 ${n(r)},0Z`
  );
}

function calloutTail(ctx: Ctx): string {
  const { w, h } = ctx;
  const dx = ctx.ratio("adj1", -0.20833);
  const dy = ctx.ratio("adj2", 0.625);
  const tipX = w / 2 + dx * (w / 2);
  const tipY = h / 2 + dy * (h / 2);
  const baseX = w / 2;
  const baseY = h / 2;
  return `M${n(baseX + 1)},${n(baseY)}L${n(tipX)},${n(tipY)}L${n(baseX - 1)},${n(baseY)}Z`;
}

define("wedgeRectCallout", (ctx) => closedPath(`${calloutBody(ctx, false)}${calloutTail(ctx)}`));
define("wedgeRoundRectCallout", (ctx) => closedPath(`${calloutBody(ctx, true)}${calloutTail(ctx)}`));
define("wedgeEllipseCallout", (ctx) =>
  closedPath(`${ellipsePath(ctx.w / 2, ctx.h / 2, ctx.w / 2, ctx.h / 2)}${calloutTail(ctx)}`),
);

/* -------------------------------------------------------------- flow chart */

define("flowChartProcess", (ctx) => closedPath(poly([[0, 0], [ctx.w, 0], [ctx.w, ctx.h], [0, ctx.h]])));
define("flowChartAlternateProcess", ROUND_RECT);
define("flowChartOfflineStorage", (ctx) => closedPath(poly([[0, 0], [ctx.w, 0], [ctx.w, ctx.h], [0, ctx.h]])));

define("flowChartDecision", (ctx) =>
  closedPath(poly([[ctx.w / 2, 0], [ctx.w, ctx.h / 2], [ctx.w / 2, ctx.h], [0, ctx.h / 2]])),
);

define("flowChartInputOutput", (ctx) => {
  const x = ctx.ratio("adj", 0.25) * ctx.ss;
  return closedPath(poly([[x, 0], [ctx.w, 0], [ctx.w - x, ctx.h], [0, ctx.h]]));
});

define("flowChartPredefinedProcess", (ctx) => {
  const x = Math.max(Math.min(ctx.ratio("adj", 0.08333) * ctx.ss, ctx.w / 2 - 1), 0.5);
  return closedPath(
    `${poly([[0, 0], [ctx.w, 0], [ctx.w, ctx.h], [0, ctx.h]])}` +
      `M${n(x)},0L${n(x)},${n(ctx.h)}M${n(ctx.w - x)},0L${n(ctx.w - x)},${n(ctx.h)}`,
  );
});

define("flowChartInternalStorage", (ctx) => {
  const x = Math.max(Math.min(ctx.ratio("adj1", 0.125) * ctx.ss, ctx.w - 1), 0.5);
  const y = Math.max(Math.min(ctx.ratio("adj2", 0.125) * ctx.ss, ctx.h - 1), 0.5);
  return closedPath(
    `${poly([[0, 0], [ctx.w, 0], [ctx.w, ctx.h], [0, ctx.h]])}` +
      `M${n(x)},0L${n(x)},${n(ctx.h)}M0,${n(y)}L${n(ctx.w)},${n(y)}`,
  );
});

define("flowChartDocument", (ctx) => {
  const { w, h } = ctx;
  return closedPath(
    `M0,0L${n(w)},0L${n(w)},${n(h * 0.82)}C${n(w * 0.75)},${n(h)} ${n(w * 0.25)},${n(h * 0.64)} 0,${n(h * 0.82)}Z`,
  );
});

define("flowChartMultidocument", (ctx) => {
  const { w, h } = ctx;
  return closedPath(
    `M0,${n(h * 0.18)}L${n(w * 0.82)},${n(h * 0.18)}L${n(w * 0.82)},${n(h)}` +
      `C${n(w * 0.6)},${n(h * 0.92)} ${n(w * 0.22)},${n(h)} 0,${n(h * 0.9)}Z` +
      `M${n(w * 0.1)},${n(h * 0.09)}L${n(w * 0.92)},${n(h * 0.09)}L${n(w * 0.92)},${n(h * 0.91)}`,
  );
});

define("flowChartTerminator", (ctx) => {
  const rx = Math.min(ctx.h / 2, ctx.w / 2);
  const { w, h } = ctx;
  return closedPath(
    `M${n(rx)},0L${n(w - rx)},0A${n(rx)},${n(h / 2)} 0 0 1 ${n(w - rx)},${n(h)}` +
      `L${n(rx)},${n(h)}A${n(rx)},${n(h / 2)} 0 0 1 ${n(rx)},0Z`,
  );
});

define("flowChartPreparation", (ctx) => {
  const x = ctx.ratio("adj", 0.25) * ctx.ss;
  return closedPath(
    poly([[x, 0], [ctx.w - x, 0], [ctx.w, ctx.h / 2], [ctx.w - x, ctx.h], [x, ctx.h], [0, ctx.h / 2]]),
  );
});

define("flowChartManualInput", (ctx) =>
  closedPath(poly([[0, ctx.h * 0.2], [ctx.w, 0], [ctx.w, ctx.h], [0, ctx.h]])),
);

define("flowChartManualOperation", (ctx) =>
  closedPath(poly([[0, 0], [ctx.w, 0], [ctx.w * 0.8, ctx.h], [ctx.w * 0.2, ctx.h]])),
);

define("flowChartConnector", (ctx) => closedPath(ellipsePath(ctx.w / 2, ctx.h / 2, ctx.w / 2, ctx.h / 2)));

define("flowChartOffpageConnector", (ctx) =>
  closedPath(poly([[0, 0], [ctx.w, 0], [ctx.w, ctx.h * 0.72], [ctx.w / 2, ctx.h], [0, ctx.h * 0.72]])),
);

define("flowChartPunchedCard", (ctx) => {
  const x = Math.min(ctx.ratio("adj", 0.125) * ctx.ss, ctx.h);
  return closedPath(poly([[0, 0], [ctx.w, 0], [ctx.w, ctx.h], [x, ctx.h], [x, x], [0, x]]));
});

define("flowChartPunchedTape", (ctx) => {
  const { w, h } = ctx;
  const y = h * 0.2;
  return closedPath(
    `M0,${n(y)}C${n(w * 0.25)},${n(y * 2)} ${n(w * 0.75)},0 ${n(w)},${n(y)}` +
      `L${n(w)},${n(h - y)}C${n(w * 0.75)},${n(h)} ${n(w * 0.25)},${n(h - y * 2)} 0,${n(h - y)}Z`,
  );
});

define("flowChartSummingJunction", (ctx) => {
  const { w, h } = ctx;
  const cx = w / 2;
  const cy = h / 2;
  const rx = w / 2;
  const ry = h / 2;
  const dx = rx * Math.SQRT1_2;
  const dy = ry * Math.SQRT1_2;
  return closedPath(
    `${ellipsePath(cx, cy, rx, ry)}` +
      `M${n(cx - dx)},${n(cy - dy)}L${n(cx + dx)},${n(cy + dy)}` +
      `M${n(cx + dx)},${n(cy - dy)}L${n(cx - dx)},${n(cy + dy)}`,
  );
});

define("flowChartOr", (ctx) =>
  closedPath(`${poly([[0, 0], [ctx.w, 0], [ctx.w, ctx.h], [0, ctx.h]])}${ellipsePath(ctx.w / 2, ctx.h / 2, ctx.w / 2, ctx.h / 2)}`),
);

define("flowChartCollate", (ctx) =>
  closedPath(
    `${poly([[0, 0], [ctx.w, 0], [ctx.w / 2, ctx.h / 2]])}${poly([[0, ctx.h], [ctx.w, ctx.h], [ctx.w / 2, ctx.h / 2]])}`,
  ),
);

define("flowChartSort", (ctx) =>
  closedPath(
    `${poly([[ctx.w / 2, 0], [ctx.w, ctx.h / 2], [ctx.w / 2, ctx.h]])}${poly([[0, ctx.h / 2], [ctx.w / 2, 0], [ctx.w / 2, ctx.h]])}`,
  ),
);

define("flowChartExtract", (ctx) => closedPath(poly([[0, ctx.h], [ctx.w / 2, 0], [ctx.w, ctx.h]])));
define("flowChartMerge", (ctx) => closedPath(poly([[0, 0], [ctx.w, 0], [ctx.w / 2, ctx.h]])));

define("flowChartOnlineStorage", (ctx) => {
  const x = Math.max(ctx.ratio("adj", 0.2) * ctx.ss, 1);
  const { w, h } = ctx;
  return closedPath(
    `M${n(x)},0L${n(w - x)},0A${n(x)},${n(h / 2)} 0 0 1 ${n(w - x)},${n(h)}L${n(x)},${n(h)}A${n(x)},${n(h / 2)} 0 0 1 ${n(x)},0Z`,
  );
});

define("flowChartMagneticDisk", (ctx) => {
  const ry = ctx.h / 6;
  const { w, h } = ctx;
  return closedPath(
    `M0,${n(ry)}A${n(w / 2)},${n(ry)} 0 0 1 ${n(w)},${n(ry)}L${n(w)},${n(h - ry)}A${n(w / 2)},${n(ry)} 0 0 1 0,${n(h - ry)}Z`,
  );
});

const DRUM: Generator = (ctx) => {
  const rx = Math.max(ctx.w / 8, 1);
  const { w, h } = ctx;
  return closedPath(
    `M${n(rx)},0L${n(w - rx)},0A${n(rx)},${n(h / 2)} 0 0 1 ${n(w - rx)},${n(h)}L${n(rx)},${n(h)}A${n(rx)},${n(h / 2)} 0 0 1 ${n(rx)},0Z`,
  );
};
define(["flowChartMagneticDrum", "flowChartMagneticTape"], DRUM);

define("flowChartDisplay", (ctx) => {
  const { w, h } = ctx;
  return closedPath(
    `M${n(w * 0.22)},0L${n(w * 0.78)},0A${n(w * 0.22)},${n(h / 2)} 0 0 1 ${n(w * 0.78)},${n(h)}` +
      `L${n(w * 0.22)},${n(h)}L${n(w * 0.06)},${n(h / 2)}Z`,
  );
});

define("flowChartDelay", (ctx) => {
  const rx = Math.min(ctx.h / 2, ctx.w / 2);
  const { w, h } = ctx;
  return closedPath(
    `M0,0L${n(w - rx)},0A${n(rx)},${n(h / 2)} 0 0 1 ${n(w - rx)},${n(h)}L0,${n(h)}Z`,
  );
});

/* ------------------------------------------------------------ lines/conns */

define(["line", "straightConnector1"], (ctx) => ({
  d: `M0,0L${n(ctx.w)},${n(ctx.h)}`,
  fillRule: "nonzero",
  closed: false,
}));

define(["bentConnector2", "bentConnector3", "bentConnector4", "bentConnector5"], (ctx) => ({
  d: `M0,0L${n(ctx.w)},0L${n(ctx.w)},${n(ctx.h)}`,
  fillRule: "nonzero",
  closed: false,
}));

define(["curvedConnector2", "curvedConnector3", "curvedConnector4", "curvedConnector5"], (ctx) => {
  const { w, h } = ctx;
  return {
    d: `M0,0C${n(w * 0.5)},0 ${n(w * 0.5)},${n(h)} ${n(w)},${n(h)}`,
    fillRule: "nonzero",
    closed: false,
  };
});

/* ------------------------------------------------------------------ stars */

/** `adj` defaults from the Office preset definitions; inner radius = adj/50000. */
const STAR_ADJUSTMENTS: Record<string, number> = {
  star4: 12_500,
  star5: 19_098,
  star6: 28_868,
  star7: 34_601,
  star8: 38_462,
  star10: 42_533,
  star12: 45_922,
  star16: 48_992,
  star24: 51_931,
  star32: 53_863,
};

for (const [preset, defaultAdj] of Object.entries(STAR_ADJUSTMENTS)) {
  const points = Number(preset.replace("star", ""));
  define(preset, (ctx) => {
    const inner = Math.min(Math.max(ctx.adj("adj", defaultAdj) / 50_000, 0), 1);
    return closedPath(starPath(ctx, points, inner));
  });
}

/* -------------------------------------------------------------- dispatch */

export function isSupportedPreset(preset: string): boolean {
  return Object.prototype.hasOwnProperty.call(GENERATORS, preset);
}

export function supportedPresetNames(): string[] {
  return Object.keys(GENERATORS).sort();
}

/**
 * Convert a preset geometry into an SVG path in the shape's pixel box.
 * Unknown presets return `null` so the caller can report a diagnostic.
 */
export function presetToPath(
  preset: string,
  adjust: Record<string, number>,
  w: number,
  h: number,
): GeometryPath | null {
  const generator = GENERATORS[preset];
  if (!generator) return null;
  if (!(w > 0) || !(h > 0)) return null;
  const ctx: Ctx = {
    w,
    h,
    ss: Math.min(w, h),
    adj: (name, fallback) => {
      const value = adjust[name];
      return typeof value === "number" && Number.isFinite(value) ? value : fallback;
    },
    ratio: (name, fallback) => {
      const value = adjust[name];
      return typeof value === "number" && Number.isFinite(value) ? value : fallback;
    },
  };
  const result = generator(ctx);
  if (!result || result.d.length === 0) return null;
  return result;
}

export function customGeometryToPath(
  commands: { c: string; pts: number[] }[],
  w: number,
  h: number,
  sourceW: number,
  sourceH: number,
): GeometryPath {
  const sx = sourceW > 0 ? w / sourceW : 1;
  const sy = sourceH > 0 ? h / sourceH : 1;
  let d = "";
  for (const command of commands) {
    const p = command.pts;
    switch (command.c) {
      case "M":
        d += `M${n((p[0] ?? 0) * sx)},${n((p[1] ?? 0) * sy)}`;
        break;
      case "L":
        d += `L${n((p[0] ?? 0) * sx)},${n((p[1] ?? 0) * sy)}`;
        break;
      case "C":
        d +=
          `C${n((p[0] ?? 0) * sx)},${n((p[1] ?? 0) * sy)} ` +
          `${n((p[2] ?? 0) * sx)},${n((p[3] ?? 0) * sy)} ` +
          `${n((p[4] ?? 0) * sx)},${n((p[5] ?? 0) * sy)}`;
        break;
      case "Q":
        d += `Q${n((p[0] ?? 0) * sx)},${n((p[1] ?? 0) * sy)} ${n((p[2] ?? 0) * sx)},${n((p[3] ?? 0) * sy)}`;
        break;
      case "Z":
        d += "Z";
        break;
      default:
        break;
    }
  }
  return { d, fillRule: "nonzero", closed: d.includes("Z") };
}
