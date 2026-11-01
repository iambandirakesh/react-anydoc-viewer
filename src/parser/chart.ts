/**
 * Chart parsing (`c:chartSpace`).
 *
 * Charts are read from the *cached* values embedded in the chart part. That is
 * exactly what PowerPoint stores for offline rendering, and it means the viewer
 * never needs to evaluate a spreadsheet formula from the embedded workbook.
 */

import type { ChartData, ChartSeries, Fill, LineStyle, RGBA } from "../model/types";
import { attr, attrBool, attrNum, child, children, find, findAll, type XmlNode } from "./xml";
import { clamp } from "./units";
import { isColorElement, resolveColorElement } from "./color";
import { parseFill, parseLine, type DrawingContext } from "./drawing";
import { parseTextBody, textBodyPlainText, type TextContext } from "./text";

const CHART_GROUP_ELEMENTS = [
  "barChart",
  "bar3DChart",
  "lineChart",
  "line3DChart",
  "pieChart",
  "pie3DChart",
  "doughnutChart",
  "areaChart",
  "area3DChart",
  "scatterChart",
  "bubbleChart",
  "radarChart",
  "ofPieChart",
  "stockChart",
  "surfaceChart",
  "surface3DChart",
] as const;

interface CachedRange {
  values: (string | number)[];
  numeric: boolean;
}

/** Read the `<c:pt>` points of one cache element. */
function readPoints(cache: XmlNode, numeric: boolean): CachedRange | null {
  const count = attrNum(child(cache, "ptCount"), "val") ?? 0;
  const values: (string | number | null)[] = new Array(Math.max(count, 0)).fill(null);
  let maxIndex = -1;
  for (const pt of children(cache, "pt")) {
    const index = attrNum(pt, "idx");
    if (index === null || index < 0) continue;
    const v = child(pt, "v");
    const raw = v ? v.text : "";
    while (values.length <= index) values.push(null);
    if (numeric) {
      const parsed = Number(raw);
      values[index] = Number.isFinite(parsed) ? parsed : null;
    } else {
      values[index] = raw;
    }
    if (index > maxIndex) maxIndex = index;
  }
  if (values.length === 0) return null;
  const trimmed = values.slice(0, Math.max(maxIndex + 1, 0));
  if (numeric) {
    return { values: trimmed.map((value) => (typeof value === "number" ? value : Number.NaN)), numeric: true };
  }
  return { values: trimmed.map((value) => (value === null ? "" : String(value))), numeric: false };
}

/** Read `<c:numCache>`, `<c:strCache>`, `<c:numLit>`, `<c:strLit>` or a multi-level category cache. */
function readCachedRange(node: XmlNode | null): CachedRange | null {
  if (!node) return null;
  const cache =
    child(node, "numCache") ??
    child(node, "strCache") ??
    child(node, "numLit") ??
    child(node, "strLit") ??
    // Some producers nest the cache one level deeper.
    find(node, "numCache") ??
    find(node, "strCache");
  if (cache) {
    return readPoints(cache, cache.local === "numCache" || cache.local === "numLit");
  }

  // PowerPoint writes categories that reference more than one worksheet column as
  // a multi-level cache: `<c:multiLvlStrCache>` holding one `<c:lvl>` per level,
  // innermost first, each repeated for every point. Without this the categories
  // came back empty, which collapsed every bar into a single band and dropped the
  // category labels entirely. The view renders one label row, so the innermost
  // populated level is the one to draw.
  const multi = child(node, "multiLvlStrCache") ?? find(node, "multiLvlStrCache");
  if (!multi) return null;
  const levels = children(multi, "lvl")
    .map((level) => readPoints(level, false))
    .filter((range): range is CachedRange => range !== null);
  return levels.find((range) => range.values.some((value) => value !== "")) ?? levels[0] ?? null;
}

function readSeriesName(ser: XmlNode, ctx: TextContext): string | null {
  const tx = child(ser, "tx");
  if (!tx) return null;
  const cached = readCachedRange(tx);
  if (cached && cached.values.length > 0) {
    const first = cached.values[0];
    if (typeof first === "string" && first.length > 0) return first;
  }
  const rich = find(tx, "rich");
  if (rich) {
    const body = parseTextBody(rich, ctx);
    const text = textBodyPlainText(body);
    if (text.trim().length > 0) return text.trim();
  }
  return null;
}

function readTitle(title: XmlNode | null, ctx: TextContext): string | null {
  if (!title) return null;
  const tx = child(title, "tx");
  const cached = readCachedRange(tx);
  if (cached && cached.values.length > 0) {
    const text = cached.values.map((value) => String(value)).join(" ").trim();
    if (text.length > 0) return text;
  }
  const rich = tx ? find(tx, "rich") : find(title, "rich");
  if (rich) {
    const body = parseTextBody(rich, ctx);
    const text = textBodyPlainText(body);
    if (text.trim().length > 0) return text.trim();
  }
  return null;
}

function numericValues(values: (string | number)[]): (number | null)[] {
  return values.map((value) => {
    const parsed = typeof value === "number" ? value : Number(value);
    return Number.isFinite(parsed) ? parsed : null;
  });
}

export interface ChartParseContext extends DrawingContext {
  textContext: TextContext;
  /** ordered palette used when a series does not declare its own fill */
  palette: RGBA[];
}

export function parseChart(chartSpace: XmlNode | null, ctx: ChartParseContext): ChartData | null {
  if (!chartSpace) return null;
  const chart = child(chartSpace, "chart");
  if (!chart) return null;
  const plotArea = child(chart, "plotArea");
  if (!plotArea) return null;

  let group: XmlNode | null = null;
  for (const element of CHART_GROUP_ELEMENTS) {
    const found = child(plotArea, element);
    if (found) {
      group = found;
      break;
    }
  }
  if (!group) return null;

  const type = group.local;
  const is3d = type.includes("3D");
  const isDoughnut = type.startsWith("doughnut") || type.startsWith("ofPie");
  const barDir = (attr(child(group, "barDir"), "val") ?? "col") === "bar" ? "bar" : "col";
  const grouping = attr(child(group, "grouping"), "val");
  const overlapRaw = attrNum(child(group, "overlap"), "val");
  const gapRaw = attrNum(child(group, "gapWidth"), "val");
  const holeRaw = attrNum(child(group, "holeSize"), "val");
  const varyColors = attrBool(child(group, "varyColors"), "val") ??
    (type === "pieChart" || type === "pie3DChart" || type === "doughnutChart");

  const series: ChartSeries[] = [];
  let index = 0;
  for (const ser of children(group, "ser")) {
    const name = readSeriesName(ser, ctx.textContext);
    const spPr = child(ser, "spPr");
    const fill: Fill | null = spPr ? parseFillFromProps(spPr, ctx) : null;
    const line: LineStyle | null = spPr ? parseLine(spPr, ctx) : null;
    const color: RGBA | null =
      fill && fill.type === "solid"
        ? fill.color
        : fill && fill.type === "gradient" && fill.stops[0]
          ? fill.stops[0].color
          : (ctx.palette[index % Math.max(ctx.palette.length, 1)] ?? null);

    const pointColors: (RGBA | null)[] = [];
    for (const dPt of children(ser, "dPt")) {
      const pointIndex = attrNum(child(dPt, "idx"), "val");
      if (pointIndex === null || pointIndex < 0) continue;
      const props = child(dPt, "spPr");
      const pointFill = props ? parseFillFromProps(props, ctx) : null;
      while (pointColors.length <= pointIndex) pointColors.push(null);
      pointColors[pointIndex] = pointFill && pointFill.type === "solid" ? pointFill.color : null;
    }

    const catNode = child(ser, "cat") ?? child(ser, "xVal");
    const valNode = child(ser, "val") ?? child(ser, "yVal");
    const categories = readCachedRange(catNode);
    const values = readCachedRange(valNode);

    const seriesValues = values ? numericValues(values.values) : [];
    if (varyColors) {
      for (let point = 0; point < seriesValues.length; point += 1) {
        pointColors[point] ??= ctx.palette[point % Math.max(ctx.palette.length, 1)] ?? null;
      }
    }
    const seriesCategories = categories
      ? categories.numeric
        ? numericValues(categories.values).map((v) => v ?? "")
        : categories.values
      : null;
    const xValues =
      child(ser, "xVal") && categories ? numericValues(categories.values) : null;

    series.push({
      name,
      values: seriesValues,
      categories: seriesCategories,
      xValues,
      color,
      fill,
      line,
      pointColors: pointColors.length > 0 ? pointColors : null,
    });
    index += 1;
  }

  const categoriesFromSeries = series.find((s) => s.categories && s.categories.length > 0)?.categories;
  const categories: (string | number)[] = categoriesFromSeries ? [...categoriesFromSeries] : [];

  const legend = child(chart, "legend");
  const showLegend = legend !== null;
  const legendPos = attr(child(legend, "legendPos"), "val") ?? "r";
  const legendOverlay = attrBool(child(legend, "overlay"), "val") === true;

  const plots = children(plotArea, "dLbls");
  const groupLabels = child(group, "dLbls");
  // PowerPoint also writes label settings on the series (and even on single
  // points), so falling back to the first series keeps `showVal`/`showPercent`.
  const firstSeries = children(group, "ser")[0];
  const seriesLabels = firstSeries ? child(firstSeries, "dLbls") : null;
  const labelSource = groupLabels ?? seriesLabels ?? plots[0] ?? null;
  const showValueLabels = attrBool(child(labelSource, "showVal"), "val") === true;
  const showPercent = attrBool(child(labelSource, "showPercent"), "val") === true;

  const valueAxis = child(plotArea, "valAx");
  const categoryAxis = child(plotArea, "catAx") ?? child(plotArea, "dateAx");
  const style = attr(child(chartSpace, "style"), "val");

  return {
    type,
    title: readTitle(child(chart, "title"), ctx.textContext),
    series,
    categories,
    showLegend,
    legendPos,
    legendOverlay,
    barDir,
    grouping,
    overlap: overlapRaw,
    gapWidth: gapRaw,
    // `c:holeSize` only applies to doughnut charts; a pie must stay solid, so the
    // 50% default is applied per chart type rather than to every pie-like chart.
    holeSize: isDoughnut ? (holeRaw === null ? 50 : clamp(holeRaw, 0, 90)) : 0,
    showValueLabels,
    showPercent,
    valueAxisTitle: readTitle(child(valueAxis, "title"), ctx.textContext),
    categoryAxisTitle: readTitle(child(categoryAxis, "title"), ctx.textContext),
    style,
  };
}

function parseFillFromProps(props: XmlNode, ctx: DrawingContext): Fill | null {
  for (const candidate of props.children) {
    if (
      candidate.local === "solidFill" ||
      candidate.local === "noFill" ||
      candidate.local === "gradFill" ||
      candidate.local === "pattFill"
    ) {
      return parseFill(candidate, ctx);
    }
  }
  return null;
}

/** Series colours are resolved lazily in the parser but declared here for reuse. */
export function seriesColorFromFont(
  node: XmlNode | null,
  ctx: ChartParseContext,
  fallback: RGBA,
): RGBA {
  const colorNode = node ? node.children.find(isColorElement) : undefined;
  if (!colorNode) return fallback;
  return resolveColorElement(colorNode, { theme: ctx.theme, colorMap: ctx.colorMap, fallback });
}

export function chartSeriesCount(chart: ChartData): number {
  return chart.series.length;
}

export function chartHasData(chart: ChartData): boolean {
  return chart.series.some((series) => series.values.some((value) => value !== null));
}

export function chartAllCachedPoints(node: XmlNode | null): number {
  return findAll(node, "pt").length;
}
