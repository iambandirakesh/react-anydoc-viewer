/**
 * Chart rendering.
 *
 * Charts are drawn from the cached plot data stored inside the presentation, so
 * no spreadsheet evaluation is ever required. This is a deliberately compact
 * renderer: it covers the common families (bar/column, line, area, pie,
 * doughnut, scatter) with axes, gridlines, a legend and optional value labels,
 * and reports anything it cannot draw rather than drawing something wrong.
 */

import type { ReactNode } from "react";
import type { ChartData, ChartNode, ChartSeries } from "../model/types";
import { emuToPx } from "../parser/units";
import { rgbaToCss } from "../parser/drawingml/color";

export const DEFAULT_CHART_PALETTE = [
  "#4472C4",
  "#ED7D31",
  "#A5A5A5",
  "#FFC000",
  "#5B9BD5",
  "#70AD47",
  "#264478",
  "#9E480E",
  "#636363",
  "#997300",
  "#255E91",
  "#43682B",
];

const AXIS_COLOR = "#D0D0D0";
const TEXT_COLOR = "#404040";

export interface ChartRenderOptions {
  palette?: string[];
  /** renders without axis labels for tiny thumbnails */
  simplified?: boolean;
}

interface PlotBox {
  x: number;
  y: number;
  width: number;
  height: number;
}

function niceTicks(min: number, max: number, count = 5): number[] {
  if (!Number.isFinite(min) || !Number.isFinite(max)) return [0, 1];
  if (min === max) return [min - 1, max + 1];
  const range = max - min;
  const rough = range / count;
  const magnitude = 10 ** Math.floor(Math.log10(rough));
  const normalized = rough / magnitude;
  const step = (normalized >= 5 ? 5 : normalized >= 2 ? 2 : 1) * magnitude;
  const start = Math.floor(min / step) * step;
  const end = Math.ceil(max / step) * step;
  const ticks: number[] = [];
  for (let value = start; value <= end + step / 2; value += step) {
    ticks.push(Math.round(value * 1e6) / 1e6);
  }
  return ticks;
}

function formatNumber(value: number): string {
  if (Math.abs(value) >= 1_000_000) return `${(value / 1_000_000).toFixed(1)}M`;
  if (Math.abs(value) >= 1000) return `${(value / 1000).toFixed(1)}k`;
  if (Number.isInteger(value)) return String(value);
  return value.toFixed(2).replace(/\.?0+$/, "");
}

const exactNumber = new Intl.NumberFormat("en-US", { maximumFractionDigits: 20 });

function pointDetails(series: ChartSeries, category: string | number | undefined, value: number): string {
  return [category, series.name, exactNumber.format(value)].filter((part) => part !== null && part !== undefined && part !== "").join(" · ");
}

function seriesColor(series: ChartSeries, index: number, palette: string[]): string {
  if (series.color) return rgbaToCss(series.color);
  return palette[index % palette.length] ?? "#4472C4";
}

/**
 * A point's own `<c:dPt>` colour wins, then the series colour, then the palette.
 * PowerPoint stores per-bar and per-slice colours this way, so ignoring it painted
 * a whole column chart in the first series' colour.
 */
function pointColor(
  series: ChartSeries,
  pointIndex: number,
  seriesIndex: number,
  palette: string[],
): string {
  const explicit = series.pointColors?.[pointIndex];
  if (explicit) return rgbaToCss(explicit);
  return seriesColor(series, seriesIndex, palette);
}

function numericValues(chart: ChartData): number[] {
  const out: number[] = [];
  for (const series of chart.series) {
    for (const value of series.values) {
      if (typeof value === "number" && Number.isFinite(value)) out.push(value);
    }
  }
  return out;
}

function isPieLike(chart: ChartData): boolean {
  return chart.type === "pieChart" || chart.type === "doughnutChart" || chart.type === "pie3DChart";
}

export function renderChart(node: ChartNode, options: ChartRenderOptions = {}): ReactNode {
  const width = emuToPx(node.xfrm.cx);
  const height = emuToPx(node.xfrm.cy);
  const chart = node.chart;
  const palette = options.palette ?? DEFAULT_CHART_PALETTE;
  const simplified = options.simplified === true;

  if (width <= 0 || height <= 0) return null;

  // Outside-end value labels need room for their ascent above the tallest bar.
  const valueLabelBand = !simplified && chart.showValueLabels && chart.type.startsWith("bar") && chart.barDir !== "bar" ? 14 : 0;
  const titleBand = (chart.title ? 26 : 8) + valueLabelBand;
  const legendBand = chart.showLegend && chart.series.length > 0 && !isPieLike(chart) ? 28 : 0;
  // A horizontal bar chart prints its category names in the left gutter, so the
  // gutter has to fit the longest one instead of clipping "Office Supplies".
  const horizontalBars = chart.barDir === "bar" && chart.type.startsWith("bar");
  const categoryLabels = chart.categories.length > 0 ? chart.categories : (chart.series[0]?.categories ?? []);
  const longestLabel = categoryLabels.reduce<number>(
    (max, label) => Math.max(max, String(label).length),
    0,
  );
  const axisLeft = simplified
    ? 8
    : horizontalBars
      ? Math.min(Math.max(20 + Math.min(longestLabel, 26) * 5.4, 52), 180)
      : 52;
  const axisBottom = simplified ? 8 : 28;
  const axisRight = horizontalBars && chart.showValueLabels && !simplified
    ? Math.max(16, Math.max(0, ...numericValues(chart).map((value) => exactNumber.format(value).length)) * 5.4 + 8)
    : 16;

  const plot: PlotBox = {
    x: axisLeft,
    y: titleBand,
    width: Math.max(width - axisLeft - axisRight, 10),
    height: Math.max(height - titleBand - axisBottom - legendBand, 10),
  };

  const nodes: ReactNode[] = [];

  if (chart.title && !simplified) {
    nodes.push(
      <text
        key="title"
        x={width / 2}
        y={18}
        textAnchor="middle"
        fill={TEXT_COLOR}
        fontSize={15}
        fontWeight={600}
      >
        {chart.title}
      </text>,
    );
  }

  if (isPieLike(chart)) {
    nodes.push(...pieNodes(chart, plot, palette, simplified));
  } else if (chart.type === "scatterChart") {
    nodes.push(...scatterNodes(chart, plot, palette, simplified));
  } else {
    nodes.push(...cartesianNodes(chart, plot, palette, simplified));
  }

  if (chart.showLegend && !simplified && chart.series.length > 0) {
    nodes.push(...legendNodes(chart, palette, width, height, isPieLike(chart)));
  }

  return (
    <svg
      viewBox={`0 0 ${width} ${height}`}
      width="100%"
      height="100%"
      preserveAspectRatio="none"
      role="img"
      aria-label={chart.title ?? `${chart.type} chart`}
      style={{ position: "absolute", inset: 0, overflow: "hidden" }}
    >
      {nodes}
    </svg>
  );
}

function cartesianNodes(
  chart: ChartData,
  plot: PlotBox,
  palette: string[],
  simplified: boolean,
): ReactNode[] {
  const nodes: ReactNode[] = [];
  const stacked = chart.type.startsWith("bar") && (chart.grouping === "stacked" || chart.grouping === "percentStacked");
  const percentStacked = stacked && chart.grouping === "percentStacked";
  const positive: number[] = [];
  const negative: number[] = [];
  const totals: number[] = [];
  if (percentStacked) {
    for (const series of chart.series) series.values.forEach((value, index) => {
      if (typeof value === "number" && Number.isFinite(value)) totals[index] = (totals[index] ?? 0) + Math.abs(value);
    });
  }
  const ranges = chart.series.map((series) => series.values.map((value, index): [number, number] | null => {
    if (typeof value !== "number" || !Number.isFinite(value)) return null;
    const amount = percentStacked ? (totals[index] ? value / totals[index]! * 100 : 0) : value;
    const accumulated = amount < 0 ? negative : positive;
    const start = stacked ? accumulated[index] ?? 0 : 0;
    accumulated[index] = start + amount;
    return [start, start + amount];
  }));
  const values = stacked ? ranges.flatMap((series) => series.flatMap((range) => range ?? [])) : numericValues(chart);
  let min = Math.min(0, ...values);
  let max = Math.max(0, ...values);
  if (min === max) {
    min = 0;
    max = max === 0 ? 1 : max * 1.2;
  }
  const ticks = simplified ? [min, max] : niceTicks(min, max, 5);
  const domainMin = ticks[0] ?? 0;
  const domainMax = ticks[ticks.length - 1] ?? 1;
  const scaleY = (value: number) =>
    plot.y + plot.height - ((value - domainMin) / (domainMax - domainMin || 1)) * plot.height;
  const scaleX = (value: number) =>
    plot.x + ((value - domainMin) / (domainMax - domainMin || 1)) * plot.width;

  const horizontal = chart.barDir === "bar" && chart.type.startsWith("bar");
  const isBar = chart.type.startsWith("bar");
  const categories = chart.categories.length > 0
    ? chart.categories
    : chart.series[0]?.categories ?? [];
  const categoryCount = Math.max(categories.length, 1);

  // Gridlines, ticks and axis labels.
  nodes.push(
    <g key="axes">
      <line
        x1={plot.x}
        y1={plot.y + plot.height}
        x2={plot.x + plot.width}
        y2={plot.y + plot.height}
        stroke={AXIS_COLOR}
      />
      <line x1={plot.x} y1={plot.y} x2={plot.x} y2={plot.y + plot.height} stroke={AXIS_COLOR} />
      {ticks.map((tick) => {
        const position = horizontal ? scaleX(tick) : scaleY(tick);
        return (
          <g key={`tick-${tick}`}>
            <line
              x1={horizontal ? position : plot.x}
              y1={horizontal ? plot.y : position}
              x2={horizontal ? position : plot.x + plot.width}
              y2={horizontal ? plot.y + plot.height : position}
              stroke={AXIS_COLOR}
              strokeDasharray="2 3"
            />
            {!simplified ? (
              <text
                x={horizontal ? position : plot.x - 6}
                y={horizontal ? plot.y + plot.height + 14 : position + 3}
                textAnchor={horizontal ? "middle" : "end"}
                fontSize={10}
                fill={TEXT_COLOR}
              >
                {percentStacked ? `${formatNumber(tick)}%` : formatNumber(tick)}
              </text>
            ) : null}
          </g>
        );
      })}
    </g>,
  );

  const categoryBand = (horizontal ? plot.height : plot.width) / categoryCount;
  const labelWidth = Math.max(0, ...categories.map((label) => Math.min(String(label).length, 18) * 5.8 + 12));
  const labelStride = horizontal ? 1 : Math.max(1, Math.ceil(labelWidth / categoryBand));
  const seriesCount = stacked ? 1 : Math.max(chart.series.length, 1);
  const barThickness = Math.max((categoryBand * 0.72) / seriesCount, 1);
  const visibleSeries = categories.map((_, index) => chart.series.flatMap((series, seriesIndex) => {
    const value = series.values[index];
    return typeof value === "number" && Number.isFinite(value) && value !== 0 ? [seriesIndex] : [];
  }));
  const barOffset = (index: number, seriesIndex: number): number => {
    if (stacked) return 0;
    const active = visibleSeries[index] ?? [];
    const position = active.indexOf(seriesIndex);
    return position < 0 ? 0 : (position - (active.length - 1) / 2) * barThickness;
  };

  categoryCount > 0 && nodes.push(
    <g key="cat-labels">
      {!simplified &&
        categories.map((label, index) => {
          if (index % labelStride !== 0) return null;
          const center = (horizontal ? plot.y : plot.x) + categoryBand * (index + 0.5);
          return (
            <text
              key={`cat-${index}`}
              x={horizontal ? plot.x - 6 : center}
              y={horizontal ? center + 3 : plot.y + plot.height + 14}
              textAnchor={horizontal ? "end" : "middle"}
              fontSize={10}
              fill={TEXT_COLOR}
            >
              {String(label).slice(0, 18)}
            </text>
          );
        })}
    </g>,
  );

  chart.series.forEach((series, seriesIndex) => {
    const color = seriesColor(series, seriesIndex, palette);
    const points: [number, number, number][] = [];
    series.values.forEach((value, index) => {
      if (typeof value !== "number" || !Number.isFinite(value)) return;
      const base = (horizontal ? plot.y : plot.x) + categoryBand * (index + 0.5);
      const offset = (seriesIndex - (seriesCount - 1) / 2) * barThickness;
      if (horizontal) {
        points.push([scaleX(value), base + offset, index]);
      } else {
        points.push([base + offset, scaleY(value), index]);
      }
    });

    if (isBar) {
      nodes.push(
        <g key={`series-${seriesIndex}`} data-chart-series={seriesIndex}>
          {series.values.map((value, index) => {
            if (typeof value !== "number" || !Number.isFinite(value)) return null;
            // A zero-length bar is a hairline artefact on the baseline; PowerPoint
            // draws nothing at all for a zero value.
            if (value === 0) return null;
            const base = (horizontal ? plot.y : plot.x) + categoryBand * (index + 0.5);
            const offset = barOffset(index, seriesIndex);
            const [start, end] = ranges[seriesIndex]?.[index] ?? [0, value];
            if (horizontal) {
              const x = scaleX(Math.min(start, end));
              const width2 = Math.abs(scaleX(end) - scaleX(start));
              return (
                <rect
                  key={index}
                  x={x}
                  y={base + offset - Math.max(barThickness - 1, 0.5) / 2}
                  width={Math.max(width2, 0.5)}
                  height={Math.max(barThickness - 1, 0.5)}
                  fill={pointColor(series, index, seriesIndex, palette)}
                >
                  <title>{pointDetails(series, categories[index], value)}</title>
                </rect>
              );
            }
            const y = scaleY(Math.max(start, end));
            const barHeight = Math.abs(scaleY(end) - scaleY(start));
            return (
              <rect
                key={index}
                x={base + offset - Math.max(barThickness - 1, 0.5) / 2}
                y={y}
                width={Math.max(barThickness - 1, 0.5)}
                height={Math.max(barHeight, 0.5)}
                fill={pointColor(series, index, seriesIndex, palette)}
              >
                <title>{pointDetails(series, categories[index], value)}</title>
              </rect>
            );
          })}
          {chart.showValueLabels && !simplified
            ? series.values.map((value, index) =>
                typeof value === "number" && Number.isFinite(value) &&
                !(value === 0 && chart.series.length > 1 && (visibleSeries[index]?.length ?? 0) > 0) ? (
                  <text
                    key={`label-${index}`}
                    x={horizontal
                      ? stacked ? (scaleX(ranges[seriesIndex]![index]![0]) + scaleX(ranges[seriesIndex]![index]![1])) / 2 : scaleX(value) + 4
                      : plot.x + categoryBand * (index + 0.5) + barOffset(index, seriesIndex)}
                    y={horizontal ? plot.y + categoryBand * (index + 0.5) + barOffset(index, seriesIndex) + 3 : stacked
                      ? (scaleY(ranges[seriesIndex]![index]![0]) + scaleY(ranges[seriesIndex]![index]![1])) / 2 + 3
                      : scaleY(value) - 4}
                    textAnchor={horizontal && !stacked ? "start" : "middle"}
                    fontSize={9}
                    fill={TEXT_COLOR}
                  >
                    {exactNumber.format(value)}
                    <title>{pointDetails(series, categories[index], value)}</title>
                  </text>
                ) : null,
              )
            : null}
        </g>,
      );
      return;
    }

    // Line / area series.
    const path = points.map(([px, py], index) => `${index === 0 ? "M" : "L"}${px.toFixed(2)},${py.toFixed(2)}`).join(" ");
    if (chart.type.startsWith("area") && points.length > 1) {
      const first = points[0]!;
      const last = points[points.length - 1]!;
      const baseline = scaleY(0);
      nodes.push(
        <path
          key={`area-${seriesIndex}`}
          d={`${path}L${last[0].toFixed(2)},${baseline.toFixed(2)}L${first[0].toFixed(2)},${baseline.toFixed(2)}Z`}
          fill={color}
          fillOpacity={0.55}
          stroke="none"
        />,
      );
    }
    if (path.length > 0) {
      nodes.push(
        <path
          key={`line-${seriesIndex}`}
          d={path}
          fill="none"
          stroke={color}
          strokeWidth={2}
          strokeLinejoin="round"
          strokeLinecap="round"
        />,
      );
    }
    if (points.length <= 60) {
      nodes.push(
        <g key={`markers-${seriesIndex}`}>
          {points.map(([px, py, index]) => (
            <circle key={index} cx={px} cy={py} r={2.5} fill={color}>
              <title>{pointDetails(series, categories[index], series.values[index]!)}</title>
            </circle>
          ))}
        </g>,
      );
    }
  });

  return nodes;
}

function scatterNodes(
  chart: ChartData,
  plot: PlotBox,
  palette: string[],
  simplified: boolean,
): ReactNode[] {
  const nodes: ReactNode[] = [];
  const xs: number[] = [];
  const ys: number[] = [];
  for (const series of chart.series) {
    for (const value of series.xValues ?? []) if (typeof value === "number") xs.push(value);
    for (const value of series.values) if (typeof value === "number") ys.push(value);
  }
  const xTicks = niceTicks(Math.min(0, ...xs), Math.max(0, ...xs), 4);
  const yTicks = niceTicks(Math.min(0, ...ys), Math.max(0, ...ys), 4);
  const xMin = xTicks[0] ?? 0;
  const xMax = xTicks[xTicks.length - 1] ?? 1;
  const yMin = yTicks[0] ?? 0;
  const yMax = yTicks[yTicks.length - 1] ?? 1;
  const sx = (value: number) => plot.x + ((value - xMin) / (xMax - xMin || 1)) * plot.width;
  const sy = (value: number) => plot.y + plot.height - ((value - yMin) / (yMax - yMin || 1)) * plot.height;

  nodes.push(
    <g key="scatter-axes">
      <line x1={plot.x} y1={plot.y + plot.height} x2={plot.x + plot.width} y2={plot.y + plot.height} stroke={AXIS_COLOR} />
      <line x1={plot.x} y1={plot.y} x2={plot.x} y2={plot.y + plot.height} stroke={AXIS_COLOR} />
    </g>,
  );

  chart.series.forEach((series, index) => {
    const color = seriesColor(series, index, palette);
    const xValues = series.xValues ?? series.categories?.map((value) => Number(value)) ?? [];
    nodes.push(
      <g key={`scatter-${index}`}>
        {series.values.map((value, pointIndex) => {
          const xValue = xValues[pointIndex];
          if (typeof value !== "number" || typeof xValue !== "number" || !Number.isFinite(xValue)) return null;
          return <circle key={pointIndex} cx={sx(xValue)} cy={sy(value)} r={3.5} fill={color} fillOpacity={0.85}>
            <title>{[series.name, `X: ${exactNumber.format(xValue)}`, `Y: ${exactNumber.format(value)}`].filter(Boolean).join(" · ")}</title>
          </circle>;
        })}
      </g>,
    );
  });

  if (!simplified) {
    nodes.push(
      <g key="scatter-labels">
        {yTicks.map((tick) => (
          <text key={tick} x={plot.x - 6} y={sy(tick) + 3} textAnchor="end" fontSize={10} fill={TEXT_COLOR}>
            {formatNumber(tick)}
          </text>
        ))}
        {xTicks.map((tick) => (
          <text key={tick} x={sx(tick)} y={plot.y + plot.height + 14} textAnchor="middle" fontSize={10} fill={TEXT_COLOR}>
            {formatNumber(tick)}
          </text>
        ))}
      </g>,
    );
  }
  return nodes;
}

/** Split a legend label into lines that fit `maxChars`, preferring word breaks. */
function wrapLabel(text: string, maxChars: number): string[] {
  const lines: string[] = [];
  let current = "";
  for (const word of text.split(/\s+/)) {
    const candidate = current.length === 0 ? word : `${current} ${word}`;
    if (current.length === 0 || candidate.length <= maxChars) {
      current = candidate;
    } else {
      lines.push(current);
      current = word;
    }
  }
  if (current.length > 0) lines.push(current);
  return lines.length > 0 ? lines.slice(0, 3) : [text.slice(0, maxChars)];
}

/** The text printed inside a pie or doughnut slice, matching the chart's own labels. */
function sliceLabel(value: number, total: number, chart: ChartData): string | null {
  if (chart.showPercent) return `${((value / total) * 100).toFixed(1)}%`;
  if (chart.showValueLabels) return exactNumber.format(value);
  return null;
}

function pieNodes(
  chart: ChartData,
  plot: PlotBox,
  palette: string[],
  simplified: boolean,
): ReactNode[] {
  const series = chart.series[0];
  if (!series) return [];
  const values = series.values.map((value) => (typeof value === "number" && Number.isFinite(value) ? Math.abs(value) : 0));
  const total = values.reduce((sum, value) => sum + value, 0);
  if (total <= 0) return [];

  const categories = series.categories ?? chart.categories;
  const legendWidth =
    chart.showLegend && !simplified ? Math.min(Math.max(plot.width * 0.42, 96), 260) : 0;
  const radius = Math.max(Math.min(plot.width - legendWidth, plot.height) / 2 - 6, 4);
  const cx = plot.x + (plot.width - legendWidth) / 2;
  const cy = plot.y + plot.height / 2;
  const hole = Math.max((chart.holeSize / 100) * radius, 0);
  const nodes: ReactNode[] = [];
  const maxChars = Math.max(8, Math.floor((legendWidth - 24) / 5.4));
  const legendLabels =
    chart.showLegend && !simplified
      ? values.map((_, index) => {
          const name = String(categories[index] ?? series.name ?? index + 1);
          return wrapLabel(name, maxChars);
        })
      : [];
  const legendTops: number[] = [];
  let legendCursor = plot.y + 2;
  for (const lines of legendLabels) {
    legendTops.push(legendCursor);
    legendCursor += lines.length * 12 + 4;
  }

  let angle = -Math.PI / 2;
  values.forEach((value, index) => {
    const slice = (value / total) * Math.PI * 2;
    const end = angle + slice;
    const color = pointColor(series, index, 0, palette);
    const largeArc = slice > Math.PI ? 1 : 0;
    const x1 = cx + radius * Math.cos(angle);
    const y1 = cy + radius * Math.sin(angle);
    const x2 = cx + radius * Math.cos(end);
    const y2 = cy + radius * Math.sin(end);
    let d: string;
    if (hole > 0) {
      const ix1 = cx + hole * Math.cos(end);
      const iy1 = cy + hole * Math.sin(end);
      const ix2 = cx + hole * Math.cos(angle);
      const iy2 = cy + hole * Math.sin(angle);
      d = `M${x1},${y1}A${radius},${radius} 0 ${largeArc} 1 ${x2},${y2}L${ix1},${iy1}A${hole},${hole} 0 ${largeArc} 0 ${ix2},${iy2}Z`;
    } else {
      d = `M${cx},${cy}L${x1},${y1}A${radius},${radius} 0 ${largeArc} 1 ${x2},${y2}Z`;
    }
    nodes.push(
      <path key={index} d={d} fill={color} stroke="#FFFFFF" strokeWidth={1}>
        <title>{`${pointDetails(series, categories[index], value)} · ${((value / total) * 100).toFixed(1)}%`}</title>
      </path>,
    );

    const label = sliceLabel(value, total, chart);
    if (!simplified && label) {
      const mid = angle + slice / 2;
      const labelRadius = hole > 0 ? (radius + hole) / 2 : radius * 0.62;
      nodes.push(
        <text
          key={`label-${index}`}
          x={cx + labelRadius * Math.cos(mid)}
          y={cy + labelRadius * Math.sin(mid)}
          textAnchor="middle"
          dominantBaseline="middle"
          fontSize={Math.min(Math.max(radius * 0.12, 8), 14)}
          fontWeight={600}
          fill="#FFFFFF"
        >
          {label}
        </text>,
      );
    }
    angle = end;
  });

  // Match each slice to its category using the legend's color swatch.
  legendLabels.forEach((lines, index) => {
    const top = legendTops[index] ?? plot.y;
    nodes.push(
      <g key={`legend-${index}`}>
        <rect
          x={plot.x + (plot.width - legendWidth) + 4}
          y={top}
          width={10}
          height={10}
          fill={pointColor(series, index, 0, palette)}
        />
        <text x={plot.x + (plot.width - legendWidth) + 18} y={top + 9} fontSize={10} fill={TEXT_COLOR}>
          {lines.map((line, lineIndex) => (
            <tspan key={lineIndex} x={plot.x + (plot.width - legendWidth) + 18} dy={lineIndex === 0 ? 0 : 12}>
              {/* keep the word break in the accessible text; a trailing space is invisible */}
              {lineIndex === lines.length - 1 ? line : `${line} `}
            </tspan>
          ))}
        </text>
      </g>,
    );
  });

  return nodes;
}

function legendNodes(
  chart: ChartData,
  palette: string[],
  width: number,
  height: number,
  pieLike: boolean,
): ReactNode[] {
  if (pieLike) return [];
  const items = chart.series.map((series, index) => ({
    label: series.name ?? `Series ${index + 1}`,
    color: seriesColor(series, index, palette),
  }));
  const itemWidth = Math.min(Math.max(width / Math.max(items.length, 1), 60), 140);
  const totalWidth = itemWidth * items.length;
  const startX = Math.max((width - totalWidth) / 2, 4);
  const y = height - 12;
  return items.map((item, index) => (
    <g key={`legend-item-${index}`}>
      <rect x={startX + index * itemWidth} y={y - 9} width={10} height={10} fill={item.color} />
      <text x={startX + index * itemWidth + 14} y={y} fontSize={11} fill={TEXT_COLOR}>
        {item.label.slice(0, 18)}
      </text>
    </g>
  ));
}
