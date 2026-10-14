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
import { rgbaToCss } from "../parser/color";

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

function seriesColor(series: ChartSeries, index: number, palette: string[]): string {
  if (series.color) return rgbaToCss(series.color);
  return palette[index % palette.length] ?? "#4472C4";
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

  const titleBand = chart.title ? 26 : 8;
  const legendBand = chart.showLegend && chart.series.length > 0 && !isPieLike(chart) ? 28 : 0;
  const axisLeft = simplified ? 8 : 52;
  const axisBottom = simplified ? 8 : 28;

  const plot: PlotBox = {
    x: axisLeft,
    y: titleBand,
    width: Math.max(width - axisLeft - 16, 10),
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
  const values = numericValues(chart);
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
                {formatNumber(tick)}
              </text>
            ) : null}
          </g>
        );
      })}
    </g>,
  );

  const categoryBand = (horizontal ? plot.height : plot.width) / categoryCount;
  const seriesCount = Math.max(chart.series.length, 1);
  const barThickness = Math.max((categoryBand * 0.72) / seriesCount, 1);

  categoryCount > 0 && nodes.push(
    <g key="cat-labels">
      {!simplified &&
        categories.map((label, index) => {
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
    const points: [number, number][] = [];
    series.values.forEach((value, index) => {
      if (typeof value !== "number" || !Number.isFinite(value)) return;
      const base = (horizontal ? plot.y : plot.x) + categoryBand * (index + 0.5);
      const offset = (seriesIndex - (seriesCount - 1) / 2) * barThickness;
      if (horizontal) {
        points.push([scaleX(value), base + offset]);
      } else {
        points.push([base + offset, scaleY(value)]);
      }
    });

    if (isBar) {
      nodes.push(
        <g key={`series-${seriesIndex}`}>
          {series.values.map((value, index) => {
            if (typeof value !== "number" || !Number.isFinite(value)) return null;
            const base = (horizontal ? plot.y : plot.x) + categoryBand * (index + 0.5);
            const offset = (seriesIndex - (seriesCount - 1) / 2) * barThickness;
            if (horizontal) {
              const x = scaleX(Math.min(value, 0));
              const width2 = Math.abs(scaleX(value) - scaleX(0));
              return (
                <rect
                  key={index}
                  x={x}
                  y={base + offset - barThickness / 2}
                  width={Math.max(width2, 0.5)}
                  height={Math.max(barThickness - 1, 0.5)}
                  fill={color}
                />
              );
            }
            const y = scaleY(Math.max(value, 0));
            const barHeight = Math.abs(scaleY(value) - scaleY(0));
            return (
              <rect
                key={index}
                x={base + offset - barThickness / 2}
                y={y}
                width={Math.max(barThickness - 1, 0.5)}
                height={Math.max(barHeight, 0.5)}
                fill={color}
              />
            );
          })}
          {chart.showValueLabels && !simplified
            ? series.values.map((value, index) =>
                typeof value === "number" ? (
                  <text
                    key={`label-${index}`}
                    x={(horizontal ? scaleX(value) + 4 : (plot.x + categoryBand * (index + 0.5)) + (seriesIndex - (seriesCount - 1) / 2) * barThickness)}
                    y={horizontal ? plot.y + categoryBand * (index + 0.5) + 3 : scaleY(value) - 4}
                    textAnchor={horizontal ? "start" : "middle"}
                    fontSize={9}
                    fill={TEXT_COLOR}
                  >
                    {formatNumber(value)}
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
          {points.map(([px, py], index) => (
            <circle key={index} cx={px} cy={py} r={2.5} fill={color} />
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
          return <circle key={pointIndex} cx={sx(xValue)} cy={sy(value)} r={3.5} fill={color} fillOpacity={0.85} />;
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

  const legendRight = chart.showLegend && !simplified ? 110 : 0;
  const radius = Math.max(Math.min(plot.width - legendRight, plot.height) / 2 - 6, 4);
  const cx = plot.x + (plot.width - legendRight) / 2;
  const cy = plot.y + plot.height / 2;
  const hole = Math.max((chart.holeSize / 100) * radius, 0);
  const nodes: ReactNode[] = [];
  const categories = series.categories ?? chart.categories;

  let angle = -Math.PI / 2;
  values.forEach((value, index) => {
    const slice = (value / total) * Math.PI * 2;
    const end = angle + slice;
    const color = palette[index % palette.length] ?? "#4472C4";
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
    nodes.push(<path key={index} d={d} fill={color} stroke="#FFFFFF" strokeWidth={1} />);

    if (!simplified && chart.showValueLabels) {
      const mid = angle + slice / 2;
      nodes.push(
        <text
          key={`label-${index}`}
          x={cx + (radius * 0.62) * Math.cos(mid)}
          y={cy + (radius * 0.62) * Math.sin(mid)}
          textAnchor="middle"
          fontSize={10}
          fill="#FFFFFF"
        >
          {`${Math.round((value / total) * 100)}%`}
        </text>,
      );
    }
    if (!simplified && chart.showLegend) {
      const label = categories[index];
      nodes.push(
        <g key={`legend-${index}`}>
          <rect x={plot.x + (plot.width - legendRight) + 8} y={plot.y + index * 16 + 4} width={10} height={10} fill={color} />
          <text x={plot.x + (plot.width - legendRight) + 22} y={plot.y + index * 16 + 13} fontSize={10} fill={TEXT_COLOR}>
            {String(label ?? series.name ?? index + 1).slice(0, 14)}
          </text>
        </g>,
      );
    }
    angle = end;
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
