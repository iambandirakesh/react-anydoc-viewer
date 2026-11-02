import type { ShapeNodeUnion, ThemeData, Transform } from "../../model/types";
import { readBinary, readText, sniffImageKind, type Archive } from "../archive";
import { chartHasData, parseChart } from "../chart";
import { accentPalette } from "../drawing";
import { relationshipsFor } from "../relationships";
import { parseShapeTree, type PartResolver } from "../shapes";
import { DEFAULT_COLOR_MAP } from "../theme";
import { attrNum, child, find, parseXml, relAttr, type XmlNode } from "../xml";
import type { SheetPageData } from "./model";

import { createChartReferenceResolver } from "./chartReferences";

const PX_TO_EMU = 9525;
const SUPPORTED_CHARTS = new Set(["barChart", "bar3DChart", "lineChart", "line3DChart", "areaChart", "area3DChart", "pieChart", "pie3DChart", "doughnutChart", "scatterChart"]);

/** Follow worksheet -> drawing -> chart/image relationships; external assets stay blocked. */
export function readSheetDrawings(archive: Archive, sheetPart: string, root: XmlNode, data: SheetPageData, theme: ThemeData, warn: (code: string, message: string) => void): ShapeNodeUnion[] {
  const nodes: ShapeNodeUnion[] = [];
  const sheetRelationships = relationshipsFor(archive, sheetPart);
  const drawingId = relAttr(child(root, "drawing"), "id");
  if (!drawingId) return nodes;
  const drawingRel = sheetRelationships.get(drawingId);
  if (!drawingRel || drawingRel.external) { warn("xlsx.drawing.relationship", "Worksheet drawing is missing or external and was skipped."); return nodes; }
  try {
    const resolveChartReferences = createChartReferenceResolver(archive);
    const drawingRoot = parseXml(readText(archive, drawingRel.target) ?? "");
    const rels = relationshipsFor(archive, drawingRel.target);
    const context = { theme, colorMap: DEFAULT_COLOR_MAP, warn, resolveRel: (id: string) => rels.get(id) ?? null };
    const resolver: PartResolver = {
      drawing: () => context,
      text: () => context,
      placeholderTransform: () => null,
      chart: () => null,
      mediaKind: (part) => { const bytes = readBinary(archive, part); return bytes ? sniffImageKind(bytes) : null; },
      warn,
    };
    const widthAt = (index: number) => {
      if (data.widths[index] !== undefined) return data.widths[index]!;
      let width = Math.min(Math.max(Math.floor((attrNum(child(root, "sheetFormatPr"), "defaultColWidth") ?? 8.43) * 7 + 5), 1), 2048);
      for (const col of child(root, "cols")?.children ?? []) {
        if (index + 1 >= (attrNum(col, "min") ?? 1) && index + 1 <= (attrNum(col, "max") ?? 1)) {
          width = col.attrs.hidden === "1" ? 0 : Math.floor((attrNum(col, "width") ?? 8.43) * 7 + 5);
        }
      }
      return width;
    };
    const rowNodes = new Map((child(root, "sheetData")?.children ?? []).map((row) => [(attrNum(row, "r") ?? 0) - 1, row]));
    const heightAt = (index: number) => {
      if (data.heights?.[index] !== undefined) return data.heights[index]!;
      const row = rowNodes.get(index);
      return row?.attrs.hidden === "1" ? 0 : (attrNum(row, "ht") ?? attrNum(child(root, "sheetFormatPr"), "defaultRowHeight") ?? 15) * 96 / 72;
    };
    const columnOffsets = [0], rowOffsets = [0];
    for (let index = 0; index < 200; index++) columnOffsets.push(columnOffsets[index]! + widthAt(index) * PX_TO_EMU);
    for (let index = 0; index < 5000; index++) rowOffsets.push(rowOffsets[index]! + heightAt(index) * PX_TO_EMU);
    const marker = (node: XmlNode | null) => {
      const col = Math.max(0, Math.min(200, Number(child(node, "col")?.text ?? 0)));
      const row = Math.max(0, Math.min(5000, Number(child(node, "row")?.text ?? 0)));
      return { x: (columnOffsets[col] ?? NaN) + Number(child(node, "colOff")?.text ?? 0), y: (rowOffsets[row] ?? NaN) + Number(child(node, "rowOff")?.text ?? 0) };
    };
    for (const [index, anchor] of (drawingRoot?.children ?? []).slice(0, 200).entries()) {
      try {
      if (!["twoCellAnchor", "oneCellAnchor", "absoluteAnchor"].includes(anchor.local)) continue;
      const from = anchor.local === "absoluteAnchor" ? { x: attrNum(child(anchor, "pos"), "x") ?? 0, y: attrNum(child(anchor, "pos"), "y") ?? 0 } : marker(child(anchor, "from"));
      const to = child(anchor, "to") ? marker(child(anchor, "to")) : null;
      const xfrm: Transform = { ...from, cx: to ? to.x - from.x : attrNum(child(anchor, "ext"), "cx") ?? 0, cy: to ? to.y - from.y : attrNum(child(anchor, "ext"), "cy") ?? 0, rot: 0, flipH: false, flipV: false };
      if (![xfrm.x, xfrm.y, xfrm.cx, xfrm.cy].every(Number.isFinite) || xfrm.cx <= 0 || xfrm.cy <= 0 || xfrm.cx > 100000000 || xfrm.cy > 100000000) { warn("xlsx.drawing.anchor", "Invalid or oversized drawing anchor was skipped."); continue; }
      const chartRef = find(anchor, "chart");
      const pic = child(anchor, "pic");
      if (chartRef) {
        const rel = rels.get(relAttr(chartRef, "id") ?? "");
        if (!rel || rel.external) { warn("xlsx.chart.relationship", "Chart reference is missing or external."); continue; }
        const chartRoot = parseXml(readText(archive, rel.target) ?? "");
        resolveChartReferences(chartRoot);
        const chart = parseChart(chartRoot, { ...context, textContext: context, palette: accentPalette(theme, DEFAULT_COLOR_MAP) });
        if (!chart || !SUPPORTED_CHARTS.has(chart.type) || !chartHasData(chart)) { warn("xlsx.chart.unsupported", "An embedded chart has no supported cached plot data."); continue; }
        nodes.push({ kind: "chart", id: `xlsx-chart-${index}`, name: null, xfrm, chart });
      } else if (pic) {
        const container: XmlNode = { name: "spTree", local: "spTree", attrs: {}, text: "", children: [pic] };
        for (const node of parseShapeTree(container, drawingRel.target, resolver)) nodes.push({ ...node, xfrm: { ...node.xfrm, ...from, cx: xfrm.cx, cy: xfrm.cy } });
      } else warn("xlsx.drawing.unsupported", "A worksheet drawing object is not yet supported.");
      } catch {
        warn("xlsx.drawing.unreadable", `Worksheet drawing ${index + 1} could not be read; other drawings are retained.`);
      }
    }
    if ((drawingRoot?.children.length ?? 0) > 200) warn("xlsx.drawing.limit", "Only the first 200 drawing anchors are shown.");
  } catch {
    warn("xlsx.drawing.unreadable", "A worksheet drawing or chart part could not be read.");
  }
  return nodes;
}
