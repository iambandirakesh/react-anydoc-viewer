/**
 * DrawingML table parsing (`a:tbl`).
 *
 * Table styles (`tableStyles.xml`) are a large, template-driven feature. We do
 * not implement them; instead, when a table declares no explicit cell borders we
 * apply a neutral grid so the table is still legible, and report that the style
 * was approximated.
 */

import type { Fill, RGBA, TableBorders, TableCell, TableData, TableRow } from "../model/types";
import { attr, attrBool, attrNum, child, children, type XmlNode } from "./xml";
import { emuToPx, ptToPx } from "./units";
import { isColorElement, resolveColorElement } from "./color";
import { parseFill, type DrawingContext } from "./drawing";
import { parseTextBody, type TextContext } from "./text";

const DEFAULT_CELL_INSETS = { l: 91_440, t: 45_720, r: 91_440, b: 45_720 };

function parseBorderEdge(node: XmlNode | null, ctx: DrawingContext): TableBorders[keyof TableBorders] {
  if (!node) return null;
  const width = attrNum(node, "w");
  const colorNode = node.children.find(isColorElement);
  const color: RGBA | null = colorNode
    ? resolveColorElement(colorNode, { theme: ctx.theme, colorMap: ctx.colorMap })
    : null;
  const dash = attr(child(node, "prstDash"), "val");
  if (width === 0 || child(node, "noFill")) {
    return { width: 0, color: null, dash: null };
  }
  return {
    width: width === null ? 12_700 : width,
    color,
    dash: dash ?? null,
  };
}

function parseCellBorders(tcPr: XmlNode | null, ctx: DrawingContext): TableBorders | null {
  if (!tcPr) return null;
  const borders: TableBorders = {};
  let any = false;
  for (const [key, local] of [["l", "lnL"], ["t", "lnT"], ["r", "lnR"], ["b", "lnB"]] as const) {
    const node = child(tcPr, local);
    if (!node) continue;
    borders[key] = parseBorderEdge(node, ctx);
    any = true;
  }
  return any ? borders : null;
}

function parseCellText(
  tc: XmlNode,
  ctx: TextContext,
): import("../model/types").TextBody | null {
  const txBody = child(tc, "txBody");
  if (!txBody) return null;
  return parseTextBody(txBody, ctx);
}

export interface TableParseOptions {
  textContext: TextContext;
  onApproximatedStyle?: () => void;
}

export function parseTable(tbl: XmlNode | null, options: TableParseOptions): TableData | null {
  if (!tbl) return null;
  const ctx = options.textContext;
  const tblPr = child(tbl, "tblPr");
  const styleId = attr(child(tblPr, "tableStyleId"), "val") ?? attr(tblPr, "tableStyleId");

  const grid = child(tbl, "tblGrid");
  const colWidths: number[] = [];
  for (const col of children(grid, "gridCol")) {
    colWidths.push(attrNum(col, "w") ?? 0);
  }

  const rows: TableRow[] = [];
  let explicitBorders = false;

  for (const tr of children(tbl, "tr")) {
    const height = attrNum(tr, "h") ?? 0;
    const cells: TableCell[] = [];
    for (const tc of children(tr, "tc")) {
      const tcPr = child(tc, "tcPr");
      const gridSpan = attrNum(tc, "gridSpan") ?? attrNum(tcPr, "gridSpan") ?? 1;
      const rowSpan = attrNum(tc, "rowSpan") ?? attrNum(tcPr, "rowSpan") ?? 1;
      const hMerge = attrBool(tc, "hMerge") === true;
      const vMerge = attrBool(tc, "vMerge") === true;
      const borders = parseCellBorders(tcPr, ctx);
      if (borders) explicitBorders = true;
      const fill: Fill | null = parseFillFromCell(tcPr, ctx);
      const insets = {
        l: emuToPx(attrNum(tcPr, "marL") ?? DEFAULT_CELL_INSETS.l),
        t: emuToPx(attrNum(tcPr, "marT") ?? DEFAULT_CELL_INSETS.t),
        r: emuToPx(attrNum(tcPr, "marR") ?? DEFAULT_CELL_INSETS.r),
        b: emuToPx(attrNum(tcPr, "marB") ?? DEFAULT_CELL_INSETS.b),
      };
      const anchorAttr = attr(tcPr, "anchor");
      const anchor =
        anchorAttr === "ctr" || anchorAttr === "b" || anchorAttr === "t" ? anchorAttr : null;
      cells.push({
        gridSpan: Math.max(1, Math.round(gridSpan)),
        rowSpan: Math.max(1, Math.round(rowSpan)),
        hMerge,
        vMerge,
        fill,
        borders,
        insets,
        anchor,
        text: parseCellText(tc, ctx),
        merged: hMerge || vMerge,
      });
    }
    rows.push({ height, cells });
  }

  if (colWidths.length === 0) {
    const maxCells = rows.reduce((max, row) => Math.max(max, row.cells.length), 0);
    for (let i = 0; i < maxCells; i += 1) colWidths.push(0);
    options.onApproximatedStyle?.();
  }

  if (!explicitBorders && rows.length > 0) {
    // Neutral grid so an unstyled table is still readable.
    const fallback: TableBorders = {
      t: { width: 12_700, color: { r: 0xd0, g: 0xd0, b: 0xd0, a: 1 }, dash: null },
      l: { width: 12_700, color: { r: 0xd0, g: 0xd0, b: 0xd0, a: 1 }, dash: null },
      b: { width: 12_700, color: { r: 0xd0, g: 0xd0, b: 0xd0, a: 1 }, dash: null },
      r: { width: 12_700, color: { r: 0xd0, g: 0xd0, b: 0xd0, a: 1 }, dash: null },
    };
    for (const row of rows) {
      for (const cell of row.cells) {
        if (!cell.borders) cell.borders = { ...fallback };
      }
    }
    options.onApproximatedStyle?.();
  }

  return {
    colWidths,
    rows,
    firstRow: attrBool(tblPr, "firstRow") !== false,
    bandRow: attrBool(tblPr, "bandRow") === true,
    styleId,
  };
}

function parseFillFromCell(tcPr: XmlNode | null, ctx: DrawingContext): Fill | null {
  if (!tcPr) return null;
  for (const candidate of tcPr.children) {
    if (
      candidate.local === "solidFill" ||
      candidate.local === "noFill" ||
      candidate.local === "gradFill" ||
      candidate.local === "blipFill" ||
      candidate.local === "pattFill"
    ) {
      const fill = parseFill(candidate, ctx);
      if (fill) return fill;
    }
  }
  return null;
}

export function tableRowHeights(table: TableData, totalHeight: number): number[] {
  const declared = table.rows.map((row) => emuToPx(row.height));
  const sum = declared.reduce((a, b) => a + b, 0);
  if (sum <= 0 || totalHeight <= 0) {
    const each = table.rows.length > 0 ? emuToPx(totalHeight) / table.rows.length : 0;
    return declared.map(() => each);
  }
  const scale = emuToPx(totalHeight) / sum;
  return declared.map((height) => height * scale);
}

export function tableColumnWidths(table: TableData): number[] {
  const declared = table.colWidths.map((width) => emuToPx(width));
  const sum = declared.reduce((a, b) => a + b, 0);
  if (sum > 0) return declared;
  return declared.map(() => 0);
}

export function borderWidthPx(emu: number): number {
  return Math.max(emuToPx(emu), ptToPx(0.5));
}
