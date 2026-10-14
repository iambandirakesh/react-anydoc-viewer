/**
 * Table rendering.
 *
 * Cells are absolutely positioned from the declared grid so that merged cells,
 * per-edge borders and per-cell insets all land exactly where PowerPoint puts
 * them, which an HTML `<table>` cannot express reliably.
 */

import type { CSSProperties, ReactNode } from "react";
import type { TableCell, TableNode, TableRow } from "../model/types";
import { emuToPx } from "../parser/units";
import { cssBackground } from "./style";
import { renderTextBody, type RenderTextOptions } from "./RenderText";

export interface TableRenderOptions extends RenderTextOptions {
  buildLink: (href: string, tooltip: string | undefined, children: ReactNode, key: string) => ReactNode;
}

function edgeStyle(width: number, color: string | undefined, horizontal: boolean, atStart: boolean): CSSProperties {
  const thickness = Math.max(width, 0.75);
  return {
    position: "absolute",
    backgroundColor: color ?? "transparent",
    ...(horizontal
      ? { left: 0, right: 0, height: `${thickness}px`, [atStart ? "top" : "bottom"]: 0 }
      : { top: 0, bottom: 0, width: `${thickness}px`, [atStart ? "left" : "right"]: 0 }),
  };
}

function cellBorders(cell: TableCell): ReactNode[] {
  const borders = cell.borders;
  if (!borders) return [];
  const nodes: ReactNode[] = [];
  const push = (edge: "t" | "l" | "b" | "r", horizontal: boolean, atStart: boolean) => {
    const value = borders[edge];
    if (!value || value.width <= 0) return;
    nodes.push(
      <span
        key={edge}
        aria-hidden="true"
        style={edgeStyle(emuToPx(value.width), value.color ? `rgba(${Math.round(value.color.r)}, ${Math.round(value.color.g)}, ${Math.round(value.color.b)}, ${value.color.a})` : undefined, horizontal, atStart)}
      />,
    );
  };
  push("t", true, true);
  push("b", true, false);
  push("l", false, true);
  push("r", false, false);
  return nodes;
}

export function renderTable(
  node: TableNode,
  options: TableRenderOptions,
): ReactNode {
  const table = node.table;
  const frameWidth = emuToPx(node.xfrm.cx);
  const frameHeight = emuToPx(node.xfrm.cy);

  const declaredCols = table.colWidths.map(emuToPx);
  const declaredTotal = declaredCols.reduce((sum, value) => sum + value, 0);
  const columns =
    declaredTotal > 0 && frameWidth > 0
      ? declaredCols.map((value) => (value * frameWidth) / declaredTotal)
      : declaredCols.map(() => (declaredCols.length > 0 ? frameWidth / declaredCols.length : 0));

  const rows: TableRow[] = table.rows;
  const declaredRowHeights = rows.map((row) => emuToPx(row.height));
  const rowTotal = declaredRowHeights.reduce((sum, value) => sum + value, 0);
  const heights =
    rowTotal > 0 && frameHeight > 0
      ? declaredRowHeights.map((value) => (value * frameHeight) / rowTotal)
      : declaredRowHeights.map(() => (rows.length > 0 ? frameHeight / rows.length : 0));

  const offsetsX: number[] = [];
  let x = 0;
  for (const width of columns) {
    offsetsX.push(x);
    x += width;
  }
  const offsetsY: number[] = [];
  let y = 0;
  for (const height of heights) {
    offsetsY.push(y);
    y += height;
  }

  const occupied: boolean[][] = rows.map(() => columns.map(() => false));
  const cells: ReactNode[] = [];

  rows.forEach((row, rowIndex) => {
    let columnCursor = 0;
    row.cells.forEach((cell, cellIndex) => {
      while (columnCursor < columns.length && occupied[rowIndex]?.[columnCursor]) columnCursor += 1;
      if (columnCursor >= columns.length) return;
      const colSpan = Math.max(1, Math.min(cell.gridSpan, columns.length - columnCursor));
      const rowSpan = Math.max(1, Math.min(cell.rowSpan, rows.length - rowIndex));

      for (let r = rowIndex; r < rowIndex + rowSpan; r += 1) {
        for (let c = columnCursor; c < columnCursor + colSpan; c += 1) {
          const rowFlags = occupied[r];
          if (rowFlags) rowFlags[c] = true;
        }
      }

      const left = offsetsX[columnCursor] ?? 0;
      const top = offsetsY[rowIndex] ?? 0;
      let width = 0;
      for (let c = columnCursor; c < columnCursor + colSpan; c += 1) width += columns[c] ?? 0;
      let height = 0;
      for (let r = rowIndex; r < rowIndex + rowSpan; r += 1) height += heights[r] ?? 0;

      if (cell.merged && rowSpan === 1 && colSpan === 1 && cell.hMerge) {
        // Cell skipped by an explicit horizontal merge.
        columnCursor += colSpan;
        return;
      }

      const cellStyle: CSSProperties = {
        position: "absolute",
        left: `${left}px`,
        top: `${top}px`,
        width: `${width}px`,
        height: `${height}px`,
        boxSizing: "border-box",
        background: cssBackground(cell.fill),
        overflow: "hidden",
      };

      const inner: CSSProperties = {
        position: "absolute",
        left: `${cell.insets.l}px`,
        top: `${cell.insets.t}px`,
        width: `${Math.max(width - cell.insets.l - cell.insets.r, 0)}px`,
        height: `${Math.max(height - cell.insets.t - cell.insets.b, 0)}px`,
        display: "flex",
        flexDirection: "column",
        justifyContent:
          cell.anchor === "ctr" ? "center" : cell.anchor === "b" ? "flex-end" : "flex-start",
      };

      const text = cell.text
        ? renderTextBody(cell.text, options.buildLink, options).node
        : null;

      cells.push(
        <div key={`c${rowIndex}-${cellIndex}`} style={cellStyle} data-cell={`${rowIndex}-${columnCursor}`}>
          {cellBorders(cell)}
          {cell.text ? (
            <div style={inner}>
              <div style={{ width: "100%" }}>{text}</div>
            </div>
          ) : null}
        </div>,
      );
      columnCursor += colSpan;
    });
  });

  return <>{cells}</>;
}
