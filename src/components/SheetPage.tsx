/**
 * Spreadsheet page rendering.
 *
 * A worksheet is a scrollable grid, not a paginated page, so each sheet gets a
 * fixed page box with its own scroll container. The table is plain HTML: values
 * are React text nodes, never markup, and a cell is never interpreted as a
 * formula. Row numbers and column letters are drawn as sticky headers so the
 * grid remains readable while scrolling.
 *
 * The viewer's thumbnails call the same renderer with `interactive: false`;
 * there the scrollbars and sticky positioning are suppressed and the top-left
 * corner is shown, which is what a thumbnail should be.
 */

import { useLayoutEffect, useRef, useState, type CSSProperties, type ReactNode } from "react";
import { selectionText } from "./worksheetUtils";
import type { PageRenderer } from "../renderer/pageRenderer";
import type { SheetCell, SheetPageData } from "../parser/xlsx/model";
import type { MediaStore } from "../parser/media";
import { ShapeNodeView, createLinkBuilder } from "../renderer/RenderShape";
import { emuToPx } from "../parser/units";
import { DEFAULT_CHART_PALETTE } from "../renderer/RenderChart";

const ROW_HEADER_WIDTH_PX = 44;
const ROW_HEIGHT_PX = 20;

/** Excel-style column label: A, B, …, Z, AA, AB, … */
export function columnLabel(index: number): string {
  let value = Math.max(index, 0) + 1;
  let label = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

function cellClass(cell: SheetCell): string {
  return [
    "radv-sheet-cell",
    cell.numeric ? "radv-sheet-numeric" : "",
    cell.align === "right" ? "radv-sheet-align-right" : "",
    cell.align === "center" ? "radv-sheet-align-center" : "",
  ]
    .filter(Boolean)
    .join(" ");
}

function canOverflow(cell: SheetCell): boolean {
  return Boolean(cell.text) && !cell.numeric && cell.align === "left" && cell.style?.whiteSpace !== "pre-wrap";
}

function textWidth(cell: SheetCell, row: SheetCell[], index: number, widths: number[]): CSSProperties {
  if (!canOverflow(cell)) return {};
  let end = index + cell.colSpan;
  while (end < row.length && !row[end]?.text && !row[end]?.covered) end++;
  // A used range is not a clipping boundary: blank worksheet space continues after it.
  return end === row.length
    ? { width: "max-content", minWidth: "100%" }
    : { width: `${widths.slice(index, end).reduce((sum, width) => sum + width, 0)}px` };
}

export interface SheetPageProps {
  data: SheetPageData;
  width: number;
  height: number;
  interactive: boolean;
  media?: MediaStore;
  warn?: (code: string, message: string) => void;
  savedScroll?: { left: number; top: number };
}

export function SheetPage({ data, width, height, interactive, media, warn, savedScroll }: SheetPageProps): ReactNode {
  const scrollRef = useRef<HTMLDivElement>(null);
  const [selection, setSelection] = useState<{ anchor: [number, number]; focus: [number, number] } | null>(null);
  const dragging = useRef(false);
  useLayoutEffect(() => {
    const element = scrollRef.current;
    if (element && savedScroll) { element.scrollLeft = savedScroll.left; element.scrollTop = savedScroll.top; }
    setSelection(null);
  }, [data, savedScroll]);
  const selected = (row: number, col: number) => selection && row >= Math.min(selection.anchor[0], selection.focus[0]) && row <= Math.max(selection.anchor[0], selection.focus[0]) && col >= Math.min(selection.anchor[1], selection.focus[1]) && col <= Math.max(selection.anchor[1], selection.focus[1]);
  const totalWidth = data.widths.reduce((sum, value) => sum + value, 0) + ROW_HEADER_WIDTH_PX;
  const rowHeight = (index: number) => data.heights?.[index] ?? ROW_HEIGHT_PX;
  const contentWidth = Math.max(totalWidth, ...(data.drawings ?? []).map((node) => emuToPx(node.xfrm.x + node.xfrm.cx) + ROW_HEADER_WIDTH_PX));
  const contentHeight = Math.max(data.rows.reduce((sum, _, index) => sum + rowHeight(index), 22), ...(data.drawings ?? []).map((node) => emuToPx(node.xfrm.y + node.xfrm.cy))) + 22;
  const hasRows = data.rows.length > 0 && data.rows[0]?.length !== 0;

  return (
    <div
      className="radv-sheet-page"
      data-testid="rv-sheet-page"
      style={{ width: `${width}px`, height: `${height}px` }}
    >
      <div
        ref={scrollRef}
        tabIndex={interactive ? 0 : undefined}
        aria-label={`${data.name} worksheet; select cells and press Control or Command C to copy`}
        onScroll={(event) => { if (savedScroll) { savedScroll.left = event.currentTarget.scrollLeft; savedScroll.top = event.currentTarget.scrollTop; } }}
        onPointerUp={() => { dragging.current = false; }}
        onPointerLeave={() => { dragging.current = false; }}
        onCopy={(event) => { if (!selection) return; event.clipboardData.setData("text/plain", selectionText(data, selection.anchor, selection.focus)); event.preventDefault(); }}
        className={`radv-sheet-scroll${interactive ? "" : " radv-sheet-scroll-static"}`}
        style={{ width: `${width}px`, height: `${height}px` }}
      >
        <div className="radv-sheet-surface" style={{ position: "relative", width: `${contentWidth}px`, minHeight: `${contentHeight}px` }}>
        {hasRows ? (
          <table
            className="radv-sheet"
            style={{ width: `${totalWidth}px` }}
            aria-label={data.name}
            role={interactive ? "grid" : undefined}
            aria-rowcount={data.rows.length + 1}
            aria-colcount={data.widths.length + 1}
            data-gridlines={data.showGridLines === false ? "false" : "true"}
          >
            <colgroup>
              <col style={{ width: `${ROW_HEADER_WIDTH_PX}px` }} />
              {data.widths.map((columnWidth, index) => (
                <col key={index} style={{ width: `${columnWidth}px`, visibility: columnWidth === 0 ? "collapse" : undefined }} />
              ))}
            </colgroup>
            <thead>
              <tr>
                <th className="radv-sheet-corner" aria-hidden="true" />
                {data.widths.map((_, index) => (
                  <th key={index} className="radv-sheet-column-head" scope="col" style={{ display: data.widths[index] === 0 ? "none" : undefined }}>
                    {columnLabel(index)}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {data.rows.map((row, rowIndex) => (
                <tr key={rowIndex} style={{ height: `${rowHeight(rowIndex)}px`, display: rowHeight(rowIndex) === 0 ? "none" : undefined }}>
                  <th className="radv-sheet-row-head" scope="row">
                    {rowIndex + 1}
                  </th>
                  {row.map((cell, cellIndex) =>
                    cell.covered ? null : (
                      <td
                        key={cellIndex}
                        className={`${cellClass(cell)}${selected(rowIndex, cellIndex) ? " radv-sheet-selected" : ""}`}
                        role={interactive ? "gridcell" : undefined}
                        aria-selected={interactive ? Boolean(selected(rowIndex, cellIndex)) : undefined}
                        aria-label={`${columnLabel(cellIndex)}${rowIndex + 1}: ${cell.text}`}
                        data-cell={`${rowIndex}:${cellIndex}`}
                        tabIndex={interactive && (selection ? selection.focus[0] === rowIndex && selection.focus[1] === cellIndex : rowIndex === 0 && cellIndex === 0) ? 0 : -1}
                        onPointerDown={interactive ? (event) => { dragging.current = true; setSelection({ anchor: event.shiftKey && selection ? selection.anchor : [rowIndex, cellIndex], focus: [rowIndex, cellIndex] }); } : undefined}
                        onPointerEnter={interactive ? (event) => { if (dragging.current && event.buttons === 1) setSelection((previous) => ({ anchor: previous?.anchor ?? [rowIndex, cellIndex], focus: [rowIndex, cellIndex] })); } : undefined}
                        onKeyDown={interactive ? (event) => {
                          const delta = event.key === "ArrowRight" ? [0, 1] : event.key === "ArrowLeft" ? [0, -1] : event.key === "ArrowDown" ? [1, 0] : event.key === "ArrowUp" ? [-1, 0] : null;
                          if (!delta) return;
                          event.preventDefault(); event.stopPropagation();
                          let r = rowIndex, c = cellIndex;
                          do { r += delta[0]!; c += delta[1]!; } while (r >= 0 && c >= 0 && r < data.rows.length && c < data.widths.length && (data.rows[r]?.[c]?.covered || rowHeight(r) === 0 || data.widths[c] === 0));
                          if (r < 0 || c < 0 || r >= data.rows.length || c >= data.widths.length) return;
                          setSelection({ anchor: event.shiftKey && selection ? selection.anchor : [r, c], focus: [r, c] });
                          scrollRef.current?.querySelector<HTMLElement>(`[data-cell="${r}:${c}"]`)?.focus();
                        } : undefined}
                        colSpan={cell.colSpan > 1 ? cell.colSpan : undefined}
                        rowSpan={cell.rowSpan > 1 ? cell.rowSpan : undefined}
                        title={cell.text.length > 0 ? cell.text : undefined}
                        style={{ ...cell.style, position: "relative", overflow: canOverflow(cell) ? "visible" : undefined, zIndex: cell.text ? 1 : undefined, display: data.widths.slice(cellIndex, cellIndex + cell.colSpan).every((width) => width === 0) ? "none" : undefined }}
                      >
                        {cell.dataBar && <div className="radv-sheet-data-bar" style={{ left: cell.dataBar.left === undefined ? undefined : `${cell.dataBar.left}%`, width: `${cell.dataBar.width}%`, background: cell.dataBar.color }} />}
                        <div className="radv-sheet-cell-content" style={{ ...textWidth(cell, row, cellIndex, data.widths), minHeight: data.autoRowHeight ? `${rowHeight(rowIndex) - 1}px` : undefined, height: data.autoRowHeight ? "auto" : `${Math.max(0, data.rows.slice(rowIndex, rowIndex + cell.rowSpan).reduce((sum, _, index) => sum + rowHeight(rowIndex + index), 0) - 1)}px`, alignItems: cell.style?.verticalAlign === "middle" ? "center" : cell.style?.verticalAlign === "top" ? "flex-start" : "flex-end" }}>
                        <span style={{ width: "100%", minWidth: 0, position: "relative", whiteSpace: cell.style?.whiteSpace ?? "pre" }}>
                        {cell.icon && <span aria-label={cell.icon.label} style={{ color: cell.icon.color, marginRight: 4 }}>{cell.icon.text}</span>}{cell.hideValue ? "" : cell.text}
                        </span>
                        </div>
                      </td>
                    ),
                  )}
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          !data.drawings?.length && <div className="radv-sheet-empty" role="note">
            This sheet is empty
          </div>
        )}
        {media && (data.drawings ?? []).map((node, index) => <div key={index} className="radv-sheet-drawing" style={{ position: "absolute", left: ROW_HEADER_WIDTH_PX, top: 22, pointerEvents: interactive ? "auto" : "none" }}>
          <ShapeNodeView node={node.kind === "chart" ? { ...node, xfrm: { ...node.xfrm, y: node.xfrm.y + Math.min(12 * 9525, node.xfrm.cy / 10), cy: node.xfrm.cy - Math.min(12 * 9525, node.xfrm.cy / 10) } } : node} buildLink={createLinkBuilder(interactive)} context={{ media, scope: `sheet-${data.name}-${index}`, interactive, palette: DEFAULT_CHART_PALETTE, warn: warn ?? (() => undefined) }} />
        </div>)}
        </div>
      </div>
      {data.truncated ? (
        <div className="radv-sheet-truncated" role="note">
          Only the used range up to the configured cell limit is shown
        </div>
      ) : null}
    </div>
  );
}

/** Build the page renderer the viewer calls for a workbook. */
export function createSheetPageRenderer(pages: readonly SheetPageData[]): PageRenderer {
  const positions = pages.map(() => ({ left: 0, top: 0 }));
  const renderer: PageRenderer = ({ page, width, height, interactive, media, warn }) => {
    const data = pages[page.index];
    if (!data) return null;
    return <SheetPage key={page.index} savedScroll={interactive ? positions[page.index] : undefined} data={data} width={width} height={height} interactive={interactive} media={media} warn={warn} />;
  };
  renderer.worksheetPages = pages;
  return renderer;
}
