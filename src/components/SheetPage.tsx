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

import type { ReactNode } from "react";
import type { PageRenderer } from "../renderer/pageRenderer";
import type { SheetCell, SheetPageData } from "../parser/xlsx/model";

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

export interface SheetPageProps {
  data: SheetPageData;
  width: number;
  height: number;
  interactive: boolean;
}

export function SheetPage({ data, width, height, interactive }: SheetPageProps): ReactNode {
  const totalWidth = data.widths.reduce((sum, value) => sum + value, 0) + ROW_HEADER_WIDTH_PX;
  const hasRows = data.rows.length > 0 && data.rows[0]?.length !== 0;

  return (
    <div
      className="radv-sheet-page"
      data-testid="rv-sheet-page"
      style={{ width: `${width}px`, height: `${height}px` }}
    >
      <div
        className={`radv-sheet-scroll${interactive ? "" : " radv-sheet-scroll-static"}`}
        style={{ width: `${width}px`, height: `${height}px` }}
      >
        {hasRows ? (
          <table
            className="radv-sheet"
            style={{ width: `${totalWidth}px` }}
            aria-label={data.name}
          >
            <colgroup>
              <col style={{ width: `${ROW_HEADER_WIDTH_PX}px` }} />
              {data.widths.map((columnWidth, index) => (
                <col key={index} style={{ width: `${columnWidth}px` }} />
              ))}
            </colgroup>
            <thead>
              <tr>
                <th className="radv-sheet-corner" aria-hidden="true" />
                {data.widths.map((_, index) => (
                  <th key={index} className="radv-sheet-column-head" scope="col">
                    {columnLabel(index)}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {data.rows.map((row, rowIndex) => (
                <tr key={rowIndex} style={{ height: `${ROW_HEIGHT_PX}px` }}>
                  <th className="radv-sheet-row-head" scope="row">
                    {rowIndex + 1}
                  </th>
                  {row.map((cell, cellIndex) =>
                    cell.covered ? null : (
                      <td
                        key={cellIndex}
                        className={cellClass(cell)}
                        colSpan={cell.colSpan > 1 ? cell.colSpan : undefined}
                        rowSpan={cell.rowSpan > 1 ? cell.rowSpan : undefined}
                        title={cell.text.length > 0 ? cell.text : undefined}
                      >
                        {cell.text}
                      </td>
                    ),
                  )}
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          <div className="radv-sheet-empty" role="note">
            This sheet is empty
          </div>
        )}
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
  return ({ page, width, height, interactive }) => {
    const data = pages[page.index];
    if (!data) return null;
    return <SheetPage data={data} width={width} height={height} interactive={interactive} />;
  };
}
