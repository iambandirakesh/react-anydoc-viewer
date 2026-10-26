/**
 * Tabular page rendering.
 *
 * A page holds a window of rows and the column widths computed for the whole
 * table, so every page of one table lines up with the others. Widths are applied
 * with `table-layout: fixed` because the value of a grid is that column *n* is
 * the same width everywhere; a cell is clipped with an ellipsis instead of
 * widening its column and shifting every row below it.
 *
 * The whole table is rendered as text inside `<td>` elements — React escapes it,
 * and a cell value is never interpreted as markup or as a formula.
 */

import type { ReactNode } from "react";
import type { PageRenderer } from "../renderer/pageRenderer";
import { BODY_FONT_PX, CONTENT_MARGIN_PX, ROW_HEIGHT_PX } from "../parser/plaintext/metrics";

export interface GridPageData {
  /** the rows of this page, header included when the table has one */
  rows: string[][];
  /** column widths in pixels, shared by every page of the table */
  widths: number[];
  /** true when the first row of `rows` is the table's header */
  header: boolean;
}

export interface GridPageProps extends GridPageData {
  width: number;
  height: number;
  /** name of the sheet or file, used for the accessible label */
  label?: string;
}

export function GridPage({ rows, widths, header, width, height, label }: GridPageProps): ReactNode {
  const totalWidth = widths.reduce((sum, value) => sum + value, 0);
  return (
    <div
      className="radv-grid-page"
      data-testid="rv-grid-page"
      style={{ width: `${width}px`, height: `${height}px`, padding: `${CONTENT_MARGIN_PX}px`, fontSize: `${BODY_FONT_PX}px` }}
    >
      <table className="radv-grid" style={{ width: `${totalWidth}px` }} aria-label={label}>
        <colgroup>
          {widths.map((columnWidth, index) => (
            <col key={index} style={{ width: `${columnWidth}px` }} />
          ))}
        </colgroup>
        <tbody>
          {rows.map((row, rowIndex) => (
            <tr key={rowIndex} style={{ height: `${ROW_HEIGHT_PX}px` }}>
              {row.map((cell, cellIndex) => {
                const Tag = header && rowIndex === 0 ? "th" : "td";
                return (
                  <Tag key={cellIndex} className={header && rowIndex === 0 ? "radv-grid-head" : "radv-grid-cell"}>
                    {cell}
                  </Tag>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/** Build the page renderer the viewer calls for a table. */
export function createGridPageRenderer(pages: readonly GridPageData[], label?: string): PageRenderer {
  return ({ page, width, height }) => {
    const chunk = pages[page.index];
    if (!chunk) return null;
    return (
      <GridPage
        rows={chunk.rows}
        widths={chunk.widths}
        header={chunk.header}
        width={width}
        height={height}
        label={label}
      />
    );
  };
}
