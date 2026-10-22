/**
 * Page geometry and text metrics for the plain-text formats.
 *
 * These files have no page setup of their own, so the viewer supplies one: US
 * Letter, the same default the Word parser uses, in landscape when a table needs
 * the width. Pagination and rendering must agree exactly, or a page loses its
 * last line, so both sides import the numbers from here rather than each
 * deriving their own.
 *
 * The one approximation is glyph advance: a monospace face advances by 0.6em for
 * the whole Courier/Menlo/Consolas class, so a line measured as fitting at
 * `CHAR_WIDTH_PX` will fit. The font stack lists those faces explicitly for the
 * same reason — a variable-width fallback would break the grid.
 */

import { pxToEmu } from "../units";

export type Orientation = "portrait" | "landscape";

export interface PageBox {
  widthPx: number;
  heightPx: number;
}

/** US Letter at 96dpi. */
const PORTRAIT: PageBox = { widthPx: 816, heightPx: 1056 };
const LANDSCAPE: PageBox = { widthPx: 1056, heightPx: 816 };

export function pageBox(orientation: Orientation): PageBox {
  return orientation === "landscape" ? LANDSCAPE : PORTRAIT;
}

/** Margin inside the page box, all four sides. */
export const CONTENT_MARGIN_PX = 48;

/** Body text: 12px on a 16px baseline. */
export const BODY_FONT_PX = 12;
export const LINE_HEIGHT_PX = 16;
export const CHAR_WIDTH_PX = BODY_FONT_PX * 0.6;
/** Row height of a table row in the grid pages. */
export const ROW_HEIGHT_PX = 20;
/** Cell padding, split across the two sides of a column. */
export const CELL_PADDING_PX = 12;
/** Widest a single column is allowed to get, in characters. */
export const MAX_COLUMN_CHARS = 40;

export function contentBox(orientation: Orientation): { widthPx: number; heightPx: number } {
  const box = pageBox(orientation);
  return {
    widthPx: Math.max(box.widthPx - CONTENT_MARGIN_PX * 2, CHAR_WIDTH_PX),
    heightPx: Math.max(box.heightPx - CONTENT_MARGIN_PX * 2, LINE_HEIGHT_PX),
  };
}

/** How many monospace characters fit on one line. */
export function charsPerLine(orientation: Orientation): number {
  return Math.max(Math.floor(contentBox(orientation).widthPx / CHAR_WIDTH_PX), 1);
}

/**
 * How many baselines fit on one page. One is held back so a line that the
 * metrics underestimate cannot be pushed out of the page box.
 */
export function linesPerPage(orientation: Orientation): number {
  return Math.max(Math.floor(contentBox(orientation).heightPx / LINE_HEIGHT_PX) - 1, 1);
}

/** How many table rows fit on one page, header row included. */
export function rowsPerPage(orientation: Orientation): number {
  return Math.max(Math.floor(contentBox(orientation).heightPx / ROW_HEIGHT_PX) - 1, 1);
}

export function boxEmu(orientation: Orientation): { width: number; height: number } {
  const box = pageBox(orientation);
  return { width: pxToEmu(box.widthPx), height: pxToEmu(box.heightPx) };
}

/**
 * The faces the metrics assume. Kept as one string so the pages and the
 * measuring code cannot drift apart.
 */
export const MONOSPACE_STACK = 'Menlo, Consolas, "Courier New", monospace';

/** Tabs become spaces before anything is measured, so a tab cannot shift a column. */
export function expandTabs(line: string, size = 4): string {
  if (!line.includes("\t")) return line;
  let out = "";
  for (const character of line) {
    if (character === "\t") {
      const next = size - (out.length % size);
      out += " ".repeat(next);
    } else {
      out += character;
    }
  }
  return out;
}

/** Strip the control characters a text page would otherwise render as garbage. */
export function sanitizeLine(line: string): string {
  let out = "";
  for (let i = 0; i < line.length; i += 1) {
    const code = line.charCodeAt(i);
    if (code === 0x09) {
      out += line[i];
      continue;
    }
    // Keep printable characters; drop C0/C1 controls (a lone \r is handled by
    // the line splitter before this point).
    if (code < 0x20 || code === 0x7f) continue;
    if (code >= 0x80 && code <= 0x9f) continue;
    out += line[i];
  }
  return out;
}

/** Split decoded text into lines, tolerating every line ending. */
export function splitLines(text: string): string[] {
  return text.replace(/\r\n?/g, "\n").split("\n");
}
