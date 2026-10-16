/**
 * Pagination for plain text.
 *
 * Text has no page breaks of its own, so lines are wrapped to the page's column
 * count and then distributed over pages of a fixed line count. Both numbers come
 * from `metrics.ts`, which the renderer imports as well: wrapping here and
 * rendering there have to agree, or a page shows fewer lines than it holds.
 *
 * A file with more pages than {@link MAX_PAGES} is truncated rather than turned
 * into a ten-thousand-page model, and the truncation is reported.
 */

import { expandTabs, linesPerPage, charsPerLine, sanitizeLine, splitLines, type Orientation } from "./metrics";

/** Ceiling on pages produced from one text file. */
export const MAX_PAGES = 1000;

export interface TextPage {
  /** already wrapped to the page's column count */
  lines: string[];
}

export interface Pagination {
  pages: TextPage[];
  /** wrap the whole file? */
  complete: boolean;
  totalLines: number;
  /** source lines that had to be broken to fit the page width */
  wrappedLines: number;
}

/**
 * Break a line to fit, preferring the last space before the limit so words stay
 * together; a long token with no space is cut at the limit. Indentation is kept,
 * because columns are meaningful in logs and fixed-width data.
 */
export function wrapLine(line: string, limit: number): string[] {
  if (limit < 1) return [line];
  if (line.length <= limit) return [line];
  const out: string[] = [];
  let rest = line;
  while (rest.length > limit) {
    const space = rest.lastIndexOf(" ", limit);
    const cut = space > Math.floor(limit * 0.6) ? space + 1 : limit;
    out.push(rest.slice(0, cut));
    rest = rest.slice(cut);
  }
  out.push(rest);
  return out;
}

export interface PaginateOptions {
  orientation?: Orientation;
  /** how wide the page is, in characters (defaults to the page geometry) */
  columns?: number;
  /** how many lines a page holds (defaults to the page geometry) */
  rows?: number;
  maxPages?: number;
}

export function paginateText(text: string, options: PaginateOptions = {}): Pagination {
  const orientation = options.orientation ?? "portrait";
  const limit = options.columns ?? charsPerLine(orientation);
  const perPage = options.rows ?? linesPerPage(orientation);
  const maxPages = options.maxPages ?? MAX_PAGES;

  const sources = splitLines(text);
  const pages: TextPage[] = [];
  let current: string[] = [];
  let wrappedLines = 0;
  let complete = true;

  for (const source of sources) {
    const wrapped = wrapLine(expandTabs(sanitizeLine(source)), limit);
    if (wrapped.length > 1) wrappedLines += 1;
    for (const line of wrapped) {
      if (pages.length >= maxPages) {
        complete = false;
        break;
      }
      current.push(line);
      if (current.length >= perPage) {
        pages.push({ lines: current });
        current = [];
      }
    }
    if (!complete) break;
  }

  if (current.length > 0 && pages.length < maxPages) pages.push({ lines: current });

  return { pages, complete, totalLines: sources.length, wrappedLines };
}
