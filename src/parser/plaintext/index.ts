/**
 * Plain-text, CSV and source-code documents.
 *
 * A `.txt`, `.log`, `.json`, `.csv`, `.md` or `.html` file has no page setup, no
 * relationships and no shapes, so it becomes one page per chunk of content, drawn
 * by its own page renderer — the same seam the PDF adapter uses. Navigation,
 * thumbnails, zoom, fit modes, presentation mode and the imperative handle all
 * work unchanged, because none of them know what a page contains.
 *
 * Nothing in these files is ever interpreted as markup or script: every character
 * reaches the DOM as text. An HTML file is no exception — it is read and paginated
 * like any other text, never parsed, never framed and never given an origin, so
 * nothing in it can run a script or load anything.
 */

import type { DocumentFormat } from "../../errors";
import type { DocumentModel, PageModel } from "../../model/types";
import { createEmptyMediaStore, type MediaStore } from "../media";
import { DEFAULT_COLOR_MAP, createEmptyTheme } from "../drawingml/theme";
import { Diagnostics, DEFAULT_MAX_DIAGNOSTICS } from "../diagnostics";
import type { PageContent } from "../pageContent";
import type { SheetPageData } from "../xlsx/model";
import { classifyPlainText, classifyTextBytes, type Classification, type PlainTextKind } from "./classify";
import { decodeTextFile } from "./decode";
import { paginateText, MAX_PAGES, type TextPage } from "./paginate";
import { delimiterMessage, parseCsv, raggedRowMessage, type CsvTable } from "./csv";
import {
  boxEmu,
  CELL_PADDING_PX,
  CHAR_WIDTH_PX,
  MAX_COLUMN_CHARS,
  contentBox,
  type Orientation,
} from "./metrics";

/**
 * What a set of bytes is, if it is one of the text-based formats.
 *
 * `null` means "not text at all" — the bytes did not decode as text, so the
 * caller must not treat them as a document. A plain `.txt` file and a `.csv` are
 * told apart by content, since neither carries a signature.
 */
export function detectPlainTextKind(bytes: Uint8Array): DocumentFormat | null {
  const classification = classifyTextBytes(bytes);
  if (classification.confidence <= 0) return null;
  return classification.kind === "csv" ? "csv" : "text";
}

export interface ParsedPlainText {
  format: DocumentFormat;
  model: DocumentModel;
  media: MediaStore;
  content: PageContent;
  dispose(): void;
}

export interface ParsePlainTextOptions {
  /** the classification detection already made; the bytes are classified here when absent */
  classification?: Classification;
  /** the caller's override for an ambiguous text file */
  contentHint?: PlainTextKind;
  maxDiagnostics?: number;
}

/** A table laid out as rows plus column widths, before it becomes a worksheet. */
interface GridPageData {
  /** the rows of this page, header included when the table has one */
  rows: string[][];
  /** column widths in pixels, shared by every page of the table */
  widths: number[];
  /** true when the first row of `rows` is the table's header */
  header: boolean;
}

/** How many rows are measured when sizing columns; a huge table sizes from its head. */
const COLUMN_SAMPLE_ROWS = 500;
/** Above this size a JSON file is shown as it is instead of being pretty-printed. */
const JSON_PRETTY_LIMIT = 2 * 1024 * 1024;

function clamp(value: number, min: number, max: number): number {
  return value < min ? min : value > max ? max : value;
}

/** One page per chunk of text; the renderer draws whichever page is on screen. */
function textPages(pages: readonly TextPage[]): PageModel[] {
  return pages.map((_, index) => ({
    index,
    partName: `text#page${index + 1}`,
    name: null,
    slideshowHidden: false,
    shapes: [],
    inheritedShapes: [],
    background: { type: "solid", color: { r: 255, g: 255, b: 255, a: 1 } },
    diagnostics: [],
  }));
}

function gridPages(pages: readonly GridPageData[]): PageModel[] {
  return pages.map((_, index) => ({
    index,
    partName: `grid#page${index + 1}`,
    name: null,
    slideshowHidden: false,
    shapes: [],
    inheritedShapes: [],
    background: { type: "solid", color: { r: 255, g: 255, b: 255, a: 1 } },
    diagnostics: [],
  }));
}

/** Column widths in pixels, from the widest value in the head of each column. */
function columnWidths(table: CsvTable): number[] {
  const sample = table.rows.slice(0, COLUMN_SAMPLE_ROWS);
  return table.header.map((title, index) => {
    let widest = title.length;
    for (const row of sample) {
      const length = (row[index] ?? "").length;
      if (length > widest) widest = length;
    }
    return Math.min(320, clamp(widest, 4, MAX_COLUMN_CHARS) * CHAR_WIDTH_PX + CELL_PADDING_PX);
  });
}

interface GridLayout {
  pages: GridPageData[];
  widths: number[];
  orientation: Orientation;
}

/** Keep the entire parsed table in one scrollable worksheet. */
function layoutGrid(table: CsvTable): GridLayout {
  const widths = columnWidths(table);
  const orientation: Orientation = widths.reduce((a, b) => a + b, 0) > contentBox("portrait").widthPx ? "landscape" : "portrait";
  return { pages: [{ rows: table.header.length ? [table.header, ...table.rows] : table.rows, widths, header: table.header.length > 0 }], widths, orientation };
}

/** Assemble the model both kinds of page share. */
function buildModel(
  slides: PageModel[],
  orientation: Orientation,
  diagnostics: Diagnostics,
): DocumentModel {
  const box = boxEmu(orientation);
  return {
    width: box.width,
    height: box.height,
    slides,
    theme: createEmptyTheme(),
    colorMap: DEFAULT_COLOR_MAP,
    textStyles: { title: [], body: [], other: [] },
    title: null,
    author: null,
    partCount: slides.length,
    diagnostics: diagnostics.list(),
  };
}

/**
 * Parse a text-based file into pages.
 *
 * The result is dispatched before the OPC/ZIP paths, because these files are not
 * containers: the bytes are decoded, classified (unless the caller already did),
 * and laid out.
 */
export function parsePlainTextDocument(
  bytes: Uint8Array,
  options: ParsePlainTextOptions = {},
): ParsedPlainText {
  const diagnostics = new Diagnostics(options.maxDiagnostics ?? DEFAULT_MAX_DIAGNOSTICS);
  const decoded = decodeTextFile(bytes);
  const classification = options.classification ?? classifyPlainText(decoded.text);
  const kind: PlainTextKind = options.contentHint ?? classification.kind;

  if (decoded.lossy) {
    diagnostics.add({
      severity: "warning",
      code: "text.encoding",
      message:
        "The file is not valid UTF-8 and was decoded as Windows-1252; some characters may be shown incorrectly.",
    });
  }
  if (!options.contentHint && kind !== "text") {
    diagnostics.add({
      severity: "info",
      code: `text.classified.${kind}`,
      message: `This file has no signature, so its type was decided from its content: ${classification.reason}.`,
    });
  }

  let text = decoded.text;
  if (kind === "json" && text.length <= JSON_PRETTY_LIMIT) {
    try {
      text = JSON.stringify(JSON.parse(text), null, 2);
    } catch {
      // Classification already parsed it; if that changed, the source is shown.
    }
  } else if (kind === "json") {
    diagnostics.add({
      severity: "info",
      code: "text.json.unformatted",
      message: `This JSON file is larger than ${Math.round(
        JSON_PRETTY_LIMIT / (1024 * 1024),
      )}MiB, so it is shown unformatted.`,
    });
  }
  if (kind === "html") {
    diagnostics.add({
      severity: "warning",
      code: "text.html.source",
      message:
        "This is an HTML file, shown as source text: nothing in it is executed, and no request is made on its behalf.",
    });
  }
  if (kind === "markdown") {
    diagnostics.add({
      severity: "info",
      code: "text.markdown.source",
      message: "Markdown is currently shown as source text.",
    });
  }

  const media = createEmptyMediaStore();

  if (kind === "csv") {
    const table = parseCsv(text, {
      ...(classification.delimiter && !options.contentHint ? { delimiter: classification.delimiter } : {}),
    });
    const layout = layoutGrid(table);
    if (table.raggedRows > 0) {
      diagnostics.add({
        severity: "warning",
        code: "text.csv.ragged",
        message: raggedRowMessage(table) ?? "",
      });
    }
    if (table.truncated) {
      diagnostics.add({
        severity: "warning",
        code: "text.csv.truncated",
        message: `This table has ${table.totalRows} data rows; the first ${table.rows.length} are shown.`,
      });
    }
    diagnostics.add({
      severity: "info",
      code: "text.csv.delimiter",
      message: delimiterMessage(table),
    });

    const sourceColumns = Math.max(table.columns, ...layout.pages[0]!.rows.map((row) => row.length));
    const columns = Math.min(200, sourceColumns);
    const rowLimit = Math.max(1, Math.floor(10000 / Math.max(1, columns)));
    const rows = layout.pages[0]!.rows.slice(0, rowLimit);
    const truncated = table.truncated || columns < sourceColumns || rows.length < layout.pages[0]!.rows.length;
    if (truncated && !table.truncated) diagnostics.add({ severity: "warning", code: "text.csv.cell-limit", message: "Only the first 10,000 cells and 200 columns are shown in the table." });
    const data: SheetPageData = {
      name: "Table",
      rows: rows.map((row, rowIndex) => Array.from({ length: columns }, (_, colIndex) => ({
        text: row[colIndex] ?? "", align: "left" as const, numeric: false, covered: false, colSpan: 1, rowSpan: 1,
        style: { whiteSpace: "pre-wrap" as const, overflowWrap: "anywhere", verticalAlign: "middle", ...(layout.pages[0]!.header && rowIndex === 0 ? { fontWeight: 700, backgroundColor: "#edf0f5" } : {}) },
      }))),
      widths: Array.from({ length: columns }, (_, index) => layout.widths[index] ?? 120),
      heights: rows.map(() => 26), autoRowHeight: true, showGridLines: true, truncated,
    };
    const slides = gridPages(layout.pages);
    slides[0]!.name = "Table";
    return {
      format: "csv",
      model: buildModel(slides, layout.orientation, diagnostics),
      media,
      content: { kind: "sheet", pages: [data] },
      dispose: () => media.dispose(),
    };
  }

  const pagination = paginateText(text, { orientation: "portrait" });
  if (!pagination.complete) {
    diagnostics.add({
      severity: "warning",
      code: "text.truncated",
      message: `The file has more than ${MAX_PAGES} pages of text; only the first ${MAX_PAGES} are shown.`,
    });
  }

  const slides = textPages(pagination.pages);
  return {
    format: "text",
    model: buildModel(slides, "portrait", diagnostics),
    media,
    content: { kind: "text", pages: pagination.pages },
    dispose: () => media.dispose(),
  };
}
