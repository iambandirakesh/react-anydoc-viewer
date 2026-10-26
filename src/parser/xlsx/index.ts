/**
 * SpreadsheetML (`.xlsx`/`.xlsm`) parsing.
 *
 * A workbook is a set of sheets, and each sheet is a grid rather than a page of
 * shapes. The viewer's model already allows a format to draw its own pages (PDF,
 * text and images do), so a sheet becomes one page and the workbook's sheet list
 * becomes the page list. Navigation, thumbnails, zoom, presentation mode and the
 * error boundary therefore work without any format-specific viewer code.
 *
 * The parser reads values and the small amount of style data needed to display
 * them (number formats and dates). Formulas are never evaluated: the cached
 * `<v>` written by the producer is shown, and a formula without a cache is
 * reported instead of guessed at. Conditional formatting, charts, pivot tables
 * and macros are outside the read-only surface.
 */

import type { PresentationModel, SlideModel } from "../../model/types";
import { DocumentViewerError } from "../../errors";
import { readText, type Archive, type ArchiveLimits } from "../archive";
import { attr, attrBool, attrNum, child, children, find, parseXml, relAttr, type XmlNode } from "../xml";
import { parseRelationships, relTypeName, relsPartNameFor } from "../relationships";
import { Diagnostics, DEFAULT_MAX_DIAGNOSTICS } from "../diagnostics";
import { DEFAULT_COLOR_MAP, createEmptyTheme } from "../theme";
import { createEmptyMediaStore, type MediaStore } from "../media";
import { boxEmu } from "../plaintext/metrics";
import { createSheetPageRenderer } from "../../components/SheetPage";
import type { PageRenderer } from "../../renderer/pageRenderer";
import { formatNumericCell, parseStyles } from "./styles";
import type { SheetCell, SheetPageData } from "./model";

/** Upper bound on grid cells rendered in one workbook page. */
const MAX_CELLS = 10_000;
/** Upper bound on rows and columns, independent of MAX_CELLS. */
const MAX_ROWS = 5_000;
const MAX_COLUMNS = 200;
/** Stop reading a hostile sheet once this many cells have been seen. */
const MAX_PARSED_CELLS = 50_000;
/** Merged ranges are applied in order; this caps the work a hostile file can cause. */
const MAX_MERGES = 2_000;
/** Default Excel column width in characters. */
const DEFAULT_COLUMN_CHARS = 8.43;

export interface XlsxParseOptions {
  limits?: Partial<ArchiveLimits>;
  maxDiagnostics?: number;
}

export interface ParsedXlsx {
  model: PresentationModel;
  media: MediaStore;
  /** One page per worksheet, in workbook order. */
  pageData: readonly SheetPageData[];
  renderPage: PageRenderer;
  dispose(): void;
}

interface SheetRef {
  name: string;
  part: string;
  /** `worksheet`, `chartsheet`, `dialogsheet`, … */
  kind: string;
  hidden: boolean;
}

interface ParsedCell {
  text: string;
  numeric: boolean;
}

interface MergeRange {
  row: number;
  col: number;
  endRow: number;
  endCol: number;
}

interface CoreProperties {
  title: string | null;
  author: string | null;
}

function tryParse(archive: Archive, partName: string): XmlNode | null {
  try {
    const xml = readText(archive, partName);
    return xml === null ? null : parseXml(xml, { part: partName });
  } catch {
    return null;
  }
}

/** Concatenate every text node below an element, in document order. */
function collectText(node: XmlNode | null): string {
  if (!node) return "";
  let out = node.text;
  for (const item of node.children) out += collectText(item);
  return out;
}

function columnIndex(letters: string): number {
  let value = 0;
  for (const character of letters.toUpperCase()) {
    const code = character.charCodeAt(0);
    if (code < 65 || code > 90) return -1;
    value = value * 26 + (code - 64);
  }
  return value - 1;
}

/** Parse `A1` / `$C$12` into zero-based coordinates. */
function parseCellRef(ref: string): { row: number; col: number } | null {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref.trim());
  if (!match) return null;
  const col = columnIndex(match[1] ?? "");
  const row = Number(match[2]);
  if (col < 0 || !Number.isFinite(row) || row < 1) return null;
  return { row: row - 1, col };
}

/** Parse `A1:C4`; a single reference is a one-cell range. */
function parseRange(ref: string): MergeRange | null {
  const [startText, endText] = ref.split(":");
  const start = parseCellRef(startText ?? "");
  const end = parseCellRef(endText ?? startText ?? "");
  if (!start || !end) return null;
  return {
    row: Math.min(start.row, end.row),
    col: Math.min(start.col, end.col),
    endRow: Math.max(start.row, end.row),
    endCol: Math.max(start.col, end.col),
  };
}

function readSharedStrings(archive: Archive): string[] {
  const root = tryParse(archive, "xl/sharedStrings.xml");
  if (!root) return [];
  const values: string[] = [];
  for (const item of children(root, "si")) values.push(collectText(item));
  return values;
}

function readStyles(archive: Archive) {
  return parseStyles(tryParse(archive, "xl/styles.xml"));
}

function readCoreProperties(archive: Archive): CoreProperties {
  const root = tryParse(archive, "docProps/core.xml");
  if (!root) return { title: null, author: null };
  const read = (local: string): string | null => {
    const text = find(root, local)?.text.trim() ?? "";
    return text.length > 0 ? text : null;
  };
  return { title: read("title"), author: read("creator") };
}

function workbookPart(archive: Archive): string {
  const rels = parseRelationships(readText(archive, "_rels/.rels"), "");
  for (const rel of rels.values()) {
    if (relTypeName(rel.type) === "officeDocument" && !rel.external) return rel.target;
  }
  return "xl/workbook.xml";
}

function readWorkbook(archive: Archive, diagnostics: Diagnostics): { sheets: SheetRef[]; date1904: boolean } {
  const part = workbookPart(archive);
  const root = tryParse(archive, part);
  if (!root) {
    throw new DocumentViewerError("MALFORMED_DOCUMENT", "The workbook part could not be read.", {
      detail: part,
    });
  }

  const workbookPr = child(root, "workbookPr");
  const date1904 = attrBool(workbookPr, "date1904") === true;

  let relationships = new Map();
  try {
    relationships = parseRelationships(readText(archive, relsPartNameFor(part)), part);
  } catch {
    diagnostics.add({
      severity: "warning",
      code: "xlsx.relationships",
      message: "The workbook relationships could not be read; worksheets are resolved by their conventional names.",
    });
  }

  const sheets: SheetRef[] = [];
  const sheetList = child(root, "sheets");
  for (const [index, node] of children(sheetList, "sheet").entries()) {
    const name = attr(node, "name") ?? `Sheet ${index + 1}`;
    const relationshipId = relAttr(node, "id");
    const relationship = relationshipId ? relationships.get(relationshipId) : undefined;
    const fallbackPart = `xl/worksheets/sheet${index + 1}.xml`;
    const target = relationship && !relationship.external ? relationship.target : fallbackPart;
    const kind = relationship ? relTypeName(relationship.type) : "worksheet";
    sheets.push({
      name,
      part: target,
      kind,
      hidden: (attr(node, "state") ?? "").toLowerCase() === "hidden",
    });
  }

  if (sheets.length === 0) {
    diagnostics.add({
      severity: "warning",
      code: "xlsx.sheets",
      message: "The workbook declares no sheets.",
    });
  }
  return { sheets, date1904 };
}

function parseCellValue(
  node: XmlNode,
  sharedStrings: readonly string[],
  styles: ReturnType<typeof parseStyles>,
  date1904: boolean,
  warn: (code: string, message: string) => void,
): ParsedCell | null {
  const type = attr(node, "t") ?? "n";
  const styleIndex = attrNum(node, "s") ?? 0;
  const value = child(node, "v")?.text ?? "";
  const inline = child(node, "is");
  const formula = child(node, "f");

  if (type === "inlineStr" || inline) {
    const text = collectText(inline);
    return text.length > 0 ? { text, numeric: false } : null;
  }
  if (type === "s") {
    const index = Number(value);
    const text = Number.isInteger(index) ? sharedStrings[index] : undefined;
    if (text === undefined) {
      warn("xlsx.shared-string", `A shared string with index ${value} is missing.`);
      return null;
    }
    return { text, numeric: false };
  }
  if (type === "b") {
    return { text: value === "1" ? "TRUE" : "FALSE", numeric: false };
  }
  if (type === "e" || type === "str" || type === "d") {
    return value.length > 0 ? { text: value, numeric: false } : null;
  }

  if (value.length === 0) {
    if (formula) {
      warn("xlsx.formula.cache", "A formula has no cached result, so its cell is shown empty.");
    }
    return null;
  }
  const numericValue = Number(value);
  if (!Number.isFinite(numericValue)) {
    warn("xlsx.number", `The value "${value}" is not a number and is shown as text.`);
    return { text: value, numeric: false };
  }
  return {
    text: formatNumericCell(numericValue, styleIndex, styles, date1904),
    numeric: true,
  };
}

function columnWidths(root: XmlNode, count: number): number[] {
  const widths = new Array<number>(count).fill(Math.round(DEFAULT_COLUMN_CHARS * 7.2 + 12));
  const cols = child(root, "cols");
  for (const col of children(cols, "col")) {
    const min = attrNum(col, "min");
    const max = attrNum(col, "max");
    const width = attrNum(col, "width");
    if (min === null || max === null || width === null) continue;
    const px = Math.min(Math.max(Math.round(width * 7.2 + 12), 24), 480);
    for (let index = Math.max(min, 1) - 1; index <= Math.min(max, count) - 1; index += 1) {
      widths[index] = px;
    }
  }
  return widths;
}

function buildSheetPage(
  root: XmlNode,
  name: string,
  sharedStrings: readonly string[],
  styles: ReturnType<typeof parseStyles>,
  date1904: boolean,
  warn: (code: string, message: string) => void,
): SheetPageData {
  const cells = new Map<number, Map<number, ParsedCell>>();
  const merges: MergeRange[] = [];
  let maxRow = 0;
  let maxCol = 0;
  let parsedCells = 0;
  let truncated = false;
  let omittedCells = false;

  const sheetData = child(root, "sheetData");
  let nextRow = 0;
  outer: for (const rowNode of children(sheetData, "row")) {
    const explicitRow = attrNum(rowNode, "r");
    const rowIndex = explicitRow !== null ? Math.max(explicitRow - 1, 0) : nextRow;
    nextRow = rowIndex + 1;
    let nextColumn = 0;
    for (const cellNode of children(rowNode, "c")) {
      const ref = attr(cellNode, "r");
      const parsedRef = ref ? parseCellRef(ref) : null;
      const cellColumn = parsedRef ? parsedRef.col : nextColumn;
      nextColumn = cellColumn + 1;
      if (parsedCells >= MAX_PARSED_CELLS) {
        truncated = true;
        omittedCells = true;
        break outer;
      }
      const parsed = parseCellValue(cellNode, sharedStrings, styles, date1904, warn);
      if (!parsed) continue;
      parsedCells += 1;
      const row = cells.get(rowIndex) ?? new Map<number, ParsedCell>();
      row.set(cellColumn, parsed);
      cells.set(rowIndex, row);
      maxRow = Math.max(maxRow, rowIndex + 1);
      maxCol = Math.max(maxCol, cellColumn + 1);
    }
  }

  const mergeNodes = children(child(root, "mergeCells"), "mergeCell");
  for (const merge of mergeNodes.slice(0, MAX_MERGES)) {
    const range = parseRange(attr(merge, "ref") ?? "");
    if (!range) continue;
    merges.push(range);
    maxRow = Math.max(maxRow, range.endRow + 1);
    maxCol = Math.max(maxCol, range.endCol + 1);
  }
  if (mergeNodes.length > MAX_MERGES) {
    truncated = true;
    warn("xlsx.merge-limit", `Only the first ${MAX_MERGES} merged ranges are shown.`);
  }

  if (maxRow === 0 || maxCol === 0) {
    return { name, rows: [], widths: [], truncated };
  }

  const columnLimit = Math.min(maxCol, MAX_COLUMNS);
  if (maxCol > columnLimit) truncated = true;
  let rowLimit = Math.min(maxRow, MAX_ROWS);
  if (columnLimit > 0 && rowLimit * columnLimit > MAX_CELLS) {
    rowLimit = Math.max(1, Math.floor(MAX_CELLS / columnLimit));
    truncated = true;
  }
  if (maxRow > rowLimit) truncated = true;

  const blankCell = (): SheetCell => ({
    text: "",
    numeric: false,
    covered: false,
    colSpan: 1,
    rowSpan: 1,
    align: "left",
  });
  const matrix: SheetCell[][] = Array.from({ length: rowLimit }, () =>
    Array.from({ length: columnLimit }, blankCell),
  );

  for (const [rowIndex, row] of cells) {
    if (rowIndex >= rowLimit) continue;
    for (const [colIndex, parsed] of row) {
      if (colIndex >= columnLimit) continue;
      const target = matrix[rowIndex]?.[colIndex];
      if (!target) continue;
      target.text = parsed.text;
      target.numeric = parsed.numeric;
      target.align = parsed.numeric ? "right" : "left";
    }
  }

  const covered = new Set<string>();
  for (const merge of merges) {
    if (merge.row >= rowLimit || merge.col >= columnLimit) {
      truncated = true;
      continue;
    }
    const endRow = Math.min(merge.endRow, rowLimit - 1);
    const endCol = Math.min(merge.endCol, columnLimit - 1);
    if (merge.endRow > endRow || merge.endCol > endCol) truncated = true;
    const anchor = matrix[merge.row]?.[merge.col];
    if (!anchor) continue;
    if (covered.has(`${merge.row}:${merge.col}`)) {
      warn("xlsx.merge-overlap", "Two merged ranges overlap; the later one was skipped.");
      continue;
    }
    if (anchor.colSpan > 1 || anchor.rowSpan > 1) {
      warn("xlsx.merge-overlap", "A merged range starts inside another merged range; it was skipped.");
      continue;
    }
    for (let rowIndex = merge.row; rowIndex <= endRow; rowIndex += 1) {
      for (let colIndex = merge.col; colIndex <= endCol; colIndex += 1) {
        if (rowIndex === merge.row && colIndex === merge.col) continue;
        const key = `${rowIndex}:${colIndex}`;
        if (covered.has(key)) continue;
        const target = matrix[rowIndex]?.[colIndex];
        if (!target) continue;
        if (target.text.length > 0) {
          warn("xlsx.merge-content", "A cell with content is covered by a merged range; its text is hidden.");
        }
        target.covered = true;
        target.text = "";
        covered.add(key);
      }
    }
    anchor.colSpan = endCol - merge.col + 1;
    anchor.rowSpan = endRow - merge.row + 1;
  }

  if (omittedCells) {
    warn("xlsx.cell-limit", `Only the first ${MAX_PARSED_CELLS} cells of this sheet were read.`);
  }

  return {
    name,
    rows: matrix,
    widths: columnWidths(root, columnLimit),
    truncated,
  };
}

function emptyPage(index: number, part: string, name: string): SlideModel {
  return {
    index,
    partName: `${part}#page${index + 1}`,
    name,
    slideshowHidden: false,
    shapes: [],
    inheritedShapes: [],
    background: { type: "solid", color: { r: 255, g: 255, b: 255, a: 1 } },
    diagnostics: [],
  };
}

export function parseXlsxArchive(archive: Archive, options: XlsxParseOptions = {}): ParsedXlsx {
  const diagnostics = new Diagnostics(options.maxDiagnostics ?? DEFAULT_MAX_DIAGNOSTICS);
  const { sheets, date1904 } = readWorkbook(archive, diagnostics);
  if (sheets.length === 0) {
    throw new DocumentViewerError("MALFORMED_DOCUMENT", "The workbook does not contain any sheets.", {
      detail: "xl/workbook.xml",
    });
  }

  const sharedStrings = readSharedStrings(archive);
  const styles = readStyles(archive);
  const core = readCoreProperties(archive);

  const pageData: SheetPageData[] = [];
  const slides: SlideModel[] = [];

  sheets.forEach((sheet, index) => {
    const warn = (code: string, message: string) => {
      diagnostics.add({ severity: "warning", code, message, slideIndex: index, part: sheet.part });
    };
    if (sheet.hidden) {
      diagnostics.add({
        severity: "info",
        code: "xlsx.sheet.hidden",
        message: `The sheet "${sheet.name}" is hidden in the workbook and is shown here.`,
        slideIndex: index,
        part: sheet.part,
      });
    }

    let data: SheetPageData = { name: sheet.name, rows: [], widths: [], truncated: false };
    if (sheet.kind !== "worksheet") {
      warn("xlsx.sheet.kind", `The sheet "${sheet.name}" is a ${sheet.kind}, not a grid worksheet, and cannot be displayed.`);
    } else {
      const root = tryParse(archive, sheet.part);
      if (!root) {
        warn("xlsx.sheet.missing", `The worksheet part "${sheet.part}" is missing or unreadable.`);
      } else {
        data = buildSheetPage(root, sheet.name, sharedStrings, styles, date1904, warn);
      }
    }

    pageData.push(data);
    slides.push(emptyPage(index, sheet.part, sheet.name));
  });

  const page = boxEmu("landscape");
  const model: PresentationModel = {
    width: page.width,
    height: page.height,
    slides,
    theme: createEmptyTheme(),
    colorMap: DEFAULT_COLOR_MAP,
    textStyles: { title: [], body: [], other: [] },
    title: core.title,
    author: core.author,
    partCount: archive.names.length,
    diagnostics: diagnostics.list(),
  };

  const media = createEmptyMediaStore();
  return {
    model,
    media,
    pageData,
    renderPage: createSheetPageRenderer(pageData),
    dispose: () => media.dispose(),
  };
}
