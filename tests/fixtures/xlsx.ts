/**
 * XLSX fixture builder.
 *
 * Workbooks are assembled from real SpreadsheetML parts with `fflate`, the same
 * way the presentation, document and PDF fixtures are generated. The builder
 * covers shared strings, inline strings, numeric cells, booleans, error values,
 * dashed formulas with cached results, dates, percentages, merged ranges and
 * column widths, so the parser is exercised with genuine OOXML rather than a
 * hand-shaped mock.
 */

import { strToU8, zipSync } from "fflate";

export type XlsxCell =
  | string
  | number
  | boolean
  | null
  | { type: "inline"; value: string }
  | { type: "formula"; formula: string; value: string | number | null }
  | { type: "date"; serial: number }
  | { type: "datetime"; serial: number }
  | { type: "percent"; value: number }
  | { type: "error"; value: string };

export interface XlsxSheetSpec {
  name: string;
  rows: (readonly (XlsxCell | undefined)[])[];
  merges?: readonly string[];
  cols?: readonly { min: number; max: number; width: number }[];
}

export interface XlsxFixtureOptions {
  sheets: readonly XlsxSheetSpec[];
  date1904?: boolean;
  title?: string;
  creator?: string;
}

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";

export function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function columnName(index: number): string {
  let value = index + 1;
  let out = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    out = String.fromCharCode(65 + remainder) + out;
    value = Math.floor((value - 1) / 26);
  }
  return out;
}

function cellRef(row: number, col: number): string {
  return `${columnName(col)}${row + 1}`;
}

function cellXml(
  cell: XlsxCell | undefined,
  row: number,
  col: number,
  sharedStrings: string[],
  sharedIndex: Map<string, number>,
): string {
  if (cell === null || cell === undefined) return "";
  const ref = cellRef(row, col);
  const shared = (text: string): string => {
    let index = sharedIndex.get(text);
    if (index === undefined) {
      index = sharedStrings.length;
      sharedStrings.push(text);
      sharedIndex.set(text, index);
    }
    return `<c r="${ref}" t="s"><v>${index}</v></c>`;
  };

  if (typeof cell === "string") return shared(cell);
  if (typeof cell === "number") return `<c r="${ref}"><v>${cell}</v></c>`;
  if (typeof cell === "boolean") return `<c r="${ref}" t="b"><v>${cell ? 1 : 0}</v></c>`;

  switch (cell.type) {
    case "inline":
      return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(cell.value)}</t></is></c>`;
    case "formula": {
      const value = cell.value === null ? "" : `<v>${escapeXml(String(cell.value))}</v>`;
      return `<c r="${ref}"><f>${escapeXml(cell.formula)}</f>${value}</c>`;
    }
    case "date":
      return `<c r="${ref}" s="1"><v>${cell.serial}</v></c>`;
    case "datetime":
      return `<c r="${ref}" s="3"><v>${cell.serial}</v></c>`;
    case "percent":
      return `<c r="${ref}" s="2"><v>${cell.value}</v></c>`;
    case "error":
      return `<c r="${ref}" t="e"><v>${escapeXml(cell.value)}</v></c>`;
    default:
      return "";
  }
}

function sheetXml(sheet: XlsxSheetSpec, sharedStrings: string[], sharedIndex: Map<string, number>): string {
  const rowXml: string[] = [];
  sheet.rows.forEach((row, rowIndex) => {
    const cells = row
      .map((cell, colIndex) => cellXml(cell, rowIndex, colIndex, sharedStrings, sharedIndex))
      .join("");
    if (cells.length > 0) rowXml.push(`<row r="${rowIndex + 1}">${cells}</row>`);
  });

  let lastRow = 0;
  let lastCol = 0;
  sheet.rows.forEach((row, rowIndex) => {
    row.forEach((cell, colIndex) => {
      if (cell === null || cell === undefined) return;
      lastRow = Math.max(lastRow, rowIndex + 1);
      lastCol = Math.max(lastCol, colIndex + 1);
    });
  });
  for (const merge of sheet.merges ?? []) {
    const end = merge.split(":")[1] ?? merge;
    const match = /^([A-Z]+)(\d+)$/.exec(end);
    if (!match) continue;
    lastCol = Math.max(lastCol, [...(match[1] ?? "")].reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0));
    lastRow = Math.max(lastRow, Number(match[2]));
  }

  const dimension = lastRow > 0 && lastCol > 0 ? `A1:${cellRef(lastRow - 1, lastCol - 1)}` : "A1";
  const cols = (sheet.cols ?? [])
    .map((col) => `<col min="${col.min}" max="${col.max}" width="${col.width}" customWidth="1"/>`)
    .join("");
  const merges = (sheet.merges ?? [])
    .map((ref) => `<mergeCell ref="${ref}"/>`)
    .join("");

  return `${XML_HEADER}<worksheet xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}"><dimension ref="${dimension}"/>${
    cols.length > 0 ? `<cols>${cols}</cols>` : ""
  }<sheetData>${rowXml.join("")}</sheetData>${
    merges.length > 0 ? `<mergeCells count="${(sheet.merges ?? []).length}">${merges}</mergeCells>` : ""
  }</worksheet>`;
}

const STYLES_XML = `${XML_HEADER}<styleSheet xmlns="${SPREADSHEET_NS}"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm"/></numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>`;

/** Build a complete `.xlsx` package. */
export function buildXlsx(options: XlsxFixtureOptions): Uint8Array {
  const sharedStrings: string[] = [];
  const sharedIndex = new Map<string, number>();
  const sheets = options.sheets.length > 0 ? options.sheets : [{ name: "Sheet1", rows: [] }];

  const sheetParts: Record<string, Uint8Array> = {};
  sheets.forEach((sheet, index) => {
    sheetParts[`xl/worksheets/sheet${index + 1}.xml`] = strToU8(sheetXml(sheet, sharedStrings, sharedIndex));
  });

  const sheetEntries = sheets
    .map(
      (sheet, index) =>
        `<sheet name="${escapeXml(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`,
    )
    .join("");
  const sharedStringsPart = `${XML_HEADER}<sst xmlns="${SPREADSHEET_NS}" count="${sharedStrings.length}" uniqueCount="${sharedStrings.length}">${sharedStrings
    .map((value) => `<si><t xml:space="preserve">${escapeXml(value)}</t></si>`)
    .join("")}</sst>`;
  const sharedRid = `rId${sheets.length + 1}`;
  const stylesRid = `rId${sheets.length + 2}`;

  const workbook = `${XML_HEADER}<workbook xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}"><workbookPr date1904="${
    options.date1904 ? "1" : "0"
  }"/><sheets>${sheetEntries}</sheets></workbook>`;
  const workbookRels = `${XML_HEADER}<Relationships xmlns="${PACKAGE_REL_NS}">${sheets
    .map(
      (_, index) =>
        `<Relationship Id="rId${index + 1}" Type="${REL_NS}/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`,
    )
    .join("")}<Relationship Id="${sharedRid}" Type="${REL_NS}/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="${stylesRid}" Type="${REL_NS}/styles" Target="styles.xml"/></Relationships>`;
  const rootRels = `${XML_HEADER}<Relationships xmlns="${PACKAGE_REL_NS}"><Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`;
  const core = `${XML_HEADER}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${escapeXml(
    options.title ?? "Fixture workbook",
  )}</dc:title><dc:creator>${escapeXml(options.creator ?? "Fixture author")}</dc:creator></cp:coreProperties>`;
  const contentTypes = `${XML_HEADER}<Types xmlns="${CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${sheets
    .map(
      (_, index) =>
        `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
    )
    .join("")}<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`;

  return zipSync(
    {
      "[Content_Types].xml": strToU8(contentTypes),
      "_rels/.rels": strToU8(rootRels),
      "docProps/core.xml": strToU8(core),
      "xl/workbook.xml": strToU8(workbook),
      "xl/_rels/workbook.xml.rels": strToU8(workbookRels),
      "xl/sharedStrings.xml": strToU8(sharedStringsPart),
      "xl/styles.xml": strToU8(STYLES_XML),
      ...sheetParts,
    },
    { level: 6 },
  );
}
