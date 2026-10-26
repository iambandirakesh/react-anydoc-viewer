/**
 * Spreadsheet style and number-format helpers.
 *
 * An `.xlsx` stores a date as a number plus a number format, so a value cannot be
 * displayed correctly without resolving the cell's `s` index through
 * `cellXfs` and then the format code. The viewer implements the common built-in
 * formats and a conservative subset of custom ones; anything else is shown as
 * the stored value rather than guessed at.
 */

import { children, attrNum, type XmlNode } from "../xml";

/** Built-in Excel number-format ids that represent dates or times. */
const BUILT_IN_DATE_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);

/** Built-in id -> format code, for the subset the viewer can interpret. */
const BUILT_IN_CODES: Record<number, string> = {
  9: "0%",
  10: "0.00%",
  14: "mm-dd-yy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yy h:mm",
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mmss.0",
};

export interface WorkbookStyles {
  /** cellXfs index -> numFmtId */
  numFmtByStyle: number[];
  /** custom numFmtId -> formatCode */
  formatCodeById: Map<number, string>;
}

export function parseStyles(root: XmlNode | null): WorkbookStyles {
  const numFmtByStyle: number[] = [];
  const formatCodeById = new Map<number, string>();
  if (!root) return { numFmtByStyle, formatCodeById };

  const numFmts = children(root, "numFmts")[0] ?? null;
  for (const numFmt of children(numFmts, "numFmt")) {
    const id = attrNum(numFmt, "numFmtId");
    const code = numFmt.attrs["formatCode"];
    if (id !== null && code !== undefined) formatCodeById.set(id, code);
  }

  const cellXfs = children(root, "cellXfs")[0] ?? null;
  for (const xf of children(cellXfs, "xf")) {
    numFmtByStyle.push(attrNum(xf, "numFmtId") ?? 0);
  }
  return { numFmtByStyle, formatCodeById };
}

export function formatCodeFor(styleIndex: number, styles: WorkbookStyles): string | null {
  const numFmtId = styles.numFmtByStyle[styleIndex] ?? 0;
  return styles.formatCodeById.get(numFmtId) ?? BUILT_IN_CODES[numFmtId] ?? null;
}

export type DateFormatKind = "date" | "time" | "datetime";

/**
 * Decide whether a format code is a date/time format.
 *
 * Literal sections (quoted text, escaped characters and bracketed colour/locale
 * modifiers) are removed first, because `"Days" 0` must not be mistaken for a
 * date and `[h]:mm` must still be one.
 */
export function dateKindFor(formatCode: string | null, numFmtId: number | null): DateFormatKind | null {
  if (numFmtId !== null && BUILT_IN_DATE_IDS.has(numFmtId)) {
    if (numFmtId >= 18 && numFmtId <= 21) return "time";
    if (numFmtId === 45 || numFmtId === 46 || numFmtId === 47) return "time";
    if (numFmtId === 22) return "datetime";
    return "date";
  }
  if (!formatCode) return null;
  const stripped = formatCode
    .replace(/"[^"]*"/g, "")
    .replace(/\[[^\]]*\]/g, (token) => (/[hms]/i.test(token) ? token : ""))
    .replace(/\\[\s\S]/g, "");
  const hasDate = /[ymd]/i.test(stripped);
  const hasTime = /[hs]/i.test(stripped) || /AM\/PM/i.test(stripped);
  if (hasDate && hasTime) return "datetime";
  if (hasDate) return "date";
  if (hasTime) return "time";
  return null;
}

function pad(value: number): string {
  return String(value).padStart(2, "0");
}

function formatClock(milliseconds: number, includeSeconds: boolean): string {
  const totalSeconds = Math.floor(milliseconds / 1000);
  const hours = Math.floor(totalSeconds / 3600) % 24;
  const minutes = Math.floor(totalSeconds / 60) % 60;
  const seconds = totalSeconds % 60;
  return includeSeconds ? `${pad(hours)}:${pad(minutes)}:${pad(seconds)}` : `${pad(hours)}:${pad(minutes)}`;
}

/**
 * Excel's serial date to an ISO-ish string.
 *
 * The 1900 system does not contain 1900-02-29, but Excel reserves serial 60 for
 * it. Serial 60 is therefore displayed as that impossible date rather than being
 * silently mapped onto 28 February or 1 March, which would make a workbook
 * disagree with Excel by a day for every later date.
 */
export function formatSerialDate(
  serial: number,
  kind: DateFormatKind,
  date1904: boolean,
  includeSeconds = false,
): string {
  const whole = Math.floor(serial);
  const fraction = serial - whole;
  if (!date1904 && whole === 60) return kind === "time" ? "00:00:00" : "1900-02-29";
  const base = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  const dayOffset = !date1904 && whole < 60 ? whole + 1 : whole;
  const date = new Date(base + dayOffset * 86_400_000 + Math.round(fraction * 86_400_000));
  const dateText = `${date.getUTCFullYear()}-${pad(date.getUTCMonth() + 1)}-${pad(date.getUTCDate())}`;
  const timeText = formatClock(fraction * 86_400_000, includeSeconds);
  if (kind === "date") return dateText;
  if (kind === "time") return timeText;
  return `${dateText} ${timeText}`;
}

function decimalPlaces(formatCode: string): number {
  const match = /\.(0+|#+)/.exec(formatCode);
  return match ? (match[1] ?? "").length : 0;
}

function formatNumber(value: number, formatCode: string | null, numFmtId: number | null): string {
  if (!Number.isFinite(value)) return "";
  const code = formatCode ?? (numFmtId !== null ? BUILT_IN_CODES[numFmtId] ?? null : null);
  if (!code) return String(value);

  if (code.includes("%") || numFmtId === 9 || numFmtId === 10) {
    const decimals = decimalPlaces(code);
    return `${(value * 100).toFixed(decimals)}%`;
  }

  const decimals = decimalPlaces(code);
  if (decimals > 0) {
    const fixed = value.toFixed(decimals);
    // A thousands separator is a display choice in the format code, not in the
    // value. Only apply it when the code asks for one.
    if (code.includes(",")) {
      const [integer, fraction] = fixed.split(".");
      const grouped = (integer ?? "0").replace(/\B(?=(\d{3})+(?!\d))/g, ",");
      return fraction === undefined ? grouped : `${grouped}.${fraction}`;
    }
    return fixed;
  }

  return String(value);
}

/**
 * Render the cached value of a numeric cell using its style.
 *
 * `formatCode` may be null for a style that carries only the default `General`
 * format; the number is then shown as stored.
 */
export function formatNumericCell(
  value: number,
  styleIndex: number,
  styles: WorkbookStyles,
  date1904: boolean,
): string {
  const numFmtId = styles.numFmtByStyle[styleIndex] ?? 0;
  const formatCode = formatCodeFor(styleIndex, styles);
  const kind = dateKindFor(formatCode, numFmtId);
  if (kind) return formatSerialDate(value, kind, date1904, formatCode?.includes("s") ?? false);
  return formatNumber(value, formatCode, numFmtId);
}
