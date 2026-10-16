/**
 * CSV parsing.
 *
 * RFC 4180 describes a common layout, it does not guarantee one, so this is a
 * character-by-character state machine rather than a `split(",")`: a quoted field
 * may contain the delimiter, doubled quotes and line breaks, and records may end
 * with CRLF, LF or CR. Rows are reported as found — a short or long row is counted
 * and reported, never silently padded or shifted.
 *
 * Values stay text. The viewer never coerces "0123" into 123 or a date-like
 * string into a date, because a viewer that reformats data is a viewer that lies
 * about it.
 */

import { describeDelimiter, sniffDelimiter } from "./classify";

export interface CsvTable {
  header: string[];
  rows: string[][];
  delimiter: string;
  /** true when the delimiter was chosen by content rather than supplied */
  sniffed: boolean;
  /** rows that did not have the header's field count */
  raggedRows: number;
  /** rows in the file, including any beyond the row limit */
  totalRows: number;
  /** true when rows were dropped to respect the row limit */
  truncated: boolean;
  /** field count of the first record */
  columns: number;
}

export interface CsvParseOptions {
  delimiter?: string;
  /** maximum data rows kept in the model (default 20 000) */
  maxRows?: number;
  /** maximum characters parsed (default 8 MiB) */
  maxChars?: number;
}

const DEFAULT_MAX_ROWS = 20_000;
const DEFAULT_MAX_CHARS = 8 * 1024 * 1024;

/** All records in the text, honouring quotes. */
function splitRecords(text: string, delimiter: string): string[][] {
  const records: string[][] = [];
  let fields: string[] = [];
  let current = "";
  let quoted = false;
  // True once the record being read has any content at all, so a file that ends
  // in a line break does not gain a final empty record, while a final record that
  // is an empty *quoted* field is still kept.
  let recordStarted = false;

  for (let i = 0; i < text.length; i += 1) {
    const character = text[i] as string;

    if (quoted) {
      if (character === '"') {
        if (text[i + 1] === '"') {
          current += '"';
          i += 1;
        } else {
          quoted = false;
        }
      } else {
        current += character;
      }
      continue;
    }

    if (character === '"') {
      quoted = true;
      recordStarted = true;
      continue;
    }
    if (character === delimiter) {
      fields.push(current);
      current = "";
      recordStarted = true;
      continue;
    }
    if (character === "\n" || character === "\r") {
      if (character === "\r" && text[i + 1] === "\n") i += 1;
      if (recordStarted || current.length > 0 || fields.length > 0) {
        fields.push(current);
        records.push(fields);
      }
      fields = [];
      current = "";
      recordStarted = false;
      continue;
    }
    current += character;
    recordStarted = true;
  }

  // A file that does not end in a line break still has a final record; one that
  // does must not gain an empty one.
  if (recordStarted || current.length > 0 || fields.length > 0) {
    fields.push(current);
    records.push(fields);
  }
  return records;
}

export function parseCsv(text: string, options: CsvParseOptions = {}): CsvTable {
  const maxRows = options.maxRows ?? DEFAULT_MAX_ROWS;
  const maxChars = options.maxChars ?? DEFAULT_MAX_CHARS;

  const crop = text.length > maxChars ? text.slice(0, maxChars) : text;
  const sniffed = options.delimiter ? null : sniffDelimiter(crop);
  const delimiter = options.delimiter ?? sniffed?.delimiter ?? ",";

  const records = splitRecords(crop, delimiter);
  const header = records[0] ?? [];
  const columns = header.length;
  const dataRecords = records.slice(1);

  const raggedRows = dataRecords.filter((record) => record.length !== columns).length;
  const kept = dataRecords.slice(0, maxRows);

  return {
    header,
    rows: kept,
    delimiter,
    sniffed: options.delimiter === undefined && sniffed !== null,
    raggedRows,
    totalRows: dataRecords.length,
    truncated: dataRecords.length > kept.length || text.length > crop.length,
    columns,
  };
}

/** The sentence used when a row count is reported. */
export function raggedRowMessage(table: CsvTable): string | null {
  if (table.raggedRows === 0) return null;
  return `${table.raggedRows} row(s) do not have the header's ${table.columns} field(s); their values are shown as they are.`;
}

/** How the delimiter was decided, for the diagnostics panel. */
export function delimiterMessage(table: CsvTable): string {
  return table.sniffed
    ? `The delimiter was detected from the content (${describeDelimiter(table.delimiter)}).`
    : `Records are split on ${describeDelimiter(table.delimiter)}.`;
}
