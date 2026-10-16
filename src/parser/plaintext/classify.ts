/**
 * What a plain-text file actually is.
 *
 * The library's rule everywhere else is that the bytes decide and the file name
 * is ignored. That works because `.pptx`, `.docx` and `.pdf` are containers with
 * signatures. Six of the text types are not: `.txt`, `.log`, `.csv`, `.json`,
 * `.md` and `.html` are all "bytes that decode as text", and no signature tells
 * them apart.
 *
 * So this module classifies by content, in order of how specific the evidence is,
 * and a caller that knows better can override the answer with `contentHint`. The
 * name is still never consulted; the hint only changes how a file that is already
 * known to be text is presented.
 */

import { looksLikeText, decodeTextFile } from "./decode";

export type PlainTextKind = "json" | "html" | "markdown" | "csv" | "text";

/** Fields that are considered when sniffing a delimiter. */
export const DELIMITERS: readonly string[] = [",", ";", "\t", "|"];

/** Rows inspected when judging whether a delimiter is consistent. */
const DELIMITER_SAMPLE_ROWS = 12;

export interface Classification {
  kind: PlainTextKind;
  /** the delimiter, for a CSV */
  delimiter?: string;
  /** why this answer was given, for diagnostics */
  reason: string;
  /** 0..1 — how sure the classifier is, used to word the diagnostics */
  confidence: number;
}

function looksLikeJson(text: string): boolean {
  const head = text.trimStart();
  if (!head.startsWith("{") && !head.startsWith("[")) return false;
  const end = text.trimEnd();
  if (!end.endsWith("}") && !end.endsWith("]")) return false;
  try {
    JSON.parse(text);
    return true;
  } catch {
    return false;
  }
}

function looksLikeHtml(text: string): boolean {
  return /^\s*(<!doctype\s+html|<!doctype|<html[\s>]|<head[\s>]|<body[\s>]|<\?xml[^>]*>\s*<!doctype)/i.test(
    text,
  );
}

/**
 * Markdown is recognised by its block-level markers rather than by inline
 * punctuation, because a `.txt` note full of asterisks is not a document. A
 * heading, a fence or two list items is enough; a single `-` line is not.
 */
function looksLikeMarkdown(text: string): boolean {
  const lines = text.split(/\r?\n/, 200);
  let headings = 0;
  let fences = 0;
  let listItems = 0;
  let links = 0;
  for (const line of lines) {
    if (/^\s{0,3}#{1,6}\s+\S/.test(line)) headings += 1;
    else if (/^\s{0,3}(```|~~~)/.test(line)) fences += 1;
    else if (/^\s{0,3}([-*+]|\d{1,9}[.)])\s+\S/.test(line)) listItems += 1;
    else if (/\[[^\]]+\]\([^)]+\)/.test(line)) links += 1;
  }
  const signals = headings + (fences >= 2 ? 1 : 0) + (listItems >= 2 ? 1 : 0) + (links >= 1 ? 1 : 0);
  return signals >= 2 || headings >= 2 || (headings >= 1 && (listItems >= 2 || links >= 1));
}

/**
 * Count delimiters outside quoted fields. A CSV's defining property is not that
 * it has commas, it is that every record has the same shape.
 */
function countDelimitersOutsideQuotes(line: string, delimiter: string): number {
  let count = 0;
  let quoted = false;
  for (let i = 0; i < line.length; i += 1) {
    const character = line[i];
    if (character === '"') {
      if (quoted && line[i + 1] === '"') {
        i += 1;
        continue;
      }
      quoted = !quoted;
      continue;
    }
    if (!quoted && character === delimiter) count += 1;
  }
  return count;
}

/**
 * Split a sample of lines into records first, so a quoted field containing a
 * newline does not look like a broken row.
 */
function sampleRecords(text: string, rows: number): string[] {
  const records: string[] = [];
  let current = "";
  let quoted = false;
  for (let i = 0; i < text.length && records.length < rows; i += 1) {
    const character = text[i] as string;
    if (character === '"') {
      quoted = !quoted;
      current += character;
      continue;
    }
    if (!quoted && (character === "\n" || character === "\r")) {
      if (character === "\r" && text[i + 1] === "\n") i += 1;
      records.push(current);
      current = "";
      continue;
    }
    current += character;
  }
  if (current.length > 0 && records.length < rows) records.push(current);
  return records;
}

export interface DelimiterChoice {
  delimiter: string;
  /** number of fields per record when the delimiter is used */
  fields: number;
  /** records that agreed with that field count */
  agreed: number;
  /** records inspected */
  sampled: number;
}

/** Pick the delimiter whose field count is the most consistent across records. */
export function sniffDelimiter(text: string): DelimiterChoice | null {
  const records = sampleRecords(text, DELIMITER_SAMPLE_ROWS);
  if (records.length < 2) return null;

  let best: DelimiterChoice | null = null;
  for (const delimiter of DELIMITERS) {
    const counts = records.map((record) => countDelimitersOutsideQuotes(record, delimiter));
    const fields = (counts[0] as number) + 1;
    if (fields < 2) continue;
    const agreed = counts.filter((count) => count === fields - 1).length;
    // A candidate must describe most of the sample and every record must contain
    // the delimiter at least once (a trailing empty field is still a field).
    if (agreed / records.length < 0.8) continue;
    if (counts.some((count) => count === 0)) continue;
    if (!best || agreed > best.agreed || (agreed === best.agreed && fields > best.fields)) {
      best = { delimiter, fields, agreed, sampled: records.length };
    }
  }
  return best;
}

export function classifyPlainText(text: string): Classification {
  if (looksLikeJson(text)) {
    return { kind: "json", reason: "the content parses as JSON", confidence: 0.99 };
  }
  if (looksLikeHtml(text)) {
    return { kind: "html", reason: "the content starts with an HTML document marker", confidence: 0.9 };
  }
  if (looksLikeMarkdown(text)) {
    return { kind: "markdown", reason: "the content has Markdown block markers", confidence: 0.7 };
  }
  const delimiter = sniffDelimiter(text);
  if (delimiter) {
    return {
      kind: "csv",
      delimiter: delimiter.delimiter,
      reason: `every record in the sample holds ${delimiter.fields} fields separated by ${describeDelimiter(
        delimiter.delimiter,
      )}`,
      confidence: delimiter.agreed === delimiter.sampled ? 0.9 : 0.75,
    };
  }
  return { kind: "text", reason: "no other text format matched", confidence: 0.5 };
}

export function describeDelimiter(delimiter: string): string {
  switch (delimiter) {
    case "\t":
      return "a tab";
    case ",":
      return "a comma";
    case ";":
      return "a semicolon";
    case "|":
      return "a pipe";
    default:
      return `"${delimiter}"`;
  }
}

/**
 * Classify a file from its bytes.
 *
 * Only the head of the file is inspected, so classification cost does not grow
 * with file size; a delimiter or heading that first appears after 64 KiB is not
 * what decides the format.
 */
export function classifyTextBytes(bytes: Uint8Array, headLimit = 64 * 1024): Classification {
  if (!looksLikeText(bytes)) {
    return { kind: "text", reason: "the bytes are not text", confidence: 0 };
  }
  const head = bytes.subarray(0, Math.min(bytes.length, headLimit));
  return classifyPlainText(decodeTextFile(head).text);
}
