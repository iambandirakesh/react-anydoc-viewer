import { describe, expect, it } from "vitest";
import { decodeTextFile, looksLikeText } from "../../src/parser/plaintext/decode";
import { classifyPlainText, sniffDelimiter } from "../../src/parser/plaintext/classify";
import { parseCsv } from "../../src/parser/plaintext/csv";
import { wrapLine, paginateText, MAX_PAGES } from "../../src/parser/plaintext/paginate";
import { charsPerLine, linesPerPage } from "../../src/parser/plaintext/metrics";
import { detectPlainTextKind, parsePlainTextDocument } from "../../src/parser/plaintext";
import { detectFormat, loadDocument } from "../../src/parser/document";

const encoder = new TextEncoder();
const bytes = (text: string): Uint8Array => encoder.encode(text);

describe("text decoding", () => {
  it("decodes UTF-8, with and without a mark", () => {
    expect(decodeTextFile(bytes("héllo"))).toEqual({ text: "héllo", encoding: "utf-8", lossy: false });
    const withBom = new Uint8Array([0xef, 0xbb, 0xbf, ...bytes("hi")]);
    expect(decodeTextFile(withBom)).toEqual({ text: "hi", encoding: "utf-8-bom", lossy: false });
  });

  it("decodes UTF-16, with and without a mark", () => {
    const marked = new Uint8Array([0xff, 0xfe, ...Array.from(bytes("A")).flatMap((b) => [b, 0])]);
    expect(decodeTextFile(marked).text).toBe("A");
    expect(decodeTextFile(marked).encoding).toBe("utf-16le");

    // No mark: the zero-byte pattern has to be recognised instead.
    const unmarked = new Uint8Array(Array.from(bytes("hello")).flatMap((b) => [b, 0]));
    expect(decodeTextFile(unmarked).text).toBe("hello");
    expect(decodeTextFile(unmarked).encoding).toBe("utf-16le");
  });

  it("falls back to Windows-1252 and says so rather than losing characters", () => {
    const latin1 = new Uint8Array([0x63, 0x61, 0x66, 0xe9]); // "café" in Windows-1252
    const decoded = decodeTextFile(latin1);
    expect(decoded.encoding).toBe("windows-1252");
    expect(decoded.lossy).toBe(true);
    expect(decoded.text).toBe("café");
  });

  it("refuses to call binary or whitespace-only bytes text", () => {
    expect(looksLikeText(new Uint8Array([1, 2, 3, 4, 0, 5]))).toBe(false);
    expect(looksLikeText(bytes("\t\t\t"))).toBe(false);
    expect(looksLikeText(bytes("  "))).toBe(false);
    expect(looksLikeText(bytes("real"))).toBe(true);
    expect(detectPlainTextKind(new Uint8Array([1, 2, 3]))).toBeNull();
    expect(detectPlainTextKind(bytes("real"))).toBe("text");
  });
});

describe("classification", () => {
  it("recognises JSON, HTML and Markdown by content", () => {
    expect(classifyPlainText('{"a":[1,2]}').kind).toBe("json");
    expect(classifyPlainText('  {"a":1}  ').kind).toBe("json");
    expect(classifyPlainText("<!doctype html><html></html>").kind).toBe("html");
    expect(classifyPlainText("# Title\n\ntext\n\n# Two").kind).toBe("markdown");
    expect(classifyPlainText("- one\n- two\n\n[link](https://example.com)").kind).toBe("markdown");
  });

  it("does not mistake prose for another format", () => {
    expect(classifyPlainText("Dear Sir,\n\nPlease find the report attached.\n").kind).toBe("text");
    expect(classifyPlainText("* a note about a star").kind).toBe("text");
    expect(classifyPlainText('{"broken": ').kind).toBe("text");
  });

  it("picks the delimiter with the most consistent field count", () => {
    expect(sniffDelimiter("a,b\n1,2\n3,4")?.delimiter).toBe(",");
    expect(sniffDelimiter("a;b\n1;2\n3;4")?.delimiter).toBe(";");
    expect(sniffDelimiter("a\tb\n1\t2")?.delimiter).toBe("\t");
    // A comma inside a quoted field is not a separator.
    expect(sniffDelimiter('a;b\n"x,y";2\n"z";3')?.delimiter).toBe(";");
    expect(sniffDelimiter("just one line")).toBeNull();
    expect(sniffDelimiter("a,b\nc\n")).toBeNull();
  });
});

describe("csv parsing", () => {
  it("keeps quoted fields, escaped quotes and embedded newlines", () => {
    const table = parseCsv('name,note\n"Smith, Jane","said ""hi""\nsecond line"\n');
    expect(table.header).toEqual(["name", "note"]);
    expect(table.rows).toEqual([["Smith, Jane", 'said "hi"\nsecond line']]);
    expect(table.raggedRows).toBe(0);
  });

  it("handles CRLF, a missing final newline and a trailing newline alike", () => {
    expect(parseCsv("a,b\r\n1,2\r\n").rows).toEqual([["1", "2"]]);
    expect(parseCsv("a,b\n1,2").rows).toEqual([["1", "2"]]);
    expect(parseCsv("a,b\n1,2\n").rows).toEqual([["1", "2"]]);
  });

  it("counts rows that do not match the header instead of padding them", () => {
    const table = parseCsv("a,b,c\n1,2\n3,4,5\n");
    expect(table.raggedRows).toBe(1);
    expect(table.rows[0]).toEqual(["1", "2"]);
  });

  it("keeps values as text", () => {
    const table = parseCsv("id,date\n0123,2024-01-02\n");
    expect(table.rows[0]).toEqual(["0123", "2024-01-02"]);
  });

  it("truncates at the row limit and says how many rows there are", () => {
    const table = parseCsv("a\n1\n2\n3\n4\n", { maxRows: 2 });
    expect(table.rows).toHaveLength(2);
    expect(table.totalRows).toBe(4);
    expect(table.truncated).toBe(true);
  });
});

describe("text pagination", () => {
  it("wraps a long line at a break, and cuts a line with no break at all", () => {
    expect(wrapLine("short", 20)).toEqual(["short"]);
    expect(wrapLine("alpha beta gamma delta", 12)).toEqual(["alpha beta ", "gamma delta"]);
    expect(wrapLine("x".repeat(25), 10)).toEqual(["x".repeat(10), "x".repeat(10), "x".repeat(5)]);
  });

  it("fills pages to the page's line count and keeps the remainder", () => {
    const perPage = linesPerPage("portrait");
    const result = paginateText(Array.from({ length: perPage + 3 }, (_, i) => `line ${i}`).join("\n"));
    expect(result.pages).toHaveLength(2);
    expect(result.pages[0]?.lines).toHaveLength(perPage);
    expect(result.pages[1]?.lines).toHaveLength(3);
    expect(result.complete).toBe(true);
  });

  it("wraps to the page's column count so nothing is drawn off the page", () => {
    const columns = charsPerLine("portrait");
    const result = paginateText("y".repeat(columns * 2));
    expect(result.pages[0]?.lines).toEqual(["y".repeat(columns), "y".repeat(columns)]);
    expect(result.wrappedLines).toBe(1);
  });

  it("stops at the page ceiling and reports the truncation", () => {
    const perPage = linesPerPage("portrait");
    const result = paginateText(
      Array.from({ length: (MAX_PAGES + 1) * perPage }, (_, i) => `${i}`).join("\n"),
      { maxPages: 2 },
    );
    expect(result.pages).toHaveLength(2);
    expect(result.complete).toBe(false);
  });

  it("sanitises control characters and expands tabs", () => {
    const result = paginateText("a\u0001b\tc");
    expect(result.pages[0]?.lines).toEqual(["ab  c"]);
  });
});

describe("text documents", () => {
  it("detects a text file and lays it out one page per chunk", () => {
    expect(detectFormat(bytes("hello, this is a note"))).toBe("text");
    const parsed = parsePlainTextDocument(bytes("hello"));
    expect(parsed.format).toBe("text");
    expect(parsed.model.slides).toHaveLength(1);
    // A page is a whole sheet of paper, not the size of the text on it.
    expect(parsed.model.width).toBe(Math.round(816 * 9525));
    expect(parsed.renderPage).toBeTypeOf("function");
    parsed.dispose();
  });

  it("pretty-prints JSON and reports the encoding and classification", () => {
    const parsed = parsePlainTextDocument(bytes('{"a":1,"b":[2,3]}'));
    expect(parsed.format).toBe("text");
    const codes = parsed.model.diagnostics.map((diagnostic) => diagnostic.code);
    expect(codes).toContain("text.classified.json");
    parsed.dispose();
  });

  it("shows HTML as source and warns that nothing is executed", () => {
    const parsed = parsePlainTextDocument(bytes("<!doctype html><html><script>alert(1)</script></html>"));
    expect(parsed.format).toBe("text");
    expect(parsed.model.diagnostics.map((diagnostic) => diagnostic.code)).toContain("text.html.source");
    // Read as text like any other file: the markup reaches the model as characters,
    // and no renderer of its own is involved.
    expect(parsed.model.slides).toHaveLength(1);
    parsed.dispose();
  });

  it("builds a table page per chunk of rows, repeating the header", () => {
    const rows = Array.from({ length: 60 }, (_, i) => `${i},value ${i}`).join("\n");
    const parsed = parsePlainTextDocument(bytes(`id,label\n${rows}\n`));
    expect(parsed.format).toBe("csv");
    expect(parsed.model.slides.length).toBeGreaterThan(1);
    const last = parsed.model.slides.length - 1;
    expect(parsed.model.slides[last]?.index).toBe(last);
    parsed.dispose();
  });

  it("uses a landscape page for a wide table and reports clipping", () => {
    const header = Array.from({ length: 30 }, (_, i) => `column ${i}`).join(",");
    const row = Array.from({ length: 30 }, () => "value").join(",");
    const wide = parsePlainTextDocument(bytes(`${header}\n${row}\n`));
    expect(wide.model.width).toBe(Math.round(1056 * 9525));
    wide.dispose();

    const huge = parsePlainTextDocument(
      bytes(
        `${Array.from({ length: 80 }, (_, i) => `c${i}`.padEnd(40, "x")).join(",")}\n${Array.from({ length: 80 }, () =>
          "v".repeat(40),
        ).join(",")}\n`,
      ),
    );
    expect(huge.model.diagnostics.map((diagnostic) => diagnostic.code)).toContain("text.csv.width");
    huge.dispose();
  });

  it("honours an explicit content hint over the content guess", () => {
    const parsed = parsePlainTextDocument(bytes("a,b\n1,2\n"), { contentHint: "text" });
    expect(parsed.format).toBe("text");
    parsed.dispose();
  });

  it("loads through the document entry point with the format from the bytes", async () => {
    const csv = await loadDocument(bytes("a,b\n1,2\n3,4\n"));
    expect(csv.format).toBe("csv");
    expect(csv.renderPage).toBeTypeOf("function");
    csv.dispose();

    const text = await loadDocument(bytes("just a note\n"));
    expect(text.format).toBe("text");
    text.dispose();
  });
});
