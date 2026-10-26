import { describe, expect, it } from "vitest";
import { probePackage, sniffContainer, unsupportedFormatMessage } from "../../src/parser/format";
import { detectFormat } from "../../src/parser/document";
import { buildDocx, paragraph } from "../fixtures/docx";
import { buildPdf } from "../fixtures/pdf";
import { buildPptx, NS, paragraph as pptParagraph, run, sp } from "../fixtures/pptx";
import { strToU8, zipSync } from "fflate";

const bytes = (...values: number[]) => new Uint8Array(values);

describe("sniffContainer", () => {
  it("recognises a PDF, including one with bytes before the header", () => {
    expect(sniffContainer(strToU8("%PDF-1.7\n…"))).toBe("pdf");
    expect(sniffContainer(strToU8("\uFEFF  %PDF-1.4\n"))).toBe("pdf");
    expect(sniffContainer(buildPdf({ pages: ["one"] }))).toBe("pdf");
  });

  it("recognises ZIP packages and OLE compound files", () => {
    expect(sniffContainer(bytes(0x50, 0x4b, 0x03, 0x04, 0x00))).toBe("zip");
    expect(sniffContainer(bytes(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0x00))).toBe("ole");
  });

  it("treats anything else as unknown", () => {
    expect(sniffContainer(new Uint8Array(0))).toBe("unknown");
    expect(sniffContainer(strToU8("<html></html>"))).toBe("unknown");
    expect(sniffContainer(bytes(0x50, 0x4b, 0x02, 0x04))).toBe("unknown");
  });
});

describe("probePackage", () => {
  it("identifies the format from the package's main part", () => {
    expect(probePackage(["ppt/presentation.xml", "ppt/slides/slide1.xml"]).format).toBe("pptx");
    expect(probePackage(["word/document.xml", "word/styles.xml"]).format).toBe("docx");
    // Part names are case-insensitive, like OPC part-name matching.
    expect(probePackage(["Word/Document.xml"]).format).toBe("docx");
  });

  it("identifies spreadsheets and names formats it cannot render instead of guessing", () => {
    expect(probePackage(["xl/workbook.xml"]).format).toBe("xlsx");
    expect(probePackage(["xl/workbook.bin"]).label).toContain(".xlsb");
    expect(probePackage(["xl/workbook.bin"]).format).toBeNull();
    expect(probePackage(["mimetype", "content.xml"]).label).toContain("OpenDocument");
    expect(probePackage(["foo/bar.xml"]).label).toContain("unfamiliar");
  });

  it("writes a message that names the supported formats", () => {
    const message = unsupportedFormatMessage("an Excel workbook (.xlsx)");
    expect(message).toContain(".pptx");
    expect(message).toContain(".docx");
    expect(message).toContain(".pdf");
  });
});

describe("detectFormat", () => {
  it("classifies each supported container", () => {
    expect(detectFormat(buildPdf({ pages: ["one"] }))).toBe("pdf");
    expect(detectFormat(buildDocx({ body: paragraph("hello") }))).toBe("docx");
    expect(
      detectFormat(
        buildPptx({
          slides: [
            {
              part: "ppt/slides/slide1.xml",
              shapes: sp({ text: pptParagraph({ runs: run({ text: "hi" }) }) }),
              rels: undefined,
            },
          ],
        }),
      ),
    ).toBe("pptx");
  });

  it("reports binary bytes and legacy binaries separately from spreadsheets", () => {
    const spreadsheet = zipSync({
      "[Content_Types].xml": strToU8("<Types/>"),
      "xl/workbook.xml": strToU8("<workbook/>"),
    });
    expect(detectFormat(spreadsheet)).toBe("xlsx");
    // Binary bytes are not text: the text detector must not claim them.
    expect(detectFormat(bytes(1, 2, 3, 4, 5, 6, 7, 8))).toBe("unknown");
    expect(detectFormat(bytes(9, 9, 9, 9))).toBe("unknown");
    expect(detectFormat(bytes(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1))).toBe("legacy");
  });

  it("recognises a text file from its bytes, since text has no signature", () => {
    // `.txt`, `.csv` and `.json` are all "bytes that decode as text"; the reading
    // is chosen by content, and the file name is never consulted.
    expect(detectFormat(strToU8("a plain note, with a comma"))).toBe("text");
    expect(detectFormat(strToU8("a,b,c\n1,2,3\n4,5,6\n"))).toBe("csv");
    expect(detectFormat(strToU8('{"a":1}'))).toBe("text");
  });

  it("does not confuse the two OPC formats", () => {
    expect(NS.r).toBeTruthy(); // fixture sanity: namespaces are available
    expect(detectFormat(buildDocx({ body: paragraph("x") }))).not.toBe("pptx");
  });
});
