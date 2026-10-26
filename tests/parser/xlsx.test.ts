import { describe, expect, it } from "vitest";
import { renderToStaticMarkup } from "react-dom/server";
import type { ReactElement } from "react";
import { detectFormat, parseDocumentBytes, parsePackageBytes } from "../../src/parser/document";
import { readArchive } from "../../src/parser/archive";
import { parseXlsxArchive } from "../../src/parser/xlsx";
import { formatSerialDate } from "../../src/parser/xlsx/styles";
import { mergeLabels } from "../../src/components/labels";
import { buildXlsx } from "../fixtures/xlsx";

const workbook = buildXlsx({
  title: "Quarterly numbers",
  creator: "Ada Lovelace",
  sheets: [
    {
      name: "Summary",
      rows: [
        ["Name", "Score", "When", "Active"],
        ["Ada", 42, { type: "date", serial: 45292 }, true],
        ["Grace", { type: "percent", value: 0.25 }, { type: "datetime", serial: 45292.5 }, false],
        ["Merged across", "", "", ""],
        [null, { type: "formula", formula: "SUM(B2:B3)", value: 45.5 }, { type: "error", value: "#DIV/0!" }, { type: "inline", value: "inline value" }],
      ],
      merges: ["A4:C4"],
      cols: [
        { min: 1, max: 2, width: 18 },
        { min: 3, max: 3, width: 24 },
      ],
    },
    {
      name: "Second Sheet",
      rows: [[null, null, "Sparse"], [], ["A1 only"]],
    },
  ],
});

function pageMarkup(parsed: ReturnType<typeof parseXlsxArchive>, pageIndex: number): string {
  const page = parsed.model.slides[pageIndex]!;
  const element = parsed.renderPage({
    page,
    width: 1056,
    height: 816,
    scale: 1,
    interactive: true,
    media: parsed.media,
    labels: mergeLabels(undefined),
    warn: () => undefined,
  });
  return renderToStaticMarkup(element as ReactElement);
}

describe("XLSX", () => {
  it("detects the package from xl/workbook.xml and parses worksheet pages", () => {
    expect(detectFormat(workbook)).toBe("xlsx");

    const parsed = parsePackageBytes(workbook);
    expect(parsed.format).toBe("xlsx");
    expect(parsed.renderPage).toBeTypeOf("function");
    expect(parsed.model.slides).toHaveLength(2);
    expect(parsed.model.slides.map((slide) => slide.name)).toEqual(["Summary", "Second Sheet"]);
    expect(parsed.model.title).toBe("Quarterly numbers");
    expect(parsed.model.author).toBe("Ada Lovelace");
    expect(parsed.model.width).toBe(Math.round((1056 / 96) * 914_400));
    expect(parsed.model.height).toBe(Math.round((816 / 96) * 914_400));
    parsed.dispose();
  });

  it("renders shared strings, numbers, dates, percentages, booleans and cached formulas", () => {
    const parsed = parseXlsxArchive(readArchive(workbook));
    const markup = pageMarkup(parsed, 0);
    expect(markup).toContain("Name");
    expect(markup).toContain("Ada");
    expect(markup).toContain("42");
    expect(markup).toContain("2024-01-01");
    expect(markup).toContain("25.00%");
    expect(markup).toContain("TRUE");
    expect(markup).toContain("45.5");
    expect(markup).toContain("#DIV/0!");
    expect(markup).toContain("inline value");
    // React writes the DOM property name into static markup; browsers apply it as `colspan`.
    expect(markup).toMatch(/colspan="3"/i);
    expect(markup).toContain("Merged across");
    parsed.dispose();
  });

  it("keeps sparse cells at their real coordinates instead of shifting rows", () => {
    const parsed = parseXlsxArchive(readArchive(workbook));
    const markup = pageMarkup(parsed, 1);
    expect(markup).toContain("Sparse");
    expect(markup).toContain("A1 only");
    // The second sheet has a row with no cells; it must still occupy row 2.
    expect(markup).toContain('scope="row">2</th>');
    parsed.dispose();
  });

  it("decodes worksheet references through workbook relationships, not file order", () => {
    const parsed = parseXlsxArchive(readArchive(workbook));
    // A long sheet becomes several pages, so each page appends its own suffix.
    expect(parsed.model.slides[0]?.partName.startsWith("xl/worksheets/sheet1.xml")).toBe(true);
    expect(parsed.model.slides[1]?.partName.startsWith("xl/worksheets/sheet2.xml")).toBe(true);
    parsed.dispose();
  });

  it("supports the 1904 date system", () => {
    const parsed = parseXlsxArchive(
      readArchive(
        buildXlsx({
          date1904: true,
          sheets: [{ name: "Dates", rows: [[{ type: "date", serial: 0 }, { type: "date", serial: 1 }]] }],
        }),
      ),
    );
    const markup = pageMarkup(parsed, 0);
    expect(markup).toContain("1904-01-01");
    expect(markup).toContain("1904-01-02");
    parsed.dispose();
  });

  it("reports a formula without a cached value instead of inventing one", () => {
    const parsed = parseXlsxArchive(
      readArchive(
        buildXlsx({
          sheets: [{ name: "No cache", rows: [[{ type: "formula", formula: "1+1", value: null }]] }],
        }),
      ),
    );
    expect(parsed.model.diagnostics.map((diagnostic) => diagnostic.code)).toContain("xlsx.formula.cache");
    expect(pageMarkup(parsed, 0)).not.toContain("2");
    parsed.dispose();
  });

  it("is reachable through the async document entry point and the format detector", async () => {
    expect(detectFormat(workbook)).toBe("xlsx");
    const parsed = await parseDocumentBytes(workbook);
    expect(parsed.format).toBe("xlsx");
    expect(parsed.renderPage).toBeTypeOf("function");
    parsed.dispose();
  });

  it("formats serial dates across the 1900 leap-year boundary", () => {
    expect(formatSerialDate(1, "date", false)).toBe("1900-01-01");
    expect(formatSerialDate(59, "date", false)).toBe("1900-02-28");
    expect(formatSerialDate(60, "date", false)).toBe("1900-02-29");
    expect(formatSerialDate(61, "date", false)).toBe("1900-03-01");
    expect(formatSerialDate(45292, "date", false)).toBe("2024-01-01");
  });
});
