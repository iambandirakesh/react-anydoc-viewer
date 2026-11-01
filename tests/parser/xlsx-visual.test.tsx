import { describe, expect, it } from "vitest";
import { renderToStaticMarkup } from "react-dom/server";
import { buildDashboardWorkbook } from "../fixtures/xlsxDashboard";
import { buildXlsx } from "../fixtures/xlsx";
import { readArchive } from "../../src/parser/archive";
import { parseXlsxArchive } from "../../src/parser/xlsx";
import { SheetPage } from "../../src/components/SheetPage";
import { parseXml } from "../../src/parser/xml";
import { parseStyles, formatNumericCell } from "../../src/parser/xlsx/styles";
import { writeFileSync } from "node:fs";

describe("XLSX visual worksheets", () => {
  it("preserves theme styles, blank styled cells, merged cells and dimensions", () => {
    const parsed = parseXlsxArchive(readArchive(buildDashboardWorkbook()));
    const page = parsed.pageData[0]!;
    expect(page.rows[0]![0]!.style).toMatchObject({ backgroundColor: "rgb(43, 42, 92)", color: "rgb(255, 255, 255)", fontWeight: 700, fontSize: `${22 * 96 / 72}px` });
    expect(page.rows[0]![0]!.colSpan).toBe(14);
    expect(page.rows[0]![13]!.style?.backgroundColor).toBe("rgb(43, 42, 92)");
    expect(page.heights?.slice(0, 4)).toEqual([48, 40, 40, 0]);
    expect(page.widths[0]).toBe(75);
    expect(page.widths[14]).toBe(0);
    expect(page.showGridLines).toBe(false);
    expect(page.rows[2]![3]!.text).toBe("$137,348,768");
    expect(page.rows[2]![7]!.text).toBe("137.3M");
    parsed.dispose();
  });

  it("resolves charts and images at two-cell, one-cell and absolute anchors", () => {
    const parsed = parseXlsxArchive(readArchive(buildDashboardWorkbook()));
    const page = parsed.pageData[0]!;
    expect(page.drawings?.map((node) => node.kind)).toEqual(["chart", "chart", "picture"]);
    expect(page.drawings?.[0]?.xfrm).toMatchObject({ x: 0, y: 152 * 9525, cx: 525 * 9525, cy: 432 * 9525 });
    expect(page.drawings?.[1]?.xfrm).toMatchObject({ x: 525 * 9525, y: 152 * 9525, cx: 4762500 });
    const markup = renderToStaticMarkup(<SheetPage data={page} width={1056} height={816} interactive media={parsed.media} />);
    expect(markup).toContain("Revenue by category");
    expect(markup).toContain("70.7%");
    expect(markup).toContain("Dashboard logo");
    expect(markup).toContain('data-gridlines="false"');
    expect(markup).toContain("$137,348,768");
    expect(parsed.model.diagnostics).toHaveLength(0);
    parsed.dispose();
  });

  it("keeps external worksheet drawings offline and reports missing assets", () => {
    const bytes = buildXlsx({ sheets: [{ name: "External", rows: [["Safe"]] }], extraParts: {
      "xl/worksheets/sheet1.xml": '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="remote"/></worksheet>',
      "xl/worksheets/_rels/sheet1.xml.rels": '<Relationships><Relationship Id="remote" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="https://example.com/private.xml" TargetMode="External"/></Relationships>',
    } });
    const parsed = parseXlsxArchive(readArchive(bytes));
    expect(parsed.pageData[0]?.drawings).toEqual([]);
    expect(parsed.model.diagnostics.some((entry) => entry.code === "xlsx.drawing.relationship")).toBe(true);
    parsed.dispose();
  });

  it.each([
    ['#,##0', 1780539, '1,780,539'],
    ['"$"#,##0.00', 1780539.25, '$1,780,539.25'],
    ['0.0,,"M"', 137348768, '137.3M'],
    ['0.0%', .707, '70.7%'],
    ['#,##0;(#,##0)', -1250, '(1,250)'],
  ])("applies explicit numeric display format %s", (code, value, expected) => {
    const root = parseXml(`<styleSheet><numFmts><numFmt numFmtId="164" formatCode="${String(code).replaceAll('"', '&quot;')}"/></numFmts><cellXfs><xf numFmtId="164"/></cellXfs></styleSheet>`);
    expect(formatNumericCell(Number(value), 0, parseStyles(root), false)).toBe(expected);
  });

  it("can generate the dashboard for browser QA", () => {
    if (process.env.XLSX_DASHBOARD_OUT) writeFileSync(process.env.XLSX_DASHBOARD_OUT, buildDashboardWorkbook());
    expect(buildDashboardWorkbook().byteLength).toBeGreaterThan(0);
  });

  it("keeps large styled grids within the existing render cell budget", () => {
    const bytes = buildXlsx({ sheets: [{ name: "Large sheet", rows: Array.from({ length: 501 }, (_, row) => Array.from({ length: 20 }, (_, col) => row * 20 + col)) }] });
    const parsed = parseXlsxArchive(readArchive(bytes));
    expect(parsed.pageData[0]?.rows).toHaveLength(500);
    expect(parsed.pageData[0]?.rows[0]).toHaveLength(20);
    expect(parsed.pageData[0]?.truncated).toBe(true);
    if (process.env.XLSX_STRESS_OUT) writeFileSync(process.env.XLSX_STRESS_OUT, bytes);
    parsed.dispose();
  });
});

describe("worksheet conditional formatting", () => {
  it("uses cached numeric values for data bars and three-color scales", async () => {
    const { applyConditionalFormatting } = await import("../../src/parser/xlsx/conditionalFormatting");
    const parsed = parseXlsxArchive(readArchive(buildDashboardWorkbook()));
    const page = parsed.pageData[0]!;
    const root = parseXml(`<worksheet><sheetData><row><c r="A1"><v>0</v></c><c r="B1"><v>50</v></c><c r="C1"><v>100</v></c></row></sheetData>
      <conditionalFormatting sqref="A1:C1"><cfRule type="dataBar"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF0000FF"/></dataBar></cfRule>
      <cfRule type="colorScale"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FFFFFFFF"/><color rgb="FF00FF00"/></colorScale></cfRule></conditionalFormatting></worksheet>`);
    const warnings: string[] = [];
    applyConditionalFormatting(root!, page, parsed.model.theme, (code) => warnings.push(code));
    expect(page.rows[0]![0]!.dataBar?.width).toBe(10);
    expect(page.rows[0]![1]!.dataBar?.width).toBe(50);
    expect(page.rows[0]![2]!.dataBar?.width).toBe(90);
    expect(page.rows[0]![0]!.style?.backgroundColor).toBe("rgb(255, 0, 0)");
    expect(page.rows[0]![1]!.style?.backgroundColor).toBe("rgb(255, 255, 255)");
    expect(page.rows[0]![2]!.style?.backgroundColor).toBe("rgb(0, 255, 0)");
    expect(warnings).toEqual([]);
    parsed.dispose();
  });
});

it("renders charts that reference worksheet cells without chart caches", async () => {
  const { unzipSync, zipSync, strToU8 } = await import("fflate");
  const parts = unzipSync(buildDashboardWorkbook());
  parts["xl/charts/chart1.xml"] = strToU8(`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:barDir val="bar"/><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:numRef><c:f>'Dashboard'!$A$26:$A$27</c:f></c:numRef></c:cat><c:val><c:numRef><c:f>'Dashboard'!$D$26:$D$27</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`);
  const parsed = parseXlsxArchive(readArchive(zipSync(parts)));
  const first = parsed.pageData[0]!.drawings?.[0];
  expect(first?.kind).toBe("chart");
  if (first?.kind !== "chart") throw new Error("Missing reference-only chart");
  expect(first.chart.series[0]!.values).toEqual([36601510, 30585380]);
  expect(first.chart.categories).toEqual(["Cosmetics", "Office Supplies"]);
  expect(renderToStaticMarkup(<SheetPage data={parsed.pageData[0]!} width={1056} height={816} interactive media={parsed.media} />)).toContain("36,601,510");
  expect(parsed.pageData[0]!.drawings).toHaveLength(3);
  parsed.dispose();
});

it("spaces worksheet charts below headers and thins crowded category ticks without dropping data", async () => {
  const { renderChart } = await import("../../src/renderer/RenderChart");
  const parsed = parseXlsxArchive(readArchive(buildDashboardWorkbook()));
  const node = parsed.pageData[0]!.drawings![0]!;
  if (node.kind !== "chart") throw new Error("Chart fixture missing");
  const markup = renderToStaticMarkup(<SheetPage data={parsed.pageData[0]!} width={1056} height={816} interactive media={parsed.media} />);
  expect(markup).toContain("top:164px");
  const categories = Array.from({ length: 16 }, (_, i) => `Week ${i}`);
  const chart = { ...node, xfrm: { ...node.xfrm, cx: 300 * 9525 }, chart: { ...node.chart, type: "lineChart" as const, barDir: "col" as const, categories, series: [{ ...node.chart.series[0]!, categories, values: categories.map((_, i) => i + 1) }] } };
  const container = document.createElement("div");
  container.innerHTML = renderToStaticMarkup(renderChart(chart));
  const labels = [...container.querySelectorAll("text")].map((text) => text.textContent).filter((text) => text?.startsWith("Week "));
  expect(labels.length).toBeGreaterThan(1);
  expect(labels.length).toBeLessThan(16);
  expect(container.querySelectorAll("circle")).toHaveLength(16);
  parsed.dispose();
});

it("lets unwrapped merged titles extend beyond the used range", () => {
  const parsed = parseXlsxArchive(readArchive(buildXlsx({ sheets: [{ name: "Monthly", rows: [["Revenue by calendar month"], ["Month", "Revenue"], ["Jan", 100]], merges: ["A1:B1"] }] })));
  const markup = renderToStaticMarkup(<SheetPage data={parsed.pageData[0]!} width={900} height={600} interactive />);
  const container = document.createElement("div");
  container.innerHTML = markup;
  const title = container.querySelector("tbody td") as HTMLTableCellElement;
  expect(title.colSpan).toBe(2);
  expect(title.style.overflow).toBe("visible");
  expect((title.firstElementChild as HTMLElement).style.width).toBe("max-content");
  const header = container.querySelectorAll("tbody tr")[1]!.querySelector("td")!;
  expect((header.firstElementChild as HTMLElement).style.width).not.toBe("max-content");
  parsed.dispose();
});

it("supports negative data bars, icon thresholds, and comparison priorities", async () => {
  const { applyConditionalFormatting } = await import("../../src/parser/xlsx/conditionalFormatting");
  const parsed = parseXlsxArchive(readArchive(buildDashboardWorkbook()));
  const page = parsed.pageData[1]!;
  const root = parseXml(`<worksheet><sheetData><row><c r="A1"><v>-100</v></c><c r="B1"><v>0</v></c><c r="C1"><v>100</v></c></row></sheetData>
    <conditionalFormatting sqref="A1:C1"><cfRule type="dataBar" priority="4"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF0000FF"/></dataBar></cfRule>
    <cfRule type="iconSet" priority="3"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>
    <cfRule type="cellIs" priority="2" dxfId="1" operator="greaterThan"><formula>-200</formula></cfRule>
    <cfRule type="cellIs" priority="1" dxfId="0" operator="greaterThan"><formula>50</formula></cfRule></conditionalFormatting></worksheet>`)!;
  const styles = parseXml(`<styleSheet><dxfs><dxf><font><color rgb="FF00FF00"/></font></dxf><dxf><font><color rgb="FFFF0000"/></font></dxf></dxfs></styleSheet>`);
  const warnings: string[] = [];
  applyConditionalFormatting(root, page, parsed.model.theme, (code) => warnings.push(code), styles);
  expect(page.rows[0]![0]!.dataBar).toMatchObject({ left: 5, width: 45, color: "#d45a5a" });
  expect(page.rows[0]![2]!.dataBar).toMatchObject({ left: 50, width: 45 });
  expect(page.rows[0]![0]!.icon?.text).toBe("↓");
  expect(page.rows[0]![2]!.icon?.text).toBe("↑");
  expect(page.rows[0]![0]!.style?.color).toBe("rgb(255, 0, 0)");
  expect(page.rows[0]![2]!.style?.color).toBe("rgb(0, 255, 0)");
  expect(warnings).toEqual([]);
  parsed.dispose();

});
