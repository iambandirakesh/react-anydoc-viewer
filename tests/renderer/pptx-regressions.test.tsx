import { afterEach, expect, it } from "vitest";
import { cleanup, render } from "@testing-library/react";
import { buildPptx, sp, group, inch, paragraph, run, chartXml, chartFrame, NS, RED_PNG, tableXml, tableCell, tableFrame } from "../fixtures/pptx";
import { parsePresentationBytes } from "../../src/parser/pptx";
import { SlideView } from "../../src/renderer/SlideView";

afterEach(cleanup);

function view(shapes: string, extra: Partial<Parameters<typeof buildPptx>[0]> = {}) {
  const parsed = parsePresentationBytes(buildPptx({
    slides: [{ part: "ppt/slides/slide1.xml", shapes, rels: {
      rId2: { target: "../charts/chart1.xml", type: `${NS.r}/chart` },
      rId3: { target: "../media/image1.png", type: `${NS.r}/image` },
    } }],
    ...extra,
  }));
  const result = render(<SlideView slide={parsed.model.slides[0]!} media={parsed.media} width={1280} height={720} scope="regression" interactive />);
  return { parsed, ...result };
}

it.each(["col", "bar"] as const)("shows exact %s data labels and category/series hover details", (barDir) => {
  const chart = chartXml({ barDir, showValueLabels: true, series: [{ name: "Week ending May 21", categories: ["Cereal", "Cosmetics"], values: [1780539, 793518.25] }] });
  const { container } = view(chartFrame({ rId: "rId2" }), { charts: { "ppt/charts/chart1.xml": chart } });
  const labels = [...container.querySelectorAll("[data-chart-series] text")].map((node) => node.firstChild?.textContent);
  expect(labels).toEqual(["1,780,539", "793,518.25"]);
  const details = [...container.querySelectorAll("[data-chart-series] rect title")].map((node) => node.textContent);
  expect(details).toEqual(["Cereal · Week ending May 21 · 1,780,539", "Cosmetics · Week ending May 21 · 793,518.25"]);
});

it("keeps pie percentage labels and provides exact values on hover", () => {
  const chart = chartXml({ type: "pieChart", showPercent: true, series: [{ name: "Revenue", categories: ["Cereal", "Cosmetics"], values: [1780539, 793518] }] });
  const { container } = view(chartFrame({ rId: "rId2" }), { charts: { "ppt/charts/chart1.xml": chart } });
  expect(container.querySelector("svg[role='img'] path title")?.textContent).toBe("Cereal · Revenue · 1,780,539 · 69.2%");
  expect([...container.querySelectorAll("svg[role='img'] text")].some((node) => node.textContent === "69.2%")).toBe(true);
});

it.each(["col", "bar"] as const)("centers sparse %s clusters on their category and keeps populated clusters paired", (barDir) => {
  const chart = chartXml({ barDir, showValueLabels: true, series: [
    { name: "Current", categories: ["Cosmetics", "Cereal", "Both"], values: [793518, 0, 10] },
    { name: "Prior", categories: ["Cosmetics", "Cereal", "Both"], values: [0, 1780539, 20] },
  ] });
  const { container } = view(chartFrame({ rId: "rId2" }), { charts: { "ppt/charts/chart1.xml": chart } });
  const category = (name: string) => [...container.querySelectorAll("svg[role='img'] text")].find((node) => node.textContent === name)!;
  const series = container.querySelectorAll("[data-chart-series]");
  const current = series[0]!.querySelectorAll("rect");
  const prior = series[1]!.querySelectorAll("rect");
  const center = (rect: Element) => Number(rect.getAttribute(barDir === "col" ? "x" : "y")) + Number(rect.getAttribute(barDir === "col" ? "width" : "height")) / 2;
  // Horizontal category text has a three-pixel baseline offset.
  const axisCenter = (name: string) => Number(category(name).getAttribute(barDir === "col" ? "x" : "y")) - (barDir === "bar" ? 3 : 0);
  expect(center(current[0]!)).toBeCloseTo(axisCenter("Cosmetics"), 0);
  expect(center(prior[0]!)).toBeCloseTo(axisCenter("Cereal"), 0);
  expect((center(current[1]!) + center(prior[1]!)) / 2).toBeCloseTo(axisCenter("Both"), 0);
  expect(center(current[1]!)).toBeLessThan(center(prior[1]!));
  const label = series[0]!.querySelector("text")!;
  const labelCenter = Number(label.getAttribute(barDir === "col" ? "x" : "y")) - (barDir === "bar" ? 3 : 0);
  expect(labelCenter).toBeCloseTo(barDir === "col" ? axisCenter("Cosmetics") : center(current[0]!), 0);
});

it("applies recorded autofit to runs, bullets and paragraph line spacing", () => {
  const shape = sp({ text: paragraph({
    bullet: '<a:buSzPct val="100000"/><a:buChar char="•"/>',
    runs: run({ text: "Scaled", size: 2000 }),
  }) }).replace("<a:bodyPr/>", '<a:bodyPr><a:normAutofit fontScale="50000" lnSpcReduction="10000"/></a:bodyPr>');
  const { container } = view(shape);
  const p = container.querySelector("p")!;
  expect(parseFloat((container.querySelector("[data-run]") as HTMLElement).style.fontSize)).toBeCloseTo(40 / 3);
  expect(parseFloat((p.firstElementChild as HTMLElement).style.fontSize)).toBeCloseTo(40 / 3);
  expect(parseFloat(p.style.lineHeight)).toBeCloseTo(1.08);
});

it("paints a cropped image fill outside defs and retains its shape clip", () => {
  const { container } = view(sp({ fill: '<a:blipFill><a:blip r:embed="rId3"/><a:srcRect l="25000"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>' }), {
    media: [{ part: "ppt/media/image1.png", bytes: RED_PNG }],
  });
  expect(container.querySelector("defs image")).toBeNull();
  const image = container.querySelector("svg > image")!;
  expect(image).not.toBeNull();
  expect(Number(image.getAttribute("x"))).toBeLessThan(0);
  const clipId = image.getAttribute("clip-path")!.slice(5, -1);
  expect(container.querySelector(`[id="${clipId}"] path`)).not.toBeNull();
});

it("scales the group child offset with the child coordinates", () => {
  const { container } = view(group({ cx: inch(4), chCx: inch(2), chX: inch(1), children: sp({ x: inch(1), fill: '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' }) }));
  const transforms = [...container.querySelectorAll<HTMLElement>("[style]")].map((node) => node.style.transform);
  // x=96 maps to 2*(96-96)=0, rather than 2*96-96=96.
  expect(transforms).toContain("scale(2, 1) translate(-96px, 0px)");
});

it.each(["col", "bar"] as const)("stacks %s series end to end, with separate negative totals", (barDir) => {
  const chart = chartXml({ barDir, series: [
    { name: "A", categories: ["Positive", "Negative"], values: [10, -10] },
    { name: "B", categories: ["Positive", "Negative"], values: [20, -20] },
  ] }).replace('val="clustered"', 'val="stacked"');
  const { container } = view(chartFrame({ rId: "rId2" }), { charts: { "ppt/charts/chart1.xml": chart } });
  const bars = [...container.querySelectorAll("[data-chart-series] rect")];
  expect(bars).toHaveLength(4);
  const coordinate = barDir === "col" ? "x" : "y";
  expect(bars[0]!.getAttribute(coordinate)).toBe(bars[2]!.getAttribute(coordinate));
  const length = barDir === "col" ? "height" : "width";
  expect(Number(bars[2]!.getAttribute(length))).toBeCloseTo(Number(bars[0]!.getAttribute(length)) * 2);
  if (barDir === "col") {
    expect(Number(bars[2]!.getAttribute("y")) + Number(bars[2]!.getAttribute("height"))).toBeCloseTo(Number(bars[0]!.getAttribute("y")));
  } else {
    expect(Number(bars[0]!.getAttribute("x")) + Number(bars[0]!.getAttribute("width"))).toBeCloseTo(Number(bars[2]!.getAttribute("x")));
  }
});

it("normalizes each percent-stacked category to the same total height", () => {
  const chart = chartXml({ series: [
    { name: "A", categories: ["X", "Y"], values: [10, 100] },
    { name: "B", categories: ["X", "Y"], values: [30, 300] },
  ] }).replace('val="clustered"', 'val="percentStacked"');
  const { container } = view(chartFrame({ rId: "rId2" }), { charts: { "ppt/charts/chart1.xml": chart } });
  const bars = [...container.querySelectorAll("[data-chart-series] rect")];
  expect(bars[0]!.getAttribute("height")).toBe(bars[1]!.getAttribute("height"));
  expect(Number(bars[2]!.getAttribute("height"))).toBeCloseTo(Number(bars[0]!.getAttribute("height")) * 3);
});

it("assigns distinct pie colors while preserving an explicit point fill", () => {
  const chart = chartXml({ type: "pieChart", series: [{ name: "Share", categories: ["A", "B", "C"], values: [50, 30, 20], pointColors: ["FF0000", null, null] }] })
    .replace('<c:varyColors val="0"/>', '<c:varyColors val="1"/>');
  const { container } = view(chartFrame({ rId: "rId2" }), { charts: { "ppt/charts/chart1.xml": chart } });
  const fills = [...container.querySelectorAll("svg[role='img'] path")].map((path) => path.getAttribute("fill"));
  expect(fills[0]).toBe("rgb(255, 0, 0)");
  expect(new Set(fills).size).toBe(3);
});

it("retains cells following explicit horizontal and vertical merge continuations", () => {
  const anchor = tableCell({ text: "Merged", gridSpan: 2 }).replace('<a:tc gridSpan="2">', '<a:tc gridSpan="2" rowSpan="2">');
  const { container } = view(tableFrame({ table: tableXml({ colWidths: [inch(1), inch(1), inch(1)], rows: [
    { height: inch(1), cells: [anchor, tableCell({ hMerge: true }), tableCell({ text: "Top right" })] },
    { height: inch(1), cells: [tableCell({ vMerge: true }), tableCell({ hMerge: true, vMerge: true }), tableCell({ text: "Bottom right" })] },
  ] }) }));
  expect(container.querySelector('[data-cell="0-2"]')?.textContent).toContain("Top right");
  expect(container.querySelector('[data-cell="1-2"]')?.textContent).toContain("Bottom right");
  expect(container.querySelectorAll("[data-cell]")).toHaveLength(3);
});
