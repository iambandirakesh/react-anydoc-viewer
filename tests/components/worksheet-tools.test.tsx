import { afterEach, expect, it } from "vitest";
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { buildXlsx } from "../fixtures/xlsx";
import { DocumentViewer } from "../../src/components/DocumentViewer";
import { parseXlsxArchive } from "../../src/parser/xlsx";
import { readArchive } from "../../src/parser/archive";
import { findSheetMatches, selectionText, sheetCsv } from "../../src/components/worksheetUtils";

afterEach(cleanup);
const workbook = () => buildXlsx({ sheets: [{ name: "Dashboard", rows: [["Heading", "Revenue"], ["Cosmetics", 10], ["=UNSAFE()", 20]] }, { name: "Categories", rows: [["Category", "Revenue"], ["Cosmetics", 30]] }] });
it("finds cells across sheets and exports safely quoted values", () => {
  const parsed = parseXlsxArchive(readArchive(workbook()));
  expect(findSheetMatches(parsed.pageData, "cosmetics").map((match) => match.sheet)).toEqual([0, 1]);
  expect(selectionText(parsed.pageData[0]!, [2, 1], [1, 0])).toBe("Cosmetics\t10\n=UNSAFE()\t20");
  expect(sheetCsv(parsed.pageData[0]!)).toContain('"\'=UNSAFE()"');
  parsed.dispose();
});
it("navigates search results, selects ranges, and restores sheet scroll", async () => {
  const { container } = render(<DocumentViewer source={workbook()} />);
  await waitFor(() => expect(screen.queryAllByRole("tab")).toHaveLength(2));
  const scroll = container.querySelector(".radv-sheet-scroll")!;
  scroll.scrollTop = 80; scroll.scrollLeft = 20; fireEvent.scroll(scroll);
  fireEvent.click(screen.getByRole("tab", { name: "Categories" }));
  fireEvent.click(screen.getByRole("tab", { name: "Dashboard" }));
  expect(container.querySelector(".radv-sheet-scroll")!.scrollTop).toBe(80);
  const first = container.querySelector('[data-cell="1:0"]')!;
  fireEvent.pointerDown(first);
  fireEvent.keyDown(first, { key: "ArrowRight", shiftKey: true });
  expect(container.querySelectorAll('[data-cell][aria-selected="true"]')).toHaveLength(2);
  let copied = "";
  fireEvent.copy(first, { clipboardData: { setData: (_: string, value: string) => { copied = value; } } });
  expect(copied).toBe("Cosmetics\t10");
  fireEvent.change(screen.getByRole("searchbox", { name: "Search all worksheets" }), { target: { value: "cosmetics" } });
  fireEvent.click(screen.getByRole("button", { name: "Next search result" }));
  expect(screen.getByRole("tab", { name: "Categories" }).getAttribute("aria-selected")).toBe("true");
});

it("shows search for CSV but not for presentation documents", async () => {
  const { buildPptx } = await import("../fixtures/pptx");
  const { container, unmount } = render(<DocumentViewer source={buildPptx({ slides: [{ part: "ppt/slides/slide1.xml", shapes: "" }] })} />);
  await waitFor(() => expect(container.querySelector('[data-status="ready"]')).not.toBeNull());
  expect(screen.queryByRole("searchbox")).toBeNull();
  unmount();
  render(<DocumentViewer source={new TextEncoder().encode("Category,Revenue\nCosmetics,100\nClothes,200")} />);
  const search = await screen.findByRole("searchbox", { name: "Search all worksheets" });
  fireEvent.change(search, { target: { value: "Cosmetics" } });
  await waitFor(() => expect(screen.getByRole("button", { name: "Next search result" })).not.toBeDisabled());
});

it("keeps worksheet panel IDs unique across viewer instances", async () => {
  const source = workbook();
  const { container } = render(<><DocumentViewer source={source} /><DocumentViewer source={source} /></>);
  await waitFor(() => expect(container.querySelectorAll('[role="tabpanel"]')).toHaveLength(2));
  const panels = [...container.querySelectorAll<HTMLElement>('[role="tabpanel"]')];
  expect(panels[0]!.id).toBeTruthy();
  expect(panels[1]!.id).not.toBe(panels[0]!.id);
  for (const panel of panels) {
    const tabs = panel.closest(".radv-workbook")!.querySelectorAll('[role="tab"]');
    for (const tab of tabs) expect(tab.getAttribute("aria-controls")).toBe(panel.id);
  }
});
