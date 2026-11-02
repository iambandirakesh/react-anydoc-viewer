import { useEffect, useMemo, useState, type RefObject } from "react";
import type { SheetPageData } from "../parser/xlsx/model";
import { findSheetMatches, sheetCsv, selectionText, type SheetMatch } from "./worksheetUtils";
import { columnLabel } from "./SheetPage";

export function WorksheetTools({ pages, index, navigate, fitWidth, root }: { pages: readonly SheetPageData[]; index: number; navigate: (index: number) => void; fitWidth: () => void; root: RefObject<HTMLDivElement | null> }) {
  const [query, setQuery] = useState("");
  const [resultIndex, setResultIndex] = useState(0);
  const [pending, setPending] = useState<SheetMatch | null>(null);
  const [message, setMessage] = useState<{ text: string } | null>(null);
  useEffect(() => {
    if (!message) return;
    const timer = setTimeout(() => setMessage(null), 2000);
    return () => clearTimeout(timer);
  }, [message]);
  const results = useMemo(() => findSheetMatches(pages, query), [pages, query]);
  useEffect(() => {
    if (!pending || pending.sheet !== index) return;
    const element = root.current?.querySelector<HTMLElement>(`[data-cell="${pending.row}:${pending.col}"]`);
    element?.scrollIntoView?.({ block: "nearest", inline: "nearest" });
    element?.focus();
    setPending(null);
  }, [pending, index, root]);
  const jump = (next: number) => {
    if (!results.length) return;
    const position = (next + results.length) % results.length;
    setResultIndex(position); setPending(results[position]!); navigate(results[position]!.sheet);
  };
  const current = pages[index];
  return <div className="radv-worksheet-tools" role="toolbar" aria-label="Worksheet tools">
    <input type="search" aria-label="Search all worksheets" placeholder="Search worksheets" value={query} onChange={(event) => { setQuery(event.target.value); setResultIndex(0); }} onKeyDown={(event) => { if (event.key === "Enter") { event.preventDefault(); jump(event.shiftKey ? resultIndex - 1 : resultIndex); } }} />
    <button type="button" disabled={!results.length} onClick={() => jump(resultIndex - 1)} aria-label="Previous search result">↑</button>
    <button type="button" disabled={!results.length} onClick={() => jump(resultIndex + 1)} aria-label="Next search result">↓</button>
    <span role="status" aria-live="polite">{query ? results.length ? `${Math.min(resultIndex + 1, results.length)} / ${results.length}${results.length === 500 ? "+" : ""} · ${pages[results[resultIndex]?.sheet ?? 0]?.name} ${columnLabel(results[resultIndex]?.col ?? 0)}${(results[resultIndex]?.row ?? 0) + 1}` : "No matches" : ""}</span>
    <button type="button" onClick={fitWidth}>Fit width</button>
    <button type="button" onClick={async () => {
      const selected = root.current?.querySelectorAll<HTMLElement>('[data-cell][aria-selected="true"]');
      if (!current || !selected?.length) { setMessage({ text: "Select a cell or range first" }); return; }
      const coords = [...selected].map((element) => element.dataset.cell!.split(":").map(Number));
      try { await navigator.clipboard.writeText(selectionText(current, [Math.min(...coords.map(c => c[0]!)), Math.min(...coords.map(c => c[1]!))], [Math.max(...coords.map(c => c[0]!)), Math.max(...coords.map(c => c[1]!))])); setMessage({ text: "Selection copied" }); }
      catch { setMessage({ text: "Press Control or Command C inside the selected cells to copy" }); }
    }}>Copy selection</button>
    <button type="button" disabled={!current} onClick={() => {
      if (!current) return;
      const url = URL.createObjectURL(new Blob([sheetCsv(current)], { type: "text/csv;charset=utf-8" }));
      const link = document.createElement("a"); link.href = url; link.download = `${current.name.replace(/[\\/:*?"<>|]/g, "_")}.csv`; link.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    }}>Download CSV</button>
    <span role="status">{message?.text}</span>
  </div>;
}
