import type { SheetPageData } from "../parser/xlsx/model";

export function selectionText(data: SheetPageData, start: [number, number], end: [number, number]): string {
  const rows: string[] = [];
  for (let row = Math.min(start[0], end[0]); row <= Math.max(start[0], end[0]); row++) {
    const values: string[] = [];
    for (let col = Math.min(start[1], end[1]); col <= Math.max(start[1], end[1]); col++) {
      const cell = data.rows[row]?.[col];
      const text = !cell || cell.covered ? "" : cell.text;
      values.push(/[\t\r\n"]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text);
    }
    rows.push(values.join("\t"));
  }
  return rows.join("\n");
}

export function sheetCsv(data: SheetPageData): string {
  return "\uFEFF" + data.rows.map((row) => row.map((cell) => {
    let text = cell.covered ? "" : cell.text;
    // Keep text cells inert when a downloaded CSV is opened in a spreadsheet app.
    if (!cell.numeric && /^[\s]*[=+@-]/.test(text)) text = "'" + text;
    return `"${text.replace(/"/g, '""')}"`;
  }).join(",")).join("\r\n");
}

export interface SheetMatch { sheet: number; row: number; col: number; text: string }
export function findSheetMatches(pages: readonly SheetPageData[], query: string): SheetMatch[] {
  if (!query.trim()) return [];
  const needle = query.toLocaleLowerCase();
  const results: SheetMatch[] = [];
  for (const [sheet, page] of pages.entries()) for (const [row, cells] of page.rows.entries()) for (const [col, cell] of cells.entries()) {
    if (!cell.covered && cell.text.toLocaleLowerCase().includes(needle)) results.push({ sheet, row, col, text: cell.text });
    if (results.length >= 500) return results;
  }
  return results;
}
