/** A cell as it will appear on a spreadsheet page. */
export interface SheetCell {
  /** text exactly as it should be displayed; already number/date formatted */
  text: string;
  /** horizontal alignment, derived from the cell style when present */
  align: "left" | "right" | "center";
  /** true for a cell the reader would expect to be right-aligned (numbers/dates) */
  numeric: boolean;
  /** true when covered by the anchor of a merged range; it is not emitted */
  covered: boolean;
  /** number of columns this cell spans (1 when not merged) */
  colSpan: number;
  /** number of rows this cell spans (1 when not merged) */
  rowSpan: number;
}

/** One rendered worksheet. A sheet is one page, just like a slide. */
export interface SheetPageData {
  name: string;
  rows: SheetCell[][];
  /** widths in CSS pixels for each data column, without the row-number column */
  widths: number[];
  /** true when rows/columns were omitted because the sheet is very large */
  truncated: boolean;
}
