/**
 * WordprocessingML table parsing (`w:tbl`).
 *
 * Word's table styles (`w:tblStyle`) are a template feature we do not
 * implement: when a table declares no borders of its own, a neutral grid is
 * applied so the table stays legible, and the approximation is reported.
 *
 * Row heights are the one thing that cannot be read from the file — Word
 * normally stores none and computes them from the content — so they are
 * measured with the same text layout the renderer uses. Those heights are then
 * handed to the table renderer as declared heights, which is what keeps the
 * absolutely positioned cell grid aligned with the measurement.
 */

import type { Fill, RGBA, TableBorders, TableCell, TableData, TableRow, TextBody } from "../../model/types";
import { child, children, type XmlNode } from "../xml";
import { EMU_PER_PX, twipsToEmu, twipsToPx } from "../units";
import { wAttr, wNum, wVal } from "./xmlw";
import { parseHexColor } from "./styles";
import { parseParagraphBlocks, type DocxTextContext } from "./text";
import type { DocxBlock, DocxParagraphBlock, DocxTableBlock, ParagraphMeasurer } from "./model";

/** Word's default cell margins: 0.075in left and right, nothing top and bottom. */
const DEFAULT_CELL_MARGIN_TWIPS = { top: 0, left: 108, bottom: 0, right: 108 };
/** Height used for a row whose content could not be measured at all. */
const FALLBACK_ROW_HEIGHT_PX = 24;

export interface TableParseContext extends DocxTextContext {
  measure: ParagraphMeasurer;
  /** width available to the table, in EMU (the page's content width) */
  contentWidthEmu: number;
}

function parseBorderEdge(node: XmlNode | null): TableBorders[keyof TableBorders] {
  if (!node) return null;
  const value = (wVal(node) ?? "single").toLowerCase();
  if (value === "none" || value === "nil") return { width: 0, color: null, dash: null };
  const size = wNum(node, "sz") ?? 4; // eighths of a point
  const color: RGBA | null = parseHexColor(wAttr(node, "color")) ?? null;
  return {
    width: Math.round((size / 8) * 12_700),
    color,
    dash: value === "dashed" || value === "dotted" ? "dash" : null,
  };
}

function parseBorders(container: XmlNode | null): TableBorders | null {
  if (!container) return null;
  const map: [keyof TableBorders, string][] = [
    ["t", "top"],
    ["l", "left"],
    ["l", "start"],
    ["b", "bottom"],
    ["r", "right"],
    ["r", "end"],
  ];
  const borders: TableBorders = {};
  let any = false;
  for (const [key, local] of map) {
    if (borders[key] !== undefined) continue;
    const node = child(container, local);
    if (!node) continue;
    borders[key] = parseBorderEdge(node);
    any = true;
  }
  return any ? borders : null;
}

function cellFill(tcPr: XmlNode | null): Fill | null {
  const fill = wAttr(child(tcPr, "shd"), "fill");
  if (!fill || fill.toLowerCase() === "auto") return null;
  const color = parseHexColor(fill);
  return color ? { type: "solid", color } : null;
}

function cellInsets(tcPr: XmlNode | null, tableMar: XmlNode | null): TableCell["insets"] {
  const mar = child(tcPr, "tcMar") ?? tableMar;
  const read = (local: string, fallback: number): number => {
    const node = mar ? child(mar, local) : null;
    const value = node ? (wNum(node, "w") ?? wNum(node, "val")) : null;
    return twipsToPx(value ?? fallback);
  };
  const left = read("left", DEFAULT_CELL_MARGIN_TWIPS.left);
  const right = read("right", DEFAULT_CELL_MARGIN_TWIPS.right);
  return {
    // `start`/`end` are the writing-direction spellings of left/right.
    l: read("start", DEFAULT_CELL_MARGIN_TWIPS.left) || left,
    t: read("top", DEFAULT_CELL_MARGIN_TWIPS.top),
    r: read("end", DEFAULT_CELL_MARGIN_TWIPS.right) || right,
    b: read("bottom", DEFAULT_CELL_MARGIN_TWIPS.bottom),
  };
}

/** Turn a cell's paragraph blocks into one text body (and the blocks to measure). */
function cellContent(paragraphs: XmlNode[], ctx: TableParseContext): { text: TextBody | null; blocks: DocxParagraphBlock[] } {
  const blocks: DocxParagraphBlock[] = [];
  for (const paragraph of paragraphs) {
    for (const block of parseParagraphBlocks(paragraph, ctx)) {
      if (block.kind === "paragraph") blocks.push(block);
      else if (block.kind === "image") {
        ctx.warn("docx.table.image", "Images inside table cells are not supported and were skipped.");
      }
    }
  }
  if (blocks.length === 0) return { text: null, blocks };
  const first = blocks[0] as DocxParagraphBlock;
  const text: TextBody = {
    ...first.text,
    paragraphs: blocks.map((block) => block.text.paragraphs[0] as TextBody["paragraphs"][number]),
    hasText: blocks.some((block) => block.text.hasText),
  };
  return { text, blocks };
}

function declaredWidths(rows: XmlNode[][], contentWidthEmu: number): number[] | null {
  const firstRow = rows[0];
  if (!firstRow || firstRow.length === 0) return null;
  const widths: number[] = [];
  for (const tc of firstRow) {
    const tcPr = child(tc, "tcPr");
    const tcW = child(tcPr, "tcW");
    const value = tcW ? wNum(tcW, "w") : null;
    if (value === null || value <= 0) return null;
    const type = (tcW ? wAttr(tcW, "type") : "dxa") ?? "dxa";
    widths.push(type === "pct" ? Math.round((value / 5000) * contentWidthEmu) : twipsToEmu(value));
  }
  return widths.length > 0 ? widths : null;
}

function tableWidth(tblPr: XmlNode | null, gridWidths: number[], contentWidthEmu: number): number {
  const declared = child(tblPr, "tblW");
  const value = declared ? wNum(declared, "w") : null;
  const type = (declared ? wAttr(declared, "type") : "dxa") ?? "dxa";
  if (value !== null && value > 0) {
    if (type.toLowerCase() === "pct") return Math.round((value / 5000) * contentWidthEmu);
    if (type.toLowerCase() !== "auto") return twipsToEmu(value);
  }
  const gridTotal = gridWidths.reduce((sum, width) => sum + width, 0);
  return gridTotal > 0 ? gridTotal : Math.max(1, contentWidthEmu);
}

interface PendingCell {
  node: XmlNode;
  gridSpan: number;
  vMerge: "none" | "restart" | "continue";
  text: TextBody | null;
  blocks: DocxParagraphBlock[];
}

function parseVMerge(node: XmlNode | null): "none" | "restart" | "continue" {
  if (!node) return "none";
  return (wVal(node) ?? "").toLowerCase() === "restart" ? "restart" : "continue";
}

export function parseTableBlock(tbl: XmlNode, ctx: TableParseContext): DocxTableBlock | null {
  const tblPr = child(tbl, "tblPr");
  const trNodes = children(tbl, "tr");
  if (trNodes.length === 0) return null;

  const gridWidths: number[] = [];
  for (const col of children(child(tbl, "tblGrid"), "gridCol")) {
    gridWidths.push(twipsToEmu(wNum(col, "w") ?? 0));
  }

  const tableMar = child(tblPr, "tblCellMar");
  const tableBorders = parseBorders(child(tblPr, "tblBorders"));
  let explicitBorders = tableBorders !== null;
  let sawNestedTable = false;

  const rowsOfNodes: XmlNode[][] = [];
  const pending: PendingCell[][] = [];
  for (const tr of trNodes) {
    const tcNodes = children(tr, "tc");
    rowsOfNodes.push(tcNodes);
    const cells: PendingCell[] = [];
    for (const tc of tcNodes) {
      const tcPr = child(tc, "tcPr");
      const paragraphs: XmlNode[] = [];
      for (const item of tc.children) {
        if (item.local === "p") paragraphs.push(item);
        else if (item.local === "tbl") sawNestedTable = true;
      }
      const content = cellContent(paragraphs, ctx);
      cells.push({
        node: tc,
        gridSpan: Math.max(1, Math.round(wNum(child(tcPr, "gridSpan"), "val") ?? 1)),
        vMerge: parseVMerge(child(tcPr, "vMerge")),
        text: content.text,
        blocks: content.blocks,
      });
    }
    pending.push(cells);
  }

  if (sawNestedTable) {
    ctx.warn("docx.table.nested", "Nested tables are not supported and were flattened.");
  }

  /* --------------------------------------------------------- grid geometry */

  const columnCount = (() => {
    let max = 0;
    for (const cells of pending) {
      let count = 0;
      for (const cell of cells) count += cell.gridSpan;
      max = Math.max(max, count);
    }
    return max;
  })();

  const colWidths =
    gridWidths.length > 0
      ? gridWidths
      : (declaredWidths(rowsOfNodes, ctx.contentWidthEmu) ?? []);
  if (colWidths.length === 0) {
    const each = Math.round(ctx.contentWidthEmu / Math.max(1, columnCount));
    for (let i = 0; i < columnCount; i += 1) colWidths.push(each);
    ctx.warn("docx.table.grid", "A table without column widths was laid out in equal columns.");
  } else if (colWidths.length < columnCount) {
    while (colWidths.length < columnCount) colWidths.push(colWidths[colWidths.length - 1] ?? 0);
  }

  const widthEmu = tableWidth(tblPr, colWidths, ctx.contentWidthEmu);
  const columnPx = colWidths.map((emu) => emu / EMU_PER_PX);

  /* -------------------------------------------------------------- v-merges */

  // A continuation cell belongs to the cell above it, so it is dropped from the
  // row and the originating cell is given the resulting row span. That keeps the
  // renderer's column walk aligned without any placeholder cells.
  const rowSpans = new Map<string, number>();
  for (let rowIndex = 0; rowIndex < pending.length; rowIndex += 1) {
    let columnIndex = 0;
    for (const cell of pending[rowIndex] as PendingCell[]) {
      if (cell.vMerge === "restart") {
        let span = 1;
        for (let r = rowIndex + 1; r < pending.length; r += 1) {
          if (pending[r]?.[columnIndex]?.vMerge === "continue") span += 1;
          else break;
        }
        rowSpans.set(`${rowIndex}:${columnIndex}`, span);
      }
      columnIndex += cell.gridSpan;
    }
  }

  /* ------------------------------------------------------------------- rows */

  const rows: TableRow[] = [];
  let declaredHeights = false;
  let anyMeasurement = false;

  for (let rowIndex = 0; rowIndex < pending.length; rowIndex += 1) {
    const trHeight = child(child((trNodes[rowIndex] as XmlNode), "trPr"), "trHeight");
    const declaredPx = trHeight ? twipsToPx(wNum(trHeight, "val") ?? 0) : 0;
    if (declaredPx > 0) declaredHeights = true;

    const cells: TableCell[] = [];
    let columnCursor = 0;
    let contentPx = 0;

    for (const cell of pending[rowIndex] as PendingCell[]) {
      if (cell.vMerge === "continue") {
        // Dropped: the cell above spans these rows, and leaving a placeholder in
        // the list would shift every following cell one column to the right.
        columnCursor += cell.gridSpan;
        continue;
      }

      const tcPr = child(cell.node, "tcPr");
      const borders = parseBorders(child(tcPr, "tcBorders")) ?? tableBorders;
      if (borders && borders !== tableBorders) explicitBorders = true;
      const insets = cellInsets(tcPr, tableMar);
      const vAlign = (wVal(child(tcPr, "vAlign")) ?? "").toLowerCase();
      const available = Math.max(
        columnPx.slice(columnCursor, columnCursor + cell.gridSpan).reduce((sum, width) => sum + width, 0) -
          insets.l -
          insets.r,
        1,
      );
      let cellHeight = 0;
      for (const block of cell.blocks) {
        cellHeight += ctx.measure(block, available);
      }
      if (cellHeight > 0) anyMeasurement = true;
      contentPx = Math.max(contentPx, cellHeight + insets.t + insets.b);

      cells.push({
        gridSpan: cell.gridSpan,
        rowSpan: Math.max(1, rowSpans.get(`${rowIndex}:${columnCursor}`) ?? 1),
        hMerge: false,
        vMerge: cell.vMerge === "restart",
        fill: cellFill(tcPr),
        borders,
        insets,
        anchor: vAlign === "center" ? "ctr" : vAlign === "bottom" ? "b" : "t",
        text: cell.text,
        merged: false,
      });
      columnCursor += cell.gridSpan;
    }

    rows.push({
      height: Math.round(Math.max(declaredPx, contentPx) || FALLBACK_ROW_HEIGHT_PX) * EMU_PER_PX,
      cells,
    });
  }

  const tableData: TableData = {
    colWidths,
    rows,
    firstRow: true,
    bandRow: false,
    styleId: wVal(child(tblPr, "tblStyle")),
  };

  if (!explicitBorders) {
    const neutral: TableBorders = {
      t: { width: 12_700, color: { r: 0xd0, g: 0xd0, b: 0xd0, a: 1 }, dash: null },
      l: { width: 12_700, color: { r: 0xd0, g: 0xd0, b: 0xd0, a: 1 }, dash: null },
      b: { width: 12_700, color: { r: 0xd0, g: 0xd0, b: 0xd0, a: 1 }, dash: null },
      r: { width: 12_700, color: { r: 0xd0, g: 0xd0, b: 0xd0, a: 1 }, dash: null },
    };
    for (const row of rows) {
      for (const cell of row.cells) {
        if (!cell.borders) cell.borders = { ...neutral };
      }
    }
    ctx.warn(
      "docx.table.style",
      "Word table styles are not implemented; a plain grid was drawn for a table that declares no borders.",
    );
  }
  if (!declaredHeights && !anyMeasurement) {
    ctx.warn("docx.table.row-heights", "Table row heights were estimated because the content could not be measured.");
  }

  const jc = (wVal(child(tblPr, "jc")) ?? "").toLowerCase();
  return {
    kind: "table",
    table: tableData,
    heightPx: rows.reduce((sum, row) => sum + row.height / EMU_PER_PX, 0),
    widthEmu,
    align: jc === "center" ? "ctr" : jc === "right" || jc === "end" ? "r" : "l",
    keepWithNext: false,
    part: ctx.part,
  };
}

export type { DocxBlock };
