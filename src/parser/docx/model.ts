/**
 * The intermediate block model for WordprocessingML.
 *
 * A `.docx` is a flowing document, while the renderer draws absolutely
 * positioned shapes. The bridge is this list of blocks: the parser produces
 * blocks in document order and the layout step decides which page each block
 * lands on and at what offset, before turning them into the same
 * {@link SlideModel} shapes the presentation renderer already understands.
 */

import type { TableData, TextBody } from "../../model/types";

/** A single Word paragraph, already resolved to the model's text shape. */
export interface DocxParagraphBlock {
  kind: "paragraph";
  text: TextBody;
  /** style id the paragraph was formatted with (diagnostics only) */
  styleId: string | null;
  /** `w:keepNext`: do not leave this paragraph alone at the bottom of a page */
  keepWithNext: boolean;
  part: string;
}

/** A Word table, with row heights already resolved. */
export interface DocxTableBlock {
  kind: "table";
  table: TableData;
  /** total height in CSS px (declared row heights or measured content) */
  heightPx: number;
  /** width in EMU */
  widthEmu: number;
  align: "l" | "ctr" | "r";
  keepWithNext: boolean;
  part: string;
}

/** A picture. When `anchor` is set it floats at a fixed page offset. */
export interface DocxImageBlock {
  kind: "image";
  mediaPath: string | null;
  mediaType: string | null;
  alt: string | null;
  /** extent in EMU */
  width: number;
  height: number;
  /** page offset in EMU for floating pictures */
  anchor: { x: number; y: number } | null;
  /** what the offset is measured from (`w:relativeFrom`) */
  anchorRelative: { h: string; v: string } | null;
  align: "l" | "ctr" | "r";
  keepWithNext: boolean;
  part: string;
}

/** An explicit page break (`w:br w:type="page"`, section break, `w:pageBreakBefore`). */
export interface DocxBreakBlock {
  kind: "break";
  part: string;
}

export type DocxBlock = DocxParagraphBlock | DocxTableBlock | DocxImageBlock | DocxBreakBlock;

/**
 * Height of a paragraph block at a given content width, in CSS px.
 *
 * The browser implementation measures real layout; the fallback estimates from
 * the text metrics so a document is still paginated sensibly where no layout
 * engine is available (SSR, unit tests).
 */
export type ParagraphMeasurer = (block: DocxParagraphBlock, widthPx: number) => number;

/** Default font size for a paragraph block, used when a run carries none. */
export const DEFAULT_DOCX_FONT_SIZE = 14.6667; // 11pt, Word's body default
