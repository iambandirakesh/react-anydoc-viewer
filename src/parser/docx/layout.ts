/**
 * DOCX pagination.
 *
 * Turns the flattened block list into pages of absolutely positioned shapes.
 *
 * Pagination is greedy: blocks are stacked down the page's content area until
 * one would not fit, and the next page starts. Explicit breaks (a page break run,
 * a section break, `w:pageBreakBefore`) always start a new page, and a paragraph
 * marked `w:keepNext` is pulled onto the following page rather than left alone at
 * the bottom of one.
 *
 * The geometry is not guessed: every block's height comes from the measurer, and
 * the shape is given exactly that height, so the text renderer's own flow and the
 * pagination agree.
 */

import type { PictureNode, ShapeNode, ShapeNodeUnion, SlideModel, TableNode } from "../../model/types";
import { EMU_PER_PX, emuToPx } from "../units";
import type { DocxBlock, DocxImageBlock, DocxParagraphBlock, ParagraphMeasurer } from "./model";

export interface PageGeometry {
  /** page size in EMU */
  widthEmu: number;
  heightEmu: number;
  /** margins in EMU */
  margins: { top: number; right: number; bottom: number; left: number };
}

export interface LayoutOptions {
  geometry: PageGeometry;
  measure: ParagraphMeasurer;
  /** part name recorded on each page, used by diagnostics */
  part: string;
  pagePart: (index: number) => string;
}

export interface LayoutResult {
  pages: SlideModel[];
  /** paragraphs taller than a page's content area (they are clipped) */
  overflowed: number;
}

function transform(x: number, y: number, cx: number, cy: number) {
  return {
    x: Math.round(x),
    y: Math.round(y),
    cx: Math.round(cx),
    cy: Math.round(cy),
    rot: 0,
    flipH: false,
    flipV: false,
  };
}

function heightOf(block: DocxBlock, measure: ParagraphMeasurer, contentWidthPx: number): number {
  switch (block.kind) {
    case "paragraph":
      return Math.max(measure(block, contentWidthPx), 0);
    case "table":
      return block.heightPx;
    case "image":
      // A floating picture takes no space in the flow: it is drawn where its
      // anchor asks for it.
      return block.anchor ? 0 : emuToPx(block.height);
    case "break":
      return 0;
    default:
      return 0;
  }
}

function isKeepWithNext(block: DocxBlock): boolean {
  return block.kind === "paragraph" || block.kind === "table" || block.kind === "image"
    ? block.keepWithNext
    : false;
}

export function layoutDocument(blocks: DocxBlock[], options: LayoutOptions): LayoutResult {
  const { geometry, measure } = options;
  const contentWidthEmu = Math.max(geometry.widthEmu - geometry.margins.left - geometry.margins.right, 1);
  const contentWidthPx = contentWidthEmu / EMU_PER_PX;
  const contentHeightPx = Math.max(
    (geometry.heightEmu - geometry.margins.top - geometry.margins.bottom) / EMU_PER_PX,
    1,
  );

  const heights = blocks.map((block) => heightOf(block, measure, contentWidthPx));

  /* ------------------------------------------------------------ pagination */

  const pageIndexes: number[][] = [[]];
  let used = 0;
  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index] as DocxBlock;
    if (block.kind === "break") {
      if ((pageIndexes[pageIndexes.length - 1] as number[]).length > 0) {
        pageIndexes.push([]);
        used = 0;
      }
      continue;
    }
    const height = heights[index] as number;
    if (used > 0 && used + height > contentHeightPx + 0.5) {
      pageIndexes.push([]);
      used = 0;
    }
    (pageIndexes[pageIndexes.length - 1] as number[]).push(index);
    used += height;
  }

  // A `w:keepNext` block is never the last thing on a page.
  for (let page = 0; page < pageIndexes.length - 1; page += 1) {
    const current = pageIndexes[page] as number[];
    while (current.length > 0) {
      const last = current[current.length - 1] as number;
      if (!isKeepWithNext(blocks[last] as DocxBlock)) break;
      current.pop();
      (pageIndexes[page + 1] as number[]).unshift(last);
    }
  }

  /* ---------------------------------------------------------------- shapes */

  const pages: SlideModel[] = [];
  let overflowed = 0;

  pageIndexes.forEach((indexes, pageIndex) => {
    const shapes: ShapeNodeUnion[] = [];
    let cursor = 0;
    let shapeIndex = 0;

    const nextId = () => `docx-p${pageIndex}-${(shapeIndex += 1)}`;

    for (const index of indexes) {
      const block = blocks[index] as DocxBlock;
      if (block.kind === "break") continue;

      if (block.kind === "paragraph") {
        const heightPx = heights[index] as number;
        if (heightPx > contentHeightPx + 0.5) overflowed += 1;
        const shape: ShapeNode = {
          kind: "shape",
          id: nextId(),
          name: block.styleId ? `Paragraph (${block.styleId})` : "Paragraph",
          xfrm: transform(
            geometry.margins.left,
            geometry.margins.top + cursor * EMU_PER_PX,
            contentWidthEmu,
            Math.max(heightPx, 1) * EMU_PER_PX,
          ),
          geometry: { type: "none" },
          fill: null,
          line: null,
          shadow: null,
          text: block.text,
          isPlaceholder: false,
        };
        shapes.push(shape);
        cursor += heightPx;
        continue;
      }

      if (block.kind === "table") {
        const heightPx = heights[index] as number;
        if (heightPx > contentHeightPx + 0.5) overflowed += 1;
        const width = Math.min(block.widthEmu, contentWidthEmu);
        const offset =
          block.align === "ctr" ? (contentWidthEmu - width) / 2 : block.align === "r" ? contentWidthEmu - width : 0;
        const shape: TableNode = {
          kind: "table",
          id: nextId(),
          name: "Table",
          xfrm: transform(
            geometry.margins.left + Math.max(offset, 0),
            geometry.margins.top + cursor * EMU_PER_PX,
            width,
            Math.max(heightPx, 1) * EMU_PER_PX,
          ),
          table: block.table,
        };
        shapes.push(shape);
        cursor += heightPx;
        continue;
      }

      if (block.kind === "image") {
        const image = block as DocxImageBlock;
        const scale = image.width > contentWidthEmu ? contentWidthEmu / image.width : 1;
        const width = image.width * scale;
        const height = image.height * scale;
        let x = geometry.margins.left;
        let y = geometry.margins.top + cursor * EMU_PER_PX;
        if (image.anchor) {
          const horizontal = image.anchorRelative?.h ?? "page";
          const vertical = image.anchorRelative?.v ?? "page";
          x = image.anchor.x + (horizontal === "page" ? 0 : geometry.margins.left);
          y = image.anchor.y + (vertical === "page" ? 0 : geometry.margins.top);
          if (vertical === "paragraph") y += cursor * EMU_PER_PX;
        } else if (image.align === "ctr") {
          x = geometry.margins.left + (contentWidthEmu - width) / 2;
        } else if (image.align === "r") {
          x = geometry.margins.left + (contentWidthEmu - width);
        }
        const shape: PictureNode = {
          kind: "picture",
          id: nextId(),
          name: image.alt ? `Picture (${image.alt})` : "Picture",
          xfrm: transform(x, y, width, height),
          mediaPath: image.mediaPath,
          mediaType: image.mediaType,
          crop: { l: 0, t: 0, r: 0, b: 0 },
          fill: null,
          line: null,
          geometry: { type: "none" },
          shadow: null,
          text: null,
          alt: image.alt,
          unrenderableMedia: null,
        };
        shapes.push(shape);
        cursor += heightOf(block, measure, contentWidthPx);
        continue;
      }
    }

    pages.push({
      index: pageIndex,
      partName: options.pagePart(pageIndex),
      name: null,
      slideshowHidden: false,
      shapes,
      inheritedShapes: [],
      background: { type: "solid", color: { r: 255, g: 255, b: 255, a: 1 } },
      diagnostics: [],
    });
  });

  return { pages, overflowed };
}
