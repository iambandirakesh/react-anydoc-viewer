/**
 * Picture extraction from `w:drawing`.
 *
 * Both inline and floating pictures end up as ordinary model pictures; a
 * floating picture keeps its page offset so it can be placed absolutely on the
 * page its anchor paragraph lands on. Text wrapping around floating pictures is
 * not implemented (the picture is drawn over the text), which is reported as a
 * diagnostic rather than hidden.
 */

import { attr, child, find, relAttr, type XmlNode } from "../xml";
import { wAttr } from "./xmlw";

export interface DocxImage {
  mediaPath: string | null;
  /** extent in EMU */
  width: number;
  height: number;
  /** page offset in EMU when the picture floats, otherwise null */
  anchor: { x: number; y: number } | null;
  /** what `anchor` is measured from (`w:relativeFrom`), defaulting to the page */
  anchorRelative: { h: string; v: string } | null;
  /** horizontal alignment requested for an inline picture */
  align: "l" | "ctr" | "r";
  alt: string | null;
  /** true when the picture asks for text to wrap around it */
  wraps: boolean;
}

export interface DrawingContext {
  resolveRel(id: string): { target: string; external: boolean } | null;
  warn(code: string, message: string): void;
  part: string;
}

function num(value: string | null): number | null {
  if (value === null) return null;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : null;
}

function alignmentFrom(value: string | null): "l" | "ctr" | "r" {
  switch ((value ?? "").toLowerCase()) {
    case "center":
      return "ctr";
    case "right":
    case "end":
      return "r";
    default:
      return "l";
  }
}

function offsetOf(container: XmlNode | null): number | null {
  if (!container) return null;
  const offset = child(container, "posOffset");
  if (!offset) return null;
  const value = num(offset.text.trim());
  return value;
}

/** Read a `w:drawing` element into a picture description. */
export function parseDrawing(drawing: XmlNode, ctx: DrawingContext): DocxImage | null {
  const inline = child(drawing, "inline");
  const anchor = child(drawing, "anchor");
  const container = inline ?? anchor;
  if (!container) {
    ctx.warn("docx.drawing.unsupported", "A drawing could not be interpreted and was skipped.");
    return null;
  }

  const extent = child(container, "extent");
  const width = num(attr(extent, "cx")) ?? 0;
  const height = num(attr(extent, "cy")) ?? 0;
  if (width <= 0 || height <= 0) {
    ctx.warn("docx.drawing.size", "A picture without a usable size was skipped.");
    return null;
  }

  const blip = find(container, "blip");
  const relId = relAttr(blip, "embed") ?? relAttr(blip, "link");
  const relationship = relId ? ctx.resolveRel(relId) : null;
  if (!relationship) {
    ctx.warn("docx.picture.missing-relationship", "A picture references an image that is not in the package.");
  }

  const docPr = find(container, "docPr");
  const altText = wAttr(docPr, "descr") ?? wAttr(docPr, "title") ?? wAttr(docPr, "name") ?? null;

  let pageAnchor: { x: number; y: number } | null = null;
  let anchorRelative: { h: string; v: string } | null = null;
  let wraps = false;
  if (anchor) {
    const positionH = child(anchor, "positionH");
    const positionV = child(anchor, "positionV");
    const x = offsetOf(positionH);
    const y = offsetOf(positionV);
    anchorRelative = {
      h: (wAttr(positionH, "relativeFrom") ?? "page").toLowerCase(),
      v: (wAttr(positionV, "relativeFrom") ?? "page").toLowerCase(),
    };
    const wrapNone = child(anchor, "wrapNone") !== null;
    wraps = !wrapNone;
    if (x !== null && y !== null) pageAnchor = { x, y };
    if (x === null || y === null) {
      ctx.warn(
        "docx.picture.alignment",
        "A floating picture is positioned relative to the text and was placed in the flow instead.",
      );
    }
    if (wraps && pageAnchor) {
      ctx.warn(
        "docx.picture.wrap",
        "Text does not wrap around floating pictures; the picture is drawn over the text.",
      );
    }
  }

  const positionH = child(container, "positionH");
  return {
    mediaPath: relationship?.target ?? null,
    width,
    height,
    anchor: pageAnchor,
    anchorRelative,
    align: alignmentFrom(positionH ? wAttr(positionH, "align") : null),
    alt: altText && altText.length > 0 ? altText : null,
    wraps,
  };
}
