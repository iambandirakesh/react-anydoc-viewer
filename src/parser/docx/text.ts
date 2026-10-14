/**
 * Paragraph and run parsing.
 *
 * A Word paragraph maps almost directly onto the model's `TextBody` — same
 * alignment, indent, spacing, bullet and hyperlink vocabulary — so paragraphs are
 * re-encoded rather than reinterpreted, and the existing text renderer can draw
 * them.
 *
 * Two Word features have no counterpart in the model, and both are handled by
 * splitting the paragraph into blocks instead of dropping content:
 *
 *  - a picture inside a run ends the current text fragment and starts an image
 *    block, so a paragraph with an inline icon still renders both;
 *  - an explicit page break (`w:br w:type="page"`, a section break) ends the
 *    fragment and emits a break block for the layout step.
 */

import type { Hyperlink, ParaProps, Paragraph, Run, RunProps, TextBody } from "../../model/types";
import { child, type XmlNode } from "../xml";
import { sanitizeUrl } from "../../utils/sanitize";
import { parseDrawing } from "./drawing";
import type { DocxBlock, DocxImageBlock, DocxParagraphBlock } from "./model";
import type { DocxStyles, WordParaStyle } from "./styles";
import { parseParagraphProperties, parseRunProperties } from "./styles";
import { wAttr, wVal } from "./xmlw";

export interface DocxTextContext {
  styles: DocxStyles;
  /** part the paragraphs came from, for diagnostics */
  part: string;
  resolveRel(id: string): { target: string; external: boolean } | null;
  warn(code: string, message: string): void;
}

/** Word content wrappers that only group runs (`w:ins`, `w:smartTag`, …). */
const RUN_CONTAINERS = new Set(["ins", "smartTag", "sdt", "fldSimple", "dir", "bdo", "moveTo"]);

function mergeRunProps(base: RunProps | undefined, override: RunProps | undefined): RunProps {
  if (!base) return { ...(override ?? {}) };
  if (!override) return { ...base };
  const merged: RunProps = { ...base };
  for (const [key, value] of Object.entries(override)) {
    if (value !== undefined) (merged as Record<string, unknown>)[key] = value;
  }
  return merged;
}

function mergeParaProps(base: ParaProps | undefined, override: ParaProps | undefined): ParaProps {
  if (!base) return { ...(override ?? {}) };
  if (!override) return { ...base };
  const merged: ParaProps = { ...base };
  for (const [key, value] of Object.entries(override)) {
    if (value !== undefined && value !== null) (merged as Record<string, unknown>)[key] = value;
  }
  return merged;
}

function emptyTextBody(paragraph: Paragraph): TextBody {
  return {
    anchor: "t",
    wrap: "square",
    // Paragraph blocks are measured and sized individually, so the text body
    // carries no inset of its own: the paragraph's own spacing is the rhythm.
    insets: { l: 0, t: 0, r: 0, b: 0 },
    autofit: { type: "none", fontScale: 1, lineSpacingReduction: 0 },
    vertical: "horz",
    paragraphs: [paragraph],
    hasText: paragraph.runs.some((run) => run.text.trim().length > 0),
  };
}

function symbolCharacter(node: XmlNode): string | null {
  const raw = wAttr(node, "char");
  if (!raw) return null;
  const code = parseInt(raw, 16);
  if (!Number.isFinite(code) || code <= 0 || code > 0x10ffff) return null;
  try {
    return String.fromCodePoint(code);
  } catch {
    return null;
  }
}

/** Resolve a paragraph's effective paragraph properties and run defaults. */
function resolveParagraph(
  pPr: XmlNode | null,
  ctx: DocxTextContext,
): { para: ParaProps; directFlags: WordParaStyle; styleId: string | null; runBase: RunProps } {
  const direct = parseParagraphProperties(pPr);
  const styleId = direct.styleId;
  const style = ctx.styles.paragraphStyle(styleId);
  const para = mergeParaProps(mergeParaProps(ctx.styles.defaultPara, style?.para), direct.props);

  const numId = direct.numId ?? style?.numId ?? null;
  const ilvl = direct.numId !== null ? direct.ilvl : (style?.ilvl ?? 0);
  if (numId !== null) {
    const numbering = ctx.styles.numbering(numId, ilvl);
    if (numbering) {
      if (numbering.bullet) para.bullet = numbering.bullet;
      if (numbering.indent && !direct.directIndent) para.indent = numbering.indent;
    } else {
      ctx.warn("docx.numbering.missing", "A numbered paragraph references numbering that is not defined.");
    }
  }

  if (direct.hasBorders) {
    ctx.warn("docx.paragraph.borders", "Paragraph borders are not supported and were omitted.");
  }
  if (direct.hasShading) {
    ctx.warn("docx.paragraph.shading", "Paragraph shading is not supported and was omitted.");
  }

  const runBase = mergeRunProps(
    mergeRunProps(ctx.styles.defaultRun, style?.run),
    direct.runProps,
  );
  return { para, directFlags: direct, styleId, runBase };
}

/**
 * Parse one `w:p` into blocks: text fragments, pictures and page breaks, in
 * document order.
 */
export function parseParagraphBlocks(p: XmlNode, ctx: DocxTextContext): DocxBlock[] {
  const pPr = child(p, "pPr");
  const { para, runBase, styleId, directFlags } = resolveParagraph(pPr, ctx);
  const part = ctx.part;

  const blocks: DocxBlock[] = [];
  let runs: Run[] = [];
  let sawContent = false;
  let sawBreak = false;

  // `w:pageBreakBefore` starts a page. A break on an otherwise empty page is a
  // no-op in the layout step, so this is safe for the document's first paragraph.
  if (directFlags.pageBreakBefore) {
    sawBreak = true;
    blocks.push({ kind: "break", part });
  }

  const flush = () => {
    if (runs.length === 0) return;
    const paragraph: Paragraph = { props: para, runs };
    const block: DocxParagraphBlock = {
      kind: "paragraph",
      text: emptyTextBody(paragraph),
      styleId,
      keepWithNext: directFlags.keepNext,
      part,
    };
    blocks.push(block);
    runs = [];
  };

  const addRun = (run: Run) => {
    sawContent = true;
    runs.push(run);
  };

  const addBreak = () => {
    flush();
    sawBreak = true;
    blocks.push({ kind: "break", part });
  };

  const addImage = (image: DocxImageBlock) => {
    flush();
    sawContent = true;
    blocks.push(image);
  };

  const consumeRunContent = (container: XmlNode, props: RunProps, link: Hyperlink | undefined) => {
    for (const item of container.children) {
      switch (item.local) {
        case "t":
          if (item.text.length === 0) break;
          addRun(link ? { text: item.text, props, link } : { text: item.text, props });
          break;
        case "br": {
          const type = (wAttr(item, "type") ?? "text").toLowerCase();
          if (type === "page") addBreak();
          else addRun({ text: "", br: true, props });
          break;
        }
        case "tab":
          addRun({ text: "\t", props });
          break;
        case "sym": {
          const character = symbolCharacter(item);
          if (character) addRun({ text: character, props });
          break;
        }
        case "noBreakHyphen":
          addRun({ text: "\u2011", props });
          break;
        case "drawing": {
          const image = parseDrawing(item, { resolveRel: ctx.resolveRel, warn: ctx.warn, part: ctx.part });
          if (!image) break;
          addImage({
            kind: "image",
            mediaPath: image.mediaPath,
            mediaType: null,
            alt: image.alt,
            width: image.width,
            height: image.height,
            anchor: image.anchor,
            anchorRelative: image.anchorRelative,
            align: image.align,
            keepWithNext: false,
            part: ctx.part,
          });
          break;
        }
        case "pict":
          ctx.warn("docx.vml", "VML pictures are not supported and were skipped.");
          break;
        case "object":
          ctx.warn("docx.embedded-object", "Embedded OLE objects are not supported and were skipped.");
          break;
        case "AlternateContent": {
          const branch = child(item, "Fallback") ?? child(item, "Choice");
          if (branch) consumeRunContent(branch, props, link);
          break;
        }
        case "fldChar":
        case "instrText":
          // Field plumbing: the cached result is what the reader should see.
          break;
        default:
          if (RUN_CONTAINERS.has(item.local)) consumeRunContent(item, props, link);
          break;
      }
    }
  };

  const consumeRun = (r: XmlNode, link: Hyperlink | undefined) => {
    const rPr = child(r, "rPr");
    const characterStyle = ctx.styles.characterStyle(wVal(child(rPr, "rStyle")));
    const props = parseRunProperties(rPr, mergeRunProps(runBase, characterStyle ?? undefined));
    consumeRunContent(r, props, link);
  };

  const resolveHyperlink = (node: XmlNode): Hyperlink | undefined => {
    const tooltip = wAttr(node, "tooltip") ?? undefined;
    const relId = wAttr(node, "id");
    if (!relId) {
      if (wAttr(node, "anchor")) {
        // A bookmark target has no URL and there is nothing to navigate to.
        ctx.warn("docx.hyperlink.bookmark", "Bookmark hyperlinks are not navigable and render as plain text.");
      }
      return undefined;
    }
    const relationship = ctx.resolveRel(relId);
    if (!relationship) {
      ctx.warn("docx.hyperlink.missing-relationship", "A hyperlink references a relationship that is missing.");
      return undefined;
    }
    if (!relationship.external) {
      ctx.warn("docx.hyperlink.internal", "Hyperlinks to other files are not navigable and render as plain text.");
      return undefined;
    }
    const result = sanitizeUrl(relationship.target);
    if (!result.safe) {
      ctx.warn("docx.hyperlink.blocked", `A hyperlink target was blocked (${result.reason ?? "unsafe"}).`);
      return { href: "", safe: false, tooltip };
    }
    return { href: result.href, safe: true, tooltip };
  };

  const consumeNode = (container: XmlNode) => {
    for (const item of container.children) {
      switch (item.local) {
        case "r":
          consumeRun(item, undefined);
          break;
        case "hyperlink":
          consumeHyperlink(item);
          break;
        case "del":
        case "moveFrom":
          break;
        case "oMath":
        case "oMathPara":
          ctx.warn("docx.equation", "Equations are not supported and were skipped.");
          break;
        case "pict":
          ctx.warn("docx.vml", "VML pictures are not supported and were skipped.");
          break;
        case "AlternateContent": {
          const branch = child(item, "Fallback") ?? child(item, "Choice");
          if (branch) consumeNode(branch);
          break;
        }
        default:
          if (RUN_CONTAINERS.has(item.local)) consumeNode(item);
          break;
      }
    }
  };

  const consumeHyperlink = (node: XmlNode) => {
    const link = resolveHyperlink(node);
    for (const item of node.children) {
      if (item.local === "r") consumeRun(item, link);
      else if (RUN_CONTAINERS.has(item.local)) consumeHyperlink(item);
    }
  };

  consumeNode(p);

  // A paragraph with no content still takes a line in Word.
  if (!sawContent && !sawBreak && blocks.length === 0) {
    addRun({ text: "", props: runBase });
  }
  flush();

  return blocks;
}

export { mergeRunProps, mergeParaProps };
