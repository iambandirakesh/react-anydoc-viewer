/**
 * Text body parsing.
 *
 * A run's effective formatting is the result of a four-step cascade, which is
 * what makes inherited templates render correctly:
 *
 *   slide master text styles  ->  layout placeholder styles  ->  txBody lstStyle
 *     ->  paragraph `defRPr`/`endParaRPr`  ->  the run's own `rPr`
 *
 * Each step only overrides the fields it actually specifies.
 */

import type {
  Bullet,
  BulletText,
  Hyperlink,
  LvlStyle,
  ParaProps,
  Paragraph,
  RGBA,
  Run,
  RunProps,
  TextBody,
} from "../model/types";
import { attr, attrBool, attrNum, child, children, relAttr, type XmlNode } from "./xml";
import { angleUnitsToDegrees, clamp, emuToPx, hundredthsPtToPx, ptToPx } from "./units";
import { isColorElement, resolveColorElement } from "./color";
import type { DrawingContext } from "./drawing";
import { sanitizeUrl } from "../utils/sanitize";

export interface TextContext extends DrawingContext {
  /** level styles inherited from the master / layout for this placeholder */
  inherited?: (LvlStyle | null)[];
  /** default text colour (from the theme or a shape's font reference) */
  defaultColor?: RGBA;
  defaultFont?: string;
  /** relationship id -> hyperlink target resolver */
  resolveHyperlink?: (id: string) => { target: string; external: boolean } | null;
}

const DEFAULT_INSETS = { l: 91_440, t: 45_720, r: 91_440, b: 45_720 };

function mergeRunProps(base: RunProps | undefined, override: RunProps | undefined): RunProps {
  if (!base) return { ...(override ?? {}) };
  if (!override) return { ...base };
  const out: RunProps = { ...base };
  for (const [key, value] of Object.entries(override)) {
    if (value !== undefined) (out as Record<string, unknown>)[key] = value;
  }
  return out;
}

function mergeParaProps(base: ParaProps | undefined, override: ParaProps | undefined): ParaProps {
  if (!base) return { ...(override ?? {}) };
  if (!override) return { ...base };
  const out: ParaProps = { ...base };
  for (const [key, value] of Object.entries(override)) {
    if (value !== undefined) (out as Record<string, unknown>)[key] = value;
  }
  return out;
}

function readColor(node: XmlNode | null, ctx: TextContext): RGBA | undefined {
  if (!node) return undefined;
  for (const candidate of node.children) {
    if (isColorElement(candidate)) {
      return resolveColorElement(candidate, {
        theme: ctx.theme,
        colorMap: ctx.colorMap,
        fallback: ctx.defaultColor,
      });
    }
  }
  return undefined;
}

function readBulletParagraph(
  container: XmlNode | null,
  ctx: TextContext,
): { bullet: Bullet | undefined; color?: RGBA; sizePct?: number; sizePts?: number; font?: string } | undefined {
  if (!container) return undefined;
  const char = child(container, "buChar");
  const autoNum = child(container, "buAutoNum");
  const none = child(container, "buNone");
  const blip = child(container, "buBlip");
  const colorNode = child(container, "buClr");
  const color = readColor(colorNode, ctx);
  const sizePctRaw = attrNum(child(container, "buSzPct"), "val");
  const sizePtsRaw = attrNum(child(container, "buSzPts"), "val");
  const sizePct = sizePctRaw === null ? undefined : clamp(sizePctRaw / 100_000, 0.1, 4);
  const sizePts = sizePtsRaw === null ? undefined : sizePtsRaw / 100;
  const font = attr(child(container, "buFont"), "typeface") ?? undefined;

  if (none) return { bullet: { type: "none" }, color, sizePct, sizePts, font };
  if (char) {
    const bulletText: BulletText = { char: attr(char, "char") ?? "\u2022", color, sizePct, sizePts, font };
    return { bullet: { type: "char", text: bulletText }, color, sizePct, sizePts, font };
  }
  if (autoNum) {
    const bulletText: BulletText = { char: "", color, sizePct, sizePts, font };
    return {
      bullet: {
        type: "autoNum",
        scheme: attr(autoNum, "type") ?? "arabicPeriod",
        startAt: attrNum(autoNum, "startAt") ?? 1,
        text: bulletText,
      },
      color,
      sizePct,
      sizePts,
      font,
    };
  }
  if (blip) {
    const id = relAttr(child(blip, "blip") ?? blip, "embed");
    const rel = id && ctx.resolveRel ? ctx.resolveRel(id) : null;
    if (rel) return { bullet: { type: "image", mediaPath: rel.target }, color, sizePct, sizePts, font };
  }
  return undefined;
}

function parseRunProps(
  node: XmlNode | null,
  ctx: TextContext,
  base?: RunProps,
  hyperlinkId?: string,
): { props: RunProps; link?: Hyperlink } {
  const props: RunProps = { ...(base ?? {}) };
  if (!node) return { props };
  const size = attrNum(node, "sz");
  if (size !== null && size > 0) props.size = hundredthsPtToPx(size);
  const bold = attrBool(node, "b");
  if (bold !== null) props.bold = bold;
  const italic = attrBool(node, "i");
  if (italic !== null) props.italic = italic;
  const underline = attr(node, "u");
  if (underline !== null) props.underline = underline === "none" ? false : underline;
  const strike = attr(node, "strike");
  if (strike !== null) props.strike = strike === "noStrike" ? false : strike;
  const baseline = attrNum(node, "baseline");
  if (baseline !== null) props.baseline = baseline / 100_000;
  const spacing = attrNum(node, "spc");
  if (spacing !== null) props.spacing = ptToPx(spacing / 100);
  const caps = attr(node, "cap");
  if (caps !== null) props.caps = caps === "all" ? "all" : caps === "small" ? "small" : "none";
  const latin = child(node, "latin") ?? child(node, "cs") ?? child(node, "ea");
  const typeface = attr(latin, "typeface");
  if (typeface) props.font = typeface;
  // The colour lives inside `<a:solidFill>`, not directly on `rPr`.
  const solidFill = child(node, "solidFill");
  const colorNode = solidFill ? solidFill.children.find(isColorElement) : undefined;
  const color = colorNode
    ? resolveColorElement(colorNode, { theme: ctx.theme, colorMap: ctx.colorMap, fallback: ctx.defaultColor })
    : undefined;
  if (color) props.color = color;
  const highlight = readColor(child(node, "highlight"), ctx);
  if (highlight) props.highlight = highlight;

  let link: Hyperlink | undefined;
  const hlink = child(node, "hlinkClick");
  const relId = relAttr(hlink, "id") ?? hyperlinkId;
  if (hlink) {
    const tooltip = attr(hlink, "tooltip") ?? undefined;
    if (relId && ctx.resolveHyperlink) {
      const target = ctx.resolveHyperlink(relId);
      if (target && target.external) {
        const result = sanitizeUrl(target.target);
        link = { href: result.href, safe: result.safe, tooltip };
        if (!result.safe && result.reason) {
          ctx.warn?.("hyperlink.blocked", `Blocked hyperlink target (${result.reason})`);
        }
      } else if (target) {
        // Internal slide links are not navigable from an embedded viewer.
        link = { href: "", safe: false, tooltip };
        ctx.warn?.("hyperlink.internal", "Internal slide hyperlinks are not supported");
      } else {
        ctx.warn?.("hyperlink.missing-relationship", "Hyperlink references an unknown relationship");
      }
    } else if (relId === null && attr(hlink, "action")) {
      ctx.warn?.("hyperlink.action", "Action-only hyperlinks are not supported");
    }
  }
  return { props, link };
}

function parseParaProps(node: XmlNode | null, ctx: TextContext): ParaProps {
  const props: ParaProps = {};
  if (!node) return props;
  const level = attrNum(node, "lvl");
  if (level !== null) props.level = clamp(Math.round(level), 0, 8);
  const align = attr(node, "algn");
  if (align) props.align = align as ParaProps["align"];
  const marL = attrNum(node, "marL");
  const marR = attrNum(node, "marR");
  const indent = attrNum(node, "indent");
  if (marL !== null || marR !== null || indent !== null) {
    props.indent = {
      marL: emuToPx(marL ?? 0),
      marR: emuToPx(marR ?? 0),
      indent: emuToPx(indent ?? 0),
    };
  }
  const rtl = attrBool(node, "rtl");
  if (rtl !== null) props.rtl = rtl;

  const lnSpc = child(node, "lnSpc");
  if (lnSpc) {
    const pct = attrNum(child(lnSpc, "spcPct"), "val");
    const pts = attrNum(child(lnSpc, "spcPts"), "val");
    if (pct !== null) props.lineSpacing = { type: "pct", value: clamp(pct / 100_000, 0.1, 10) };
    else if (pts !== null) props.lineSpacing = { type: "pt", value: ptToPx(pts / 100) };
  }
  const before = child(node, "spcBef");
  if (before) {
    const pct = attrNum(child(before, "spcPct"), "val");
    const pts = attrNum(child(before, "spcPts"), "val");
    if (pct !== null) props.spaceBefore = clamp(pct / 100_000, 0, 10) * 12;
    else if (pts !== null) props.spaceBefore = ptToPx(pts / 100);
  }
  const after = child(node, "spcAft");
  if (after) {
    const pct = attrNum(child(after, "spcPct"), "val");
    const pts = attrNum(child(after, "spcPts"), "val");
    if (pct !== null) props.spaceAfter = clamp(pct / 100_000, 0, 10) * 12;
    else if (pts !== null) props.spaceAfter = ptToPx(pts / 100);
  }
  const bullet = readBulletParagraph(node, ctx);
  if (bullet?.bullet) props.bullet = bullet.bullet;
  return props;
}

/** Parse `<a:lstStyle>` into nine level styles. */
export function parseLevelStyles(node: XmlNode | null, ctx: TextContext): (LvlStyle | null)[] {
  const levels: (LvlStyle | null)[] = new Array(9).fill(null);
  if (!node) return levels;
  for (let level = 0; level < 9; level += 1) {
    const lvl = child(node, `lvl${level + 1}pPr`);
    if (!lvl) continue;
    const defRPr = child(lvl, "defRPr");
    const bullet = readBulletParagraph(lvl, ctx);
    const pPr = parseParaProps(lvl, ctx);
    if (bullet?.bullet && !pPr.bullet) pPr.bullet = bullet.bullet;
    levels[level] = {
      defRun: defRPr ? parseRunProps(defRPr, ctx).props : {},
      pPr,
    };
  }
  return levels;
}

function defaultBody(): TextBody {
  return {
    anchor: "t",
    wrap: "square",
    insets: {
      l: emuToPx(DEFAULT_INSETS.l),
      t: emuToPx(DEFAULT_INSETS.t),
      r: emuToPx(DEFAULT_INSETS.r),
      b: emuToPx(DEFAULT_INSETS.b),
    },
    autofit: { type: "none", fontScale: 1, lineSpacingReduction: 0 },
    vertical: "horz",
    paragraphs: [],
    hasText: false,
  };
}

export function parseTextBody(node: XmlNode | null, ctx: TextContext): TextBody | null {
  if (!node) return null;
  const body = defaultBody();

  const bodyPr = child(node, "bodyPr");
  if (bodyPr) {
    const anchor = attr(bodyPr, "anchor");
    if (anchor === "ctr" || anchor === "b" || anchor === "just" || anchor === "dist" || anchor === "t") {
      body.anchor = anchor;
    }
    const wrap = attr(bodyPr, "wrap");
    if (wrap === "none") body.wrap = "none";
    const vertical = attr(bodyPr, "vert");
    if (vertical) body.vertical = vertical;
    body.insets = {
      l: emuToPx(attrNum(bodyPr, "lIns") ?? DEFAULT_INSETS.l),
      t: emuToPx(attrNum(bodyPr, "tIns") ?? DEFAULT_INSETS.t),
      r: emuToPx(attrNum(bodyPr, "rIns") ?? DEFAULT_INSETS.r),
      b: emuToPx(attrNum(bodyPr, "bIns") ?? DEFAULT_INSETS.b),
    };
    const norm = child(bodyPr, "normAutofit");
    const shape = child(bodyPr, "spAutoFit");
    const none = child(bodyPr, "noAutofit");
    if (norm) {
      body.autofit = {
        type: "norm",
        fontScale: clamp((attrNum(norm, "fontScale") ?? 100_000) / 100_000, 0.1, 1),
        lineSpacingReduction: clamp((attrNum(norm, "lnSpcReduction") ?? 0) / 100_000, 0, 1),
      };
    } else if (shape) {
      body.autofit = { type: "shape", fontScale: 1, lineSpacingReduction: 0 };
    } else if (none) {
      body.autofit = { type: "none", fontScale: 1, lineSpacingReduction: 0 };
    }
  }

  const ownLevels = parseLevelStyles(child(node, "lstStyle"), ctx);
  const inherited = ctx.inherited ?? [];

  const paragraphs: Paragraph[] = [];
  let hasText = false;

  for (const paraNode of children(node, "p")) {
    const pPrNode = child(paraNode, "pPr");
    const ownParaProps = parseParaProps(pPrNode, ctx);
    const level = ownParaProps.level ?? 0;
    const inheritedStyle = inherited[level] ?? null;
    const ownStyle = ownLevels[level] ?? null;

    const baseRun = mergeRunProps(inheritedStyle?.defRun, ownStyle?.defRun);
    const paraDefaults = mergeRunProps(baseRun, child(pPrNode, "defRPr") ? parseRunProps(child(pPrNode, "defRPr"), ctx).props : undefined);
    const endPara = child(paraNode, "endParaRPr");
    const runBase = mergeRunProps(paraDefaults, endPara ? parseRunProps(endPara, ctx).props : undefined);

    const mergedPara = mergeParaProps(
      mergeParaProps(inheritedStyle?.pPr, ownStyle?.pPr),
      ownParaProps,
    );

    const runs: Run[] = [];
    for (const runNode of paraNode.children) {
      if (runNode.local === "r") {
        const rPr = child(runNode, "rPr");
        const textNode = child(runNode, "t");
        const { props, link } = parseRunProps(rPr, ctx, runBase);
        const text = textNode ? textNode.text : "";
        if (text.trim().length > 0) hasText = true;
        const run: Run = { text, props };
        if (link) run.link = link;
        runs.push(run);
      } else if (runNode.local === "br") {
        const { props } = parseRunProps(child(runNode, "rPr"), ctx, runBase);
        runs.push({ text: "", br: true, props });
      } else if (runNode.local === "fld") {
        const { props, link } = parseRunProps(child(runNode, "rPr"), ctx, runBase);
        const textNode = child(runNode, "t");
        const text = textNode ? textNode.text : "";
        if (text.trim().length > 0) hasText = true;
        const run: Run = { text, props, field: true };
        if (link) run.link = link;
        runs.push(run);
      }
    }

    if (runs.length === 0) {
      // An empty paragraph still occupies a line, which is how titles get their
      // spacing. Keeping it preserves vertical rhythm.
      runs.push({ text: "", props: runBase });
    }

    paragraphs.push({ props: mergedPara, runs });
  }

  if (paragraphs.length === 0) return null;
  body.paragraphs = paragraphs;
  body.hasText = hasText;
  return body;
}

/** Synthesise a text body for plain text (used for chart/table fallbacks). */
export function plainTextBody(text: string, ctx: TextContext, sizePx = 18): TextBody {
  const body = defaultBody();
  body.paragraphs = [
    {
      props: {},
      runs: [
        {
          text,
          props: { size: sizePx, color: ctx.defaultColor, font: ctx.defaultFont },
        },
      ],
    },
  ];
  body.hasText = text.length > 0;
  return body;
}

export function textBodyPlainText(body: TextBody | null): string {
  if (!body) return "";
  const parts: string[] = [];
  for (const paragraph of body.paragraphs) {
    let line = "";
    for (const run of paragraph.runs) line += run.br ? "\n" : run.text;
    parts.push(line);
  }
  return parts.join("\n");
}

export { angleUnitsToDegrees };
export type { TextBody } from "../model/types";
