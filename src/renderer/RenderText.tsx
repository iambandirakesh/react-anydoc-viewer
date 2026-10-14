/**
 * Text body rendering.
 *
 * HTML is used rather than SVG `<text>` so that line breaking, wrapping and
 * accessibility behave the way they do everywhere else on the web; the shape's
 * geometry is drawn separately in SVG underneath.
 */

import type { CSSProperties, ReactElement, ReactNode } from "react";
import type { Bullet, ParaProps, Run, TextBody } from "../model/types";
import { cssColor } from "./style";
import {
  DEFAULT_FONT_SIZE,
  DEFAULT_FONT_STACK,
  DEFAULT_INDENT,
  baseSizeOf,
  fontStack,
  paragraphStyleFor,
  runStyleFor,
  textContainerStyle,
  type TextStyleOptions,
} from "./textStyles";

export type RenderTextOptions = TextStyleOptions;

function formatAutoNumber(scheme: string, value: number): string {
  const romanize = (num: number): string => {
    const table: [number, string][] = [
      [1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"],
      [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"],
    ];
    let remaining = num;
    let out = "";
    for (const [size, numeral] of table) {
      while (remaining >= size) {
        out += numeral;
        remaining -= size;
      }
    }
    return out;
  };
  const alpha = (num: number): string => {
    let out = "";
    let value2 = num;
    while (value2 > 0) {
      const remainder = (value2 - 1) % 26;
      out = String.fromCharCode(97 + remainder) + out;
      value2 = Math.floor((value2 - 1) / 26);
    }
    return out;
  };

  const normalized = scheme.toLowerCase();
  if (normalized.startsWith("romanuc")) return romanize(value).toUpperCase();
  if (normalized.startsWith("romanlc")) return romanize(value).toLowerCase();
  if (normalized.startsWith("alphauc")) return alpha(value).toUpperCase();
  if (normalized.startsWith("alphalc")) return alpha(value);
  // arabicPeriod, arabicParenR, circleNumDbPlain, ideograph… all render as digits.
  return String(value);
}

function bulletSuffix(scheme: string): string {
  const lower = scheme.toLowerCase();
  if (lower.includes("parenboth")) return ")";
  if (lower.includes("parenr")) return ")";
  if (lower.includes("period")) return ".";
  if (lower.includes("plain")) return "";
  if (lower.includes("endash")) return "–";
  if (lower.includes("emdash")) return "—";
  return ".";
}

interface ParagraphContext {
  /** 1-based index within the current numbered-list run */
  number: number;
  baseSize: number;
}

function renderBullet(
  bullet: Bullet,
  paragraph: ParaProps,
  ctx: ParagraphContext,
  options: RenderTextOptions,
): ReactNode {
  const indent = paragraph.indent ?? DEFAULT_INDENT;
  const hang = -indent.indent;
  if (bullet.type === "none") return null;
  if (bullet.type === "image") {
    // Image bullets need the media store; the caller resolves them beforehand.
    return null;
  }

  let text: string;
  if (bullet.type === "char") {
    text = bullet.text.char;
  } else {
    text = `${formatAutoNumber(bullet.scheme, bullet.startAt + ctx.number - 1)}${bulletSuffix(bullet.scheme)}`;
  }

  const bulletSpec = bullet.text;
  const size =
    (bulletSpec.sizePts !== undefined
      ? bulletSpec.sizePts * (96 / 72)
      : ctx.baseSize * ((bulletSpec.sizePct ?? 100) / 100)) * (options.fontScale ?? 1);

  const style: CSSProperties = {
    position: "absolute",
    top: 0,
    left: 0,
    width: `${Math.max(hang, 0)}px`,
    transform: hang < 0 ? "translateX(-100%)" : undefined,
    textAlign: hang < 0 ? "right" : "left",
    fontSize: `${size}px`,
    fontFamily: fontStack(bulletSpec.font),
    color: cssColor(bulletSpec.color),
    whiteSpace: "pre",
    pointerEvents: "none",
  };

  const offset = hang < 0 ? style : { ...style, left: `${indent.indent}px`, width: `${Math.max(hang, DEFAULT_FONT_SIZE)}px` };
  return <span style={offset}>{text}</span>;
}

export interface TextRenderResult {
  node: ReactNode;
  /** true when any visible run exists */
  visible: boolean;
}

export function renderTextBody(
  body: TextBody,
  buildLink: (href: string, tooltip: string | undefined, children: ReactNode, key: string) => ReactNode,
  options: RenderTextOptions,
): TextRenderResult {
  const baseSize = baseSizeOf(body);
  const blocks: ReactNode[] = [];
  let numberRun = 0;

  body.paragraphs.forEach((paragraph, paragraphIndex) => {
    const props = paragraph.props;
    const indent = props.indent ?? DEFAULT_INDENT;

    const isNumbered = props.bullet?.type === "autoNum";
    if (isNumbered) numberRun += 1;
    else if (props.bullet && props.bullet.type !== "autoNum") numberRun = 0;

    const paragraphStyle = paragraphStyleFor(props, body, baseSize);

    const runs: ReactNode[] = [];
    paragraph.runs.forEach((run, runIndex) => {
      const key = `p${paragraphIndex}-r${runIndex}`;
      if (run.br) {
        runs.push(<br key={key} />);
        return;
      }
      if (run.text.length === 0) return;
      const child = (
        <span key={key} style={runStyleFor(run, baseSize, options)} data-run="true">
          {run.text}
        </span>
      );
      if (run.link && run.link.safe && options.interactive) {
        runs.push(buildLink(run.link.href, run.link.tooltip, child, key));
      } else {
        runs.push(child);
      }
    });

    const bullet = props.bullet ? renderBullet(props.bullet, props, { number: numberRun, baseSize }, options) : null;

    blocks.push(
      <p key={`p${paragraphIndex}`} style={paragraphStyle} data-paragraph={paragraphIndex}>
        {bullet}
        <span style={indent.indent !== 0 ? { display: "inline-block", textIndent: `${indent.indent}px` } : undefined}>
          {runs}
        </span>
      </p>,
    );
  });

  return {
    node: <div style={textContainerStyle(body)}>{blocks}</div>,
    visible: body.hasText || body.paragraphs.length > 0,
  };
}

export { DEFAULT_FONT_STACK, DEFAULT_FONT_SIZE, fontStack };
export type { ReactElement };
