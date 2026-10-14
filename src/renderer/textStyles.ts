/**
 * The style vocabulary shared by text rendering and text measurement.
 *
 * These helpers exist so that the two can never drift: the DOCX layout step
 * measures paragraph heights by building the very same markup with the very same
 * styles the renderer will later use. If they were duplicated, pagination would
 * slowly stop matching the rendered output.
 */

import type { CSSProperties } from "react";
import type { ParaProps, Run, TextBody } from "../model/types";
import { cssColor } from "./style";

export const DEFAULT_FONT_SIZE = 18;
export const DEFAULT_FONT_STACK =
  '"Calibri", "Segoe UI", "Helvetica Neue", Arial, "Noto Sans", sans-serif';

export const DEFAULT_INDENT = { marL: 0, marR: 0, indent: 0 };

export function fontStack(font: string | undefined): string {
  if (!font) return DEFAULT_FONT_STACK;
  const generic = /^(serif|sans-serif|monospace|cursive|fantasy)$/i.test(font);
  const quoted = generic || font.includes(",") ? font : `"${font.replace(/"/g, "")}"`;
  return `${quoted}, ${DEFAULT_FONT_STACK}`;
}

export interface TextStyleOptions {
  /** applied to every run when the viewer is rendering at a reduced scale */
  fontScale?: number;
  /** disables link rendering for thumbnails and print snapshots */
  interactive: boolean;
}

export function runStyleFor(run: Run, baseSize: number, options: TextStyleOptions): CSSProperties {
  const props = run.props;
  const size = (props.size ?? baseSize) * (options.fontScale ?? 1);
  const style: CSSProperties = {
    fontSize: `${size}px`,
    fontFamily: fontStack(props.font),
    fontWeight: props.bold ? 700 : 400,
    fontStyle: props.italic ? "italic" : "normal",
    color: cssColor(props.color),
    letterSpacing: props.spacing ? `${props.spacing}px` : undefined,
  };
  if (props.underline && props.underline !== "none") {
    style.textDecoration = props.strike && props.strike !== "noStrike" ? "underline line-through" : "underline";
    if (props.underline === "dbl") style.textDecorationStyle = "double";
  } else if (props.strike && props.strike !== "noStrike") {
    style.textDecoration = "line-through";
  }
  if (props.caps === "all") style.textTransform = "uppercase";
  if (props.caps === "small") style.fontVariant = "small-caps";
  if (props.highlight) style.backgroundColor = cssColor(props.highlight);
  if (props.baseline && props.baseline !== 0) {
    style.position = "relative";
    style.top = `${-props.baseline * size}px`;
    style.fontSize = `${size * 0.75}px`;
  }
  return style;
}

/**
 * The style of one paragraph inside a text body.
 *
 * Margins are written as longhands on purpose: the `margin` shorthand is not
 * reliably applied by every React version, and a stray user-agent paragraph
 * margin would shift every paragraph in the body.
 */
export function paragraphStyleFor(props: ParaProps, body: TextBody, baseSize: number): CSSProperties {
  const indent = props.indent ?? DEFAULT_INDENT;
  const align = props.align ?? "l";
  const textAlign: CSSProperties["textAlign"] =
    align === "ctr"
      ? "center"
      : align === "r"
        ? "right"
        : align === "just" || align === "dist" || align === "justLow"
          ? "justify"
          : "left";

  const lineHeight =
    props.lineSpacing?.type === "pt"
      ? `${props.lineSpacing.value}px`
      : props.lineSpacing?.type === "pct"
        ? props.lineSpacing.value
        : undefined;

  return {
    position: "relative",
    marginTop: props.spaceBefore ? `${props.spaceBefore}px` : 0,
    marginRight: `${indent.marR}px`,
    marginBottom: props.spaceAfter ? `${props.spaceAfter}px` : 0,
    marginLeft: `${indent.marL}px`,
    textAlign,
    lineHeight,
    whiteSpace: body.wrap === "none" ? "nowrap" : "pre-wrap",
    direction: props.rtl ? "rtl" : undefined,
    fontSize: `${baseSize}px`,
    fontFamily: DEFAULT_FONT_STACK,
  };
}

/** The flex column a text body's paragraphs flow in. */
export function textContainerStyle(body: TextBody): CSSProperties {
  const anchor = body.anchor;
  return {
    display: "flex",
    flexDirection: "column",
    justifyContent: anchor === "ctr" ? "center" : anchor === "b" ? "flex-end" : "flex-start",
    width: "100%",
    height: "100%",
    boxSizing: "border-box",
    padding: `${body.insets.t}px ${body.insets.r}px ${body.insets.b}px ${body.insets.l}px`,
    overflow: "visible",
  };
}

/** The base font size a text body lays its runs out with. */
export function baseSizeOf(body: TextBody): number {
  return (body.paragraphs[0]?.runs[0]?.props.size ?? DEFAULT_FONT_SIZE) as number;
}
