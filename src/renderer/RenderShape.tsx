/**
 * Shape rendering.
 *
 * Each shape becomes an absolutely positioned box containing an SVG (geometry,
 * fills, outlines, image fills) with an HTML text layer on top. That hybrid
 * keeps vector geometry exact while letting the browser lay out text, which is
 * what makes wrapped and bulleted text behave the way it does in PowerPoint.
 *
 * Coordinates are the shape's own pixel box, so `viewBox` maps 1:1 and the CSS
 * transform handles rotation/flipping in one place.
 */

import { useState, type CSSProperties, type ReactElement, type ReactNode } from "react";
import type {
  ChartNode,
  ConnectorNode,
  Fill,
  Geometry,
  GroupNode,
  LineStyle,
  PictureNode,
  ShapeNode,
  ShapeNodeUnion,
  Shadow,
  TableNode,
  Transform,
  UnsupportedNode,
} from "../model/types";
import { emuToPx } from "../parser/units";
import { customGeometryToPath, presetToPath, type GeometryPath } from "../parser/geometry";
import type { MediaStore } from "../parser/media";
import { cssShadowFilter, gradientIdFor, hashString, paintForFill, svgPaint } from "./style";
import { renderTextBody } from "./RenderText";
import { renderTable } from "./RenderTable";
import { renderChart } from "./RenderChart";

export interface ShapeRenderContext {
  media: MediaStore;
  /** namespace for generated element ids (stable across server/client) */
  scope: string;
  /** false for thumbnails: no links, no pointer interaction */
  interactive: boolean;
  palette: string[];
  warn: (code: string, message: string) => void;
}

const PLACEHOLDER_STYLE: CSSProperties = {
  position: "absolute",
  inset: 0,
  display: "flex",
  alignItems: "center",
  justifyContent: "center",
  border: "1px dashed rgba(120, 120, 120, 0.75)",
  backgroundColor: "rgba(246, 246, 246, 0.9)",
  color: "#5f6368",
  fontSize: "11px",
  fontFamily: "system-ui, -apple-system, sans-serif",
  textAlign: "center",
  padding: "4px",
  boxSizing: "border-box",
  overflow: "hidden",
  lineHeight: 1.3,
  whiteSpace: "pre-line",
};

/** Minimal input for the geometry layer: shapes, connectors and pictures all fit. */
interface GeometryLayerInput {
  id: string;
  xfrm: Transform;
  geometry: Geometry;
  /** connectors and picture outlines have no shape fill */
  fill?: Fill | null;
  line: LineStyle | null;
}

export type LinkBuilder = (
  href: string,
  tooltip: string | undefined,
  children: ReactNode,
  key: string,
) => ReactNode;

export function createLinkBuilder(interactive: boolean): LinkBuilder {
  return (href, tooltip, children, key) => {
    if (!interactive) return children;
    return (
      <a
        key={key}
        href={href}
        title={tooltip}
        target="_blank"
        rel="noopener noreferrer"
        style={{ color: "#0563c1", textDecoration: "underline" }}
      >
        {children}
      </a>
    );
  };
}

function wrapperStyle(node: ShapeNodeUnion, shadow: Shadow | null | undefined): CSSProperties {
  const style: CSSProperties = {
    position: "absolute",
    left: `${emuToPx(node.xfrm.x)}px`,
    top: `${emuToPx(node.xfrm.y)}px`,
    width: `${emuToPx(node.xfrm.cx)}px`,
    height: `${emuToPx(node.xfrm.cy)}px`,
  };
  const transform: string[] = [];
  if (node.xfrm.rot) transform.push(`rotate(${node.xfrm.rot}deg)`);
  if (node.xfrm.flipH || node.xfrm.flipV) {
    transform.push(`scale(${node.xfrm.flipH ? -1 : 1}, ${node.xfrm.flipV ? -1 : 1})`);
  }
  if (transform.length > 0) {
    style.transform = transform.join(" ");
    style.transformOrigin = "center center";
  }
  const filter = cssShadowFilter(shadow);
  if (filter) style.filter = filter.filter;
  return style;
}

function geometryPath(node: GeometryLayerInput, context: ShapeRenderContext): GeometryPath | null {
  const width = emuToPx(node.xfrm.cx);
  const height = emuToPx(node.xfrm.cy);
  if (width <= 0 || height <= 0) return null;
  const { geometry } = node;
  if (geometry.type === "none") return null;
  if (geometry.type === "custom") {
    return customGeometryToPath(geometry.commands, width, height, geometry.w, geometry.h);
  }
  const path = presetToPath(geometry.preset, geometry.adjust, width, height);
  if (path) return path;
  context.warn(
    "geometry.preset-unsupported",
    `Preset shape "${geometry.preset}" is not implemented and was drawn as a rectangle.`,
  );
  return presetToPath("rect", {}, width, height);
}

interface PaintResult {
  defs: ReactNode[];
  fillValue: string;
  fillOpacity: number;
  usesImage: boolean;
}

function resolvePaint(
  fill: Fill | null,
  nodeId: string,
  context: ShapeRenderContext,
  box: { width: number; height: number },
): PaintResult {
  if (!fill || fill.type === "none") {
    return { defs: [], fillValue: "none", fillOpacity: 1, usesImage: false };
  }

  if (fill.type === "gradient") {
    const id = gradientIdFor(context.scope, nodeId, fill);
    const stops = fill.stops.map((stop, index) => (
      <stop
        key={index}
        offset={`${Math.round(stop.pos * 100)}%`}
        stopColor={`rgb(${Math.round(stop.color.r)}, ${Math.round(stop.color.g)}, ${Math.round(stop.color.b)})`}
        stopOpacity={stop.color.a}
      />
    ));
    const def = fill.radial ? (
      <radialGradient key={id} id={id} cx="50%" cy="50%" r="50%">
        {stops}
      </radialGradient>
    ) : (
      <linearGradient
        key={id}
        id={id}
        gradientUnits="objectBoundingBox"
        x1="0%"
        y1="0%"
        x2="100%"
        y2="0%"
        gradientTransform={`rotate(${fill.angle}, 0.5, 0.5)`}
      >
        {stops}
      </linearGradient>
    );
    return { defs: [def], fillValue: `url(#${id})`, fillOpacity: 1, usesImage: false };
  }

  if (fill.type === "image") {
    const url = context.media.url(fill.mediaPath);
    if (!url) {
      context.warn("fill.image-missing", `Image fill "${fill.mediaPath}" is not available`);
      return { defs: [], fillValue: "none", fillOpacity: 1, usesImage: false };
    }
    const crop = fill.crop ?? { l: 0, t: 0, r: 0, b: 0 };
    const visibleWidth = Math.max(1 - crop.l - crop.r, 0.01);
    const visibleHeight = Math.max(1 - crop.t - crop.b, 0.01);
    const clipId = `radv-clip-${hashString(`${context.scope}|${nodeId}`)}`;
    return {
      defs: [
        <image
          key="image-fill"
          href={url}
          x={(-crop.l / visibleWidth) * box.width}
          y={(-crop.t / visibleHeight) * box.height}
          width={box.width / visibleWidth}
          height={box.height / visibleHeight}
          preserveAspectRatio="none"
          clipPath={`url(#${clipId})`}
        />,
      ],
      fillValue: "none",
      fillOpacity: 1,
      usesImage: true,
    };
  }

  const paint = paintForFill(fill);
  return { defs: [], fillValue: paint.paint, fillOpacity: paint.opacity, usesImage: false };
}

function renderGeometryLayer(
  node: GeometryLayerInput,
  context: ShapeRenderContext,
  options: { strokeOnly?: boolean } = {},
): ReactNode {
  const width = emuToPx(node.xfrm.cx);
  const height = emuToPx(node.xfrm.cy);
  const path = geometryPath(node, context);
  if (!path) return null;

  const fill = options.strokeOnly ? null : node.fill ?? null;
  const paint = resolvePaint(fill, node.id, context, { width, height });
  const strokePaint = svgPaint(null, node.line);
  const clipId = paint.usesImage ? `radv-clip-${hashString(`${context.scope}|${node.id}`)}` : null;

  return (
    <svg
      width="100%"
      height="100%"
      viewBox={`0 0 ${width} ${height}`}
      preserveAspectRatio="none"
      style={{ position: "absolute", inset: 0, overflow: "visible", pointerEvents: "none" }}
      aria-hidden="true"
    >
      <defs>
        {paint.defs}
        {clipId ? (
          <clipPath id={clipId} clipPathUnits="userSpaceOnUse">
            <path d={path.d} fillRule={path.fillRule} />
          </clipPath>
        ) : null}
      </defs>
      <path
        d={path.d}
        fill={path.closed ? paint.fillValue : "none"}
        fillOpacity={paint.fillOpacity}
        fillRule={path.fillRule}
        stroke={strokePaint.stroke === "none" ? undefined : strokePaint.stroke}
        strokeWidth={strokePaint.strokeWidth || undefined}
        strokeOpacity={strokePaint.strokeOpacity}
        strokeDasharray={strokePaint.strokeDasharray}
        strokeLinecap={strokePaint.strokeLinecap}
        strokeLinejoin={strokePaint.strokeLinejoin}
      />
    </svg>
  );
}

function renderPicture(node: PictureNode, context: ShapeRenderContext): ReactNode {
  const width = emuToPx(node.xfrm.cx);
  const height = emuToPx(node.xfrm.cy);
  const url = node.mediaPath && !node.unrenderableMedia ? context.media.url(node.mediaPath) : null;
  const outline =
    node.line && node.line.fill && node.line.fill.type !== "none"
      ? renderGeometryLayer(
          { id: node.id, xfrm: node.xfrm, geometry: node.geometry, fill: null, line: node.line },
          context,
          { strokeOnly: true },
        )
      : null;

  if (!url) {
    const label = node.unrenderableMedia
      ? `${node.unrenderableMedia} image\ncannot be displayed in a browser`
      : "Image unavailable";
    return (
      <>
        <div style={PLACEHOLDER_STYLE} role="note" data-placeholder="image">
          {label}
        </div>
        {outline}
      </>
    );
  }

  return (
    <>
      <PictureImage node={node} url={url} width={width} height={height} context={context} />
      {outline}
    </>
  );
}

/**
 * The image itself.
 *
 * A part whose bytes cannot be decoded (a truncated PNG, a lying content type)
 * would otherwise leave the browser's broken-image icon in the slide with no
 * explanation, so a decode failure is caught and reported like any other
 * unrenderable element.
 */
function PictureImage({
  node,
  url,
  width,
  height,
  context,
}: {
  node: PictureNode;
  url: string;
  width: number;
  height: number;
  context: ShapeRenderContext;
}): ReactNode {
  const [failed, setFailed] = useState(false);
  const crop = node.crop;
  const visibleWidth = Math.max(1 - crop.l - crop.r, 0.01);
  const visibleHeight = Math.max(1 - crop.t - crop.b, 0.01);

  if (failed) {
    return (
      <div style={PLACEHOLDER_STYLE} role="note" data-placeholder="image-decode">
        {"Image cannot be displayed\n(decoding failed)"}
      </div>
    );
  }

  return (
    <div style={{ position: "absolute", inset: 0, overflow: "hidden" }}>
      <img
        src={url}
        alt={node.alt ?? ""}
        draggable={false}
        onError={() => {
          setFailed(true);
          context.warn(
            "picture.decode",
            `Image "${node.mediaPath ?? "?"}" could not be decoded and was replaced with a placeholder`,
          );
        }}
        style={{
          position: "absolute",
          left: `${(-crop.l / visibleWidth) * width}px`,
          top: `${(-crop.t / visibleHeight) * height}px`,
          width: `${width / visibleWidth}px`,
          height: `${height / visibleHeight}px`,
          maxWidth: "none",
          pointerEvents: "none",
        }}
      />
    </div>
  );
}

function renderTextLayer(
  node: ShapeNode | PictureNode,
  context: ShapeRenderContext,
  buildLink: LinkBuilder,
): ReactNode {
  if (!node.text) return null;
  const result = renderTextBody(node.text, buildLink, { interactive: context.interactive });
  return (
    <div style={{ position: "absolute", inset: 0, pointerEvents: context.interactive ? "auto" : "none" }}>
      {result.node}
    </div>
  );
}

function renderUnsupported(node: UnsupportedNode): ReactNode {
  return (
    <div style={PLACEHOLDER_STYLE} role="note" data-unsupported={node.reason}>
      {node.label}
    </div>
  );
}

function renderGroup(node: GroupNode, context: ShapeRenderContext, buildLink: LinkBuilder): ReactNode {
  const width = emuToPx(node.xfrm.cx);
  const height = emuToPx(node.xfrm.cy);
  const innerWidth = emuToPx(node.chExt.cx);
  const innerHeight = emuToPx(node.chExt.cy);
  const scaleX = innerWidth > 0 ? width / innerWidth : 1;
  const scaleY = innerHeight > 0 ? height / innerHeight : 1;
  return (
    <div style={{ position: "absolute", inset: 0, overflow: "visible" }}>
      <div
        style={{
          position: "absolute",
          left: 0,
          top: 0,
          width: `${innerWidth}px`,
          height: `${innerHeight}px`,
          // Child coordinates are in the group's child space; map them into the
          // group's own box before the parent transform is applied.
          transform: `translate(${-emuToPx(node.chOff.x)}px, ${-emuToPx(node.chOff.y)}px) scale(${scaleX}, ${scaleY})`,
          transformOrigin: "0 0",
        }}
      >
        {node.children.map((child, index) => (
          <ShapeNodeView key={`${child.id}-${index}`} node={child} context={context} buildLink={buildLink} />
        ))}
      </div>
    </div>
  );
}

export interface ShapeNodeViewProps {
  node: ShapeNodeUnion;
  context: ShapeRenderContext;
  buildLink: LinkBuilder;
}

export function ShapeNodeView({ node, context, buildLink }: ShapeNodeViewProps): ReactElement | null {
  if (node.hidden) return null;

  let content: ReactNode = null;
  let shadow: Shadow | null | undefined = null;

  switch (node.kind) {
    case "shape":
      shadow = node.shadow;
      content = (
        <>
          {renderGeometryLayer(node, context)}
          {renderTextLayer(node, context, buildLink)}
        </>
      );
      break;
    case "connector":
      content = renderGeometryLayer(node as ConnectorNode, context, { strokeOnly: true });
      break;
    case "picture":
      shadow = node.shadow;
      content = renderPicture(node, context);
      break;
    case "table":
      content = renderTable(node as TableNode, { interactive: context.interactive, buildLink });
      break;
    case "chart":
      content = renderChart(node as ChartNode, { palette: context.palette });
      break;
    case "unsupported":
      content = renderUnsupported(node);
      break;
    case "group":
      content = renderGroup(node as GroupNode, context, buildLink);
      break;
    default:
      return null;
  }

  if (content === null) return null;

  const style = wrapperStyle(node, shadow);
  const label = node.description ?? node.title;

  if (node.link && node.link.safe && context.interactive) {
    return (
      <a
        href={node.link.href}
        title={node.link.tooltip}
        target="_blank"
        rel="noopener noreferrer"
        style={{ ...style, display: "block" }}
        aria-label={label}
      >
        {content}
      </a>
    );
  }

  return (
    <div
      style={style}
      data-shape-kind={node.kind}
      data-shape-name={node.name ?? undefined}
      aria-label={label}
      role={label ? "img" : undefined}
    >
      {content}
    </div>
  );
}
