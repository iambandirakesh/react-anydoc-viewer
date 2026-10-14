/**
 * Shape tree parsing.
 *
 * `p:spTree` children are walked in document order (which *is* z-order) and each
 * one becomes a renderable node. Anything that cannot be rendered becomes an
 * `unsupported` node carrying a reason, so the renderer can show a labelled
 * placeholder instead of an empty area.
 */

import type {
  ChartData,
  ChartNode,
  ConnectorNode,
  Fill,
  GroupNode,
  LineStyle,
  PictureNode,
  Placeholder,
  RGBA,
  Shadow,
  ShapeNode,
  ShapeNodeUnion,
  TableNode,
  Transform,
  UnsupportedNode,
} from "../model/types";
import { attr, attrBool, attrNum, child, relAttr, type XmlNode } from "./xml";
import { isBrowserRenderable, type ImageKind } from "./archive";
import { resolveColorElement } from "./color";
import {
  parseFillFromProps,
  parseGeometry,
  parseGroupChildTransform,
  parseLine,
  parseShadow,
  parseTransform,
  type DrawingContext,
} from "./drawing";
import { parseTextBody, type TextBody, type TextContext } from "./text";
import { parseTable } from "./table";
import { sanitizeUrl } from "../utils/sanitize";

export interface PartResolver {
  /** drawing context bound to the part that owns the shapes */
  drawing(part: string): DrawingContext;
  /** text context (including inherited level styles) for a placeholder */
  text(part: string, ph: Placeholder | null): TextContext;
  /** geometry inherited from the matching layout/master placeholder */
  placeholderTransform(part: string, ph: Placeholder | null): Transform | null;
  /** parse a chart part; implementations cache by part name */
  chart(part: string): ChartData | null;
  /** image kind of a media part, or null when the part is missing */
  mediaKind(part: string): ImageKind | null;
  warn(code: string, message: string): void;
}

export type ShapeLayer = "slide" | "inherited";

export interface ShapeTreeOptions {
  layer: ShapeLayer;
}

const TABLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/table";
const CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const DIAGRAM_URI = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const OLE_URI = "http://schemas.openxmlformats.org/presentationml/2006/ole";

const DEFAULT_TRANSFORM: Transform = { x: 0, y: 0, cx: 0, cy: 0, rot: 0, flipH: false, flipV: false };

interface CommonProps {
  id: string;
  name: string | null;
  title?: string;
  description?: string;
  hidden: boolean;
  placeholder: Placeholder | null;
  linkId: string | null;
  linkTooltip: string | null;
}

function readCommonProps(container: XmlNode | null, fallbackId: string): CommonProps {
  const cNvPr = child(container, "cNvPr");
  const phNode = child(child(container, "nvPr"), "ph");
  const hlink = child(cNvPr, "hlinkClick");
  const placeholder: Placeholder | null = phNode
    ? {
        type: attr(phNode, "type") ?? "body",
        idx: attrNum(phNode, "idx"),
        orient: attr(phNode, "orient") ?? undefined,
        size: attr(phNode, "sz") ?? undefined,
      }
    : null;
  return {
    id: attr(cNvPr, "id") ?? fallbackId,
    name: attr(cNvPr, "name"),
    title: attr(cNvPr, "title") ?? undefined,
    description: attr(cNvPr, "descr") ?? undefined,
    hidden: attrBool(cNvPr, "hidden") === true,
    placeholder,
    linkId: relAttr(hlink, "id"),
    linkTooltip: attr(hlink, "tooltip"),
  };
}

function colorOf(container: XmlNode | null, ctx: DrawingContext): RGBA | null {
  if (!container) return null;
  for (const candidate of container.children) {
    switch (candidate.local) {
      case "srgbClr":
      case "schemeClr":
      case "sysClr":
      case "prstClr":
      case "scrgbClr":
      case "hslClr":
        return resolveColorElement(candidate, { theme: ctx.theme, colorMap: ctx.colorMap });
      default:
        break;
    }
  }
  return null;
}

/** Fill/line/effect/font references supplied by `<p:style>`. */
function resolveStyleRefs(
  style: XmlNode | null,
  ctx: DrawingContext,
  current: { fill: Fill | null; line: LineStyle | null; shadow: Shadow | null },
): {
  fill: Fill | null;
  line: LineStyle | null;
  shadow: Shadow | null;
  fontRef: { font?: string; color?: RGBA };
} {
  const fontRef: { font?: string; color?: RGBA } = {};
  if (!style) return { ...current, fontRef };

  const fillRef = child(style, "fillRef");
  const fillIdx = attrNum(fillRef, "idx") ?? 0;
  if (!current.fill && fillIdx > 0) {
    const list = fillIdx >= 1001 ? ctx.theme.bgFillStyles : ctx.theme.fillStyles;
    const template = list[fillIdx >= 1001 ? fillIdx - 1001 : fillIdx - 1];
    if (template) {
      const refColor = colorOf(fillRef, ctx);
      current.fill =
        refColor && template.type === "solid" ? { type: "solid", color: refColor } : template;
    }
  }

  const lnRef = child(style, "lnRef");
  const lnIdx = attrNum(lnRef, "idx") ?? 0;
  if (!current.line && lnIdx > 0) {
    const template = ctx.theme.lineStyles[lnIdx - 1];
    if (template) {
      const refColor = colorOf(lnRef, ctx);
      current.line =
        refColor && (!template.fill || template.fill.type === "solid")
          ? { ...template, fill: { type: "solid", color: refColor } }
          : template;
    }
  }

  const effectRef = child(style, "effectRef");
  const effectIdx = attrNum(effectRef, "idx") ?? 0;
  if (!current.shadow && effectIdx > 0) {
    const template = ctx.theme.effectStyles[effectIdx - 1];
    if (template) current.shadow = template;
  }

  const fontRefNode = child(style, "fontRef");
  if (fontRefNode) {
    const idx = attr(fontRefNode, "idx");
    if (idx === "major") fontRef.font = ctx.theme.majorFont ?? undefined;
    else if (idx === "minor") fontRef.font = ctx.theme.minorFont ?? undefined;
    const color = colorOf(fontRefNode, ctx);
    if (color) fontRef.color = color;
  }

  return { ...current, fontRef };
}

function hyperlinkFor(
  ctx: TextContext,
  linkId: string | null,
  tooltip: string | null,
  warn: PartResolver["warn"],
): ShapeNode["link"] {
  if (!linkId || !ctx.resolveHyperlink) return undefined;
  const target = ctx.resolveHyperlink(linkId);
  if (!target) {
    warn("hyperlink.missing-relationship", "Shape hyperlink references an unknown relationship");
    return undefined;
  }
  if (!target.external) {
    warn("hyperlink.internal", "Internal slide hyperlinks are not supported");
    return { href: "", safe: false, tooltip: tooltip ?? undefined };
  }
  const result = sanitizeUrl(target.target);
  if (!result.safe) {
    warn("hyperlink.blocked", `Blocked hyperlink target (${result.reason ?? "unsafe"})`);
    return { href: "", safe: false, tooltip: tooltip ?? undefined };
  }
  return { href: result.href, safe: true, tooltip: tooltip ?? undefined };
}

function fillIsVisible(fill: Fill | null): boolean {
  return fill !== null && fill.type !== "none";
}

function lineIsVisible(line: LineStyle | null): boolean {
  return (
    line !== null &&
    line.width !== null &&
    line.width > 0 &&
    fillIsVisible(line.fill)
  );
}

/* -------------------------------------------------------------- individual */

function parseShapeElement(
  sp: XmlNode,
  part: string,
  resolver: PartResolver,
  options: ShapeTreeOptions,
): ShapeNodeUnion | null {
  const ctx = resolver.drawing(part);
  const nvSpPr = child(sp, "nvSpPr");
  const common = readCommonProps(nvSpPr, attr(sp, "id") ?? "0");
  const spPr = child(sp, "spPr");
  const isTextBox = attrBool(child(nvSpPr, "cNvSpPr"), "txBox") === true;

  const ownTransform = parseTransform(spPr);
  const geometry = parseGeometry(spPr, ctx);
  const styleRefs = resolveStyleRefs(child(sp, "style"), ctx, {
    fill: parseFillFromProps(spPr, ctx),
    line: parseLine(spPr, ctx),
    shadow: parseShadow(child(spPr, "effectLst") ?? spPr, ctx),
  });

  const textCtx = resolver.text(part, common.placeholder);
  const text = parseTextBody(child(sp, "txBody"), {
    ...textCtx,
    defaultColor: styleRefs.fontRef.color ?? textCtx.defaultColor,
    defaultFont: styleRefs.fontRef.font ?? textCtx.defaultFont,
  });

  const visibleText = text !== null && text.hasText;
  const visible = fillIsVisible(styleRefs.fill) || lineIsVisible(styleRefs.line) || styleRefs.shadow !== null;

  if (!visibleText && !visible) {
    // Invisible (no fill, no outline, no text). Inherited layers are also full of
    // empty placeholder prompts, which must not be painted over the slide.
    return null;
  }
  if (options.layer === "inherited" && !visibleText && common.placeholder) {
    return null;
  }

  const xfrm =
    ownTransform ??
    resolver.placeholderTransform(part, common.placeholder) ??
    DEFAULT_TRANSFORM;

  const node: ShapeNode = {
    kind: "shape",
    id: common.id,
    name: isTextBox ? common.name ?? "Text Box" : common.name,
    xfrm,
    geometry,
    fill: styleRefs.fill,
    line: styleRefs.line,
    shadow: styleRefs.shadow,
    text,
    isPlaceholder: common.placeholder !== null,
  };
  if (common.hidden) node.hidden = true;
  if (common.title) node.title = common.title;
  if (common.description) node.description = common.description;
  if (common.placeholder) node.placeholder = common.placeholder;
  const link = hyperlinkFor(textCtx, common.linkId, common.linkTooltip, resolver.warn);
  if (link) node.link = link;
  return node;
}

function parsePictureElement(
  pic: XmlNode,
  part: string,
  resolver: PartResolver,
  options: ShapeTreeOptions,
): PictureNode | null {
  const ctx = resolver.drawing(part);
  const common = readCommonProps(child(pic, "nvPicPr"), attr(pic, "id") ?? "0");
  const blipFill = child(pic, "blipFill");
  const blip = child(blipFill, "blip");
  const spPr = child(pic, "spPr");
  const xfrm = parseTransform(spPr) ?? DEFAULT_TRANSFORM;

  let mediaPath: string | null = null;
  let mediaType: string | null = null;
  let unrenderableMedia: string | null = null;

  const embedId = relAttr(blip, "embed");
  const linkId = relAttr(blip, "link");
  const relId = embedId ?? linkId;
  if (relId && ctx.resolveRel) {
    const target = ctx.resolveRel(relId);
    if (!target) {
      resolver.warn("picture.missing-relationship", "Picture references an unknown relationship");
    } else if (target.external) {
      unrenderableMedia = "Linked image";
      resolver.warn(
        "picture.linked-external",
        "Linked (external) images are not fetched: the viewer never makes network requests for slide assets",
      );
    } else {
      mediaPath = target.target;
      const kind = resolver.mediaKind(target.target);
      mediaType = kind ? String(kind) : null;
      if (kind === null) {
        resolver.warn("picture.missing-part", `Image part "${target.target}" is missing from the package`);
      } else if (!isBrowserRenderable(kind)) {
        unrenderableMedia = kind.toUpperCase();
        resolver.warn(
          "picture.unsupported-format",
          `Embedded ${kind.toUpperCase()} images cannot be decoded by browsers`,
        );
      }
    }
  } else if (!relId) {
    resolver.warn("picture.no-image", "Picture element has no image reference");
  }

  const srcRect = child(blipFill, "srcRect");
  const pct = (name: string) => {
    const value = attrNum(srcRect, name);
    return value === null ? 0 : value / 100_000;
  };

  const textCtx = resolver.text(part, common.placeholder);
  const node: PictureNode = {
    kind: "picture",
    id: common.id,
    name: common.name,
    xfrm,
    mediaPath,
    mediaType,
    crop: { l: pct("l"), t: pct("t"), r: pct("r"), b: pct("b") },
    fill: parseFillFromProps(spPr, ctx),
    line: parseLine(spPr, ctx),
    geometry: parseGeometry(spPr, ctx),
    shadow: parseShadow(child(spPr, "effectLst") ?? spPr, ctx),
    text: mediaPath === null ? parseTextBody(child(pic, "txBody"), textCtx) : null,
    alt: common.description ?? common.title ?? null,
    unrenderableMedia,
  };
  if (common.hidden) node.hidden = true;
  if (common.title) node.title = common.title;
  if (common.description) node.description = common.description;
  if (common.placeholder) node.placeholder = common.placeholder;
  const link = hyperlinkFor(textCtx, common.linkId, common.linkTooltip, resolver.warn);
  if (link) node.link = link;
  if (options.layer === "inherited" && mediaPath === null && unrenderableMedia === null) return null;
  return node;
}

function unsupported(
  base: Pick<ShapeNode, "id" | "name" | "xfrm"> & { title?: string; description?: string },
  code: string,
  label: string,
  resolver: PartResolver,
): UnsupportedNode {
  resolver.warn(code, label);
  const node: UnsupportedNode = {
    kind: "unsupported",
    id: base.id,
    name: base.name,
    xfrm: base.xfrm,
    reason: code,
    label,
    text: null,
  };
  if (base.title) node.title = base.title;
  if (base.description) node.description = base.description;
  return node;
}

function parseGraphicFrame(
  frame: XmlNode,
  part: string,
  resolver: PartResolver,
  options: ShapeTreeOptions,
): ShapeNodeUnion | null {
  const ctx = resolver.drawing(part);
  const common = readCommonProps(child(frame, "nvGraphicFramePr"), attr(frame, "id") ?? "0");
  const xfrm = parseTransform(frame) ?? DEFAULT_TRANSFORM;
  const graphic = child(frame, "graphic");
  const data = graphic ? child(graphic, "graphicData") : null;
  const uri = attr(data, "uri") ?? "";
  void options;

  const base = {
    id: common.id,
    name: common.name,
    xfrm,
    title: common.title,
    description: common.description,
  };

  if (!data) {
    return unsupported(base, "graphicFrame.no-data", "Empty graphic frame", resolver);
  }

  const tbl = child(data, "tbl");
  if (uri === TABLE_URI || tbl) {
    const table = parseTable(tbl, {
      textContext: resolver.text(part, common.placeholder),
      onApproximatedStyle: () =>
        resolver.warn(
          "table.style-approximated",
          "Table styling was approximated (table styles from tableStyles.xml are not applied)",
        ),
    });
    if (!table) {
      return unsupported(base, "table.empty", "Table could not be read", resolver);
    }
    const node: TableNode = { kind: "table", ...base, table };
    if (common.placeholder) node.placeholder = common.placeholder;
    return node;
  }

  const chartRef = child(data, "chart");
  if (uri === CHART_URI || chartRef) {
    const relId = relAttr(chartRef, "id");
    if (!relId || !ctx.resolveRel) {
      return unsupported(base, "chart.no-relationship", "Chart is missing its relationship", resolver);
    }
    const resolved = ctx.resolveRel(relId);
    if (!resolved || resolved.external) {
      return unsupported(base, "chart.bad-relationship", "Chart part could not be resolved", resolver);
    }
    const chart = resolver.chart(resolved.target);
    if (!chart) {
      return unsupported(
        base,
        "chart.unreadable",
        `Chart part "${resolved.target}" could not be parsed`,
        resolver,
      );
    }
    const node: ChartNode = { kind: "chart", ...base, chart };
    if (common.placeholder) node.placeholder = common.placeholder;
    return node;
  }

  if (uri === DIAGRAM_URI || child(data, "relIds")) {
    return unsupported(base, "smartart.unsupported", "SmartArt diagram (not supported)", resolver);
  }

  if (uri === OLE_URI || child(data, "oleObj")) {
    return unsupported(base, "ole.unsupported", "Embedded OLE object", resolver);
  }

  return unsupported(
    base,
    "graphicFrame.unsupported",
    uri.length > 0 ? `Unsupported frame (${uri.split("/").pop() ?? uri})` : "Unsupported graphic frame",
    resolver,
  );
}

function parseGroupElement(
  grp: XmlNode,
  part: string,
  resolver: PartResolver,
  options: ShapeTreeOptions,
): GroupNode | null {
  const common = readCommonProps(child(grp, "nvGrpSpPr"), attr(grp, "id") ?? "0");
  const grpSpPr = child(grp, "grpSpPr");
  const xfrm = parseTransform(grpSpPr) ?? DEFAULT_TRANSFORM;
  const childTransform = parseGroupChildTransform(grpSpPr);
  const chOff = childTransform?.off ?? { x: 0, y: 0 };
  const ext = childTransform?.ext ?? { cx: xfrm.cx, cy: xfrm.cy };
  const chExt = { cx: ext.cx > 0 ? ext.cx : 1, cy: ext.cy > 0 ? ext.cy : 1 };

  const childNodes = parseChildren(grp, part, resolver, options);
  if (childNodes.length === 0) return null;

  const node: GroupNode = {
    kind: "group",
    id: common.id,
    name: common.name,
    xfrm,
    chOff,
    chExt,
    children: childNodes,
  };
  if (common.hidden) node.hidden = true;
  return node;
}

function parseConnectorElement(
  cxn: XmlNode,
  part: string,
  resolver: PartResolver,
  options: ShapeTreeOptions,
): ConnectorNode | null {
  void options;
  const ctx = resolver.drawing(part);
  const common = readCommonProps(child(cxn, "nvCxnSpPr"), attr(cxn, "id") ?? "0");
  const spPr = child(cxn, "spPr");
  const xfrm = parseTransform(spPr) ?? DEFAULT_TRANSFORM;
  const styleRefs = resolveStyleRefs(child(cxn, "style"), ctx, {
    fill: null,
    line: parseLine(spPr, ctx),
    shadow: null,
  });
  if (!lineIsVisible(styleRefs.line)) return null;
  const node: ConnectorNode = {
    kind: "connector",
    id: common.id,
    name: common.name,
    xfrm,
    geometry: parseGeometry(spPr, ctx),
    line: styleRefs.line,
  };
  if (common.hidden) node.hidden = true;
  return node;
}

/* ---------------------------------------------------------- tree walking */

function parseSingle(
  element: XmlNode,
  part: string,
  resolver: PartResolver,
  options: ShapeTreeOptions,
): ShapeNodeUnion[] {
  switch (element.local) {
    case "sp": {
      const node = parseShapeElement(element, part, resolver, options);
      return node ? [node] : [];
    }
    case "pic": {
      const node = parsePictureElement(element, part, resolver, options);
      return node ? [node] : [];
    }
    case "graphicFrame": {
      const node = parseGraphicFrame(element, part, resolver, options);
      return node ? [node] : [];
    }
    case "grpSp": {
      const node = parseGroupElement(element, part, resolver, options);
      return node ? [node] : [];
    }
    case "cxnSp": {
      const node = parseConnectorElement(element, part, resolver, options);
      return node ? [node] : [];
    }
    case "AlternateContent": {
      const fallback = child(element, "Fallback");
      const branch = fallback ?? child(element, "Choice");
      if (!fallback) {
        resolver.warn(
          "mc.choice-only",
          "Rendered a Markup Compatibility choice branch that has no fallback",
        );
      }
      if (!branch) return [];
      const out: ShapeNodeUnion[] = [];
      for (const inner of branch.children) {
        out.push(...parseSingle(inner, part, resolver, options));
      }
      return out;
    }
    default:
      return [];
  }
}

export function parseChildren(
  container: XmlNode,
  part: string,
  resolver: PartResolver,
  options: ShapeTreeOptions,
): ShapeNodeUnion[] {
  const out: ShapeNodeUnion[] = [];
  for (const element of container.children) {
    out.push(...parseSingle(element, part, resolver, options));
  }
  return out;
}

export function parseShapeTree(
  spTree: XmlNode | null,
  part: string,
  resolver: PartResolver,
  options: ShapeTreeOptions = { layer: "slide" },
): ShapeNodeUnion[] {
  if (!spTree) return [];
  return parseChildren(spTree, part, resolver, options);
}

/** Plain text of a shape, used by tests and by the thumbnail renderer. */
export function shapeText(node: ShapeNodeUnion): string {
  if (node.kind === "shape" || node.kind === "picture" || node.kind === "table") {
    if (node.kind === "shape" || node.kind === "picture") return textOf(node.text);
  }
  if (node.kind === "table") {
    return node.table.rows
      .map((row) => row.cells.map((cell) => textOf(cell.text)).join("\t"))
      .join("\n");
  }
  if (node.kind === "group") {
    return node.children.map(shapeText).join("\n");
  }
  return "";
}

function textOf(body: TextBody | null): string {
  if (!body) return "";
  return body.paragraphs
    .map((paragraph) => paragraph.runs.map((run) => (run.br ? "\n" : run.text)).join(""))
    .join("\n");
}
