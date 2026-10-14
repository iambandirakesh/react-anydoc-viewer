/**
 * Presentation orchestration.
 *
 * Turns an OPC package into a fully resolved {@link PresentationModel}:
 * slide order from `p:sldIdLst`, inheritance from layout -> master -> theme,
 * background resolution, and per-slide diagnostics.
 *
 * Everything reusable (themes, masters, layouts, charts, the archive itself) is
 * parsed at most once per presentation.
 */

import type {
  ColorMap,
  Diagnostic,
  Fill,
  LvlStyle,
  Placeholder,
  PresentationModel,
  RGBA,
  ShapeNodeUnion,
  SlideModel,
  TextStyles,
  ThemeData,
  Transform,
} from "../model/types";
import {
  createMediaStore,
  type MediaStore,
} from "./media";
import {
  contentTypeFor,
  parseContentTypes,
  readText,
  readArchive,
  type Archive,
  type ArchiveLimits,
  type ImageKind,
} from "./archive";
import {
  attr,
  child,
  children,
  parseXml,
  type XmlNode,
} from "./xml";
import { emuToPx } from "./units";
import {
  findAllRelationships,
  findOfficeDocumentPart,
  findRelationship,
  relationshipsFor,
} from "./relationships";
import {
  DEFAULT_COLOR_MAP,
  createEmptyTheme,
  parseColorMap,
  parseTheme,
  resolveBackgroundRef,
} from "./theme";
import {
  accentPalette,
  parseFillFromProps,
  parseTransform,
  type DrawingContext,
  type RelTarget,
} from "./drawing";
import { parseLevelStyles, type TextBody, type TextContext } from "./text";
import { parseShapeTree, type PartResolver } from "./shapes";
import { parseChart, type ChartParseContext } from "./chart";
import { resolveColorElement } from "./color";
import { Diagnostics, DEFAULT_MAX_DIAGNOSTICS } from "./diagnostics";
import { DocumentViewerError } from "../errors";
import type { ChartData } from "../model/types";

const EMU_PER_INCH = 914_400;
const DEFAULT_SLIDE_WIDTH = 9_144_000; // 10in, 4:3
const DEFAULT_SLIDE_HEIGHT = 6_858_000; // 7.5in, 4:3

export interface ParseOptions {
  limits?: Partial<ArchiveLimits>;
  /** maximum number of diagnostics retained (default 200) */
  maxDiagnostics?: number;
}

export interface ParsedPresentation {
  model: PresentationModel;
  media: MediaStore;
  dispose(): void;
}

/* ------------------------------------------------------- cached sub-parts */

interface PlaceholderInfo {
  transform: Transform | null;
  levels: (LvlStyle | null)[];
}

interface MasterInfo {
  part: string;
  colorMap: ColorMap;
  theme: ThemeData;
  textStyles: TextStyles;
  shapes: ShapeNodeUnion[];
  background: Fill | null;
  placeholders: Map<string, PlaceholderInfo>;
  layoutParts: string[];
}

interface LayoutInfo {
  part: string;
  masterPart: string | null;
  showMasterShapes: boolean;
  shapes: ShapeNodeUnion[];
  background: Fill | null;
  colorMap: ColorMap | null;
  placeholders: Map<string, PlaceholderInfo>;
}

function placeholderKey(type: string, idx: number | null): string {
  return `${type}|${idx ?? ""}`;
}

/**
 * Placeholders are matched by type *and* index first (the specification's
 * identity), then by type alone, which is what producers usually emit.
 */
function lookupPlaceholder(
  map: Map<string, PlaceholderInfo>,
  ph: Placeholder | null,
): PlaceholderInfo | null {
  if (!ph || map.size === 0) return null;
  const exact = map.get(placeholderKey(ph.type, ph.idx));
  if (exact) return exact;
  const typed = map.get(placeholderKey(ph.type, null));
  if (typed) return typed;
  const prefix = `${ph.type}|`;
  for (const [key, value] of map) {
    if (key.startsWith(prefix)) return value;
  }
  return null;
}

function collectPlaceholders(
  spTree: XmlNode | null,
  ctx: TextContext,
): Map<string, PlaceholderInfo> {
  const map = new Map<string, PlaceholderInfo>();
  if (!spTree) return map;
  const stack: XmlNode[] = [...spTree.children];
  while (stack.length > 0) {
    const node = stack.shift() as XmlNode;
    if (node.local === "sp") {
      const nvSpPr = child(node, "nvSpPr");
      const phNode = child(child(nvSpPr, "nvPr"), "ph");
      if (phNode) {
        const ph: Placeholder = {
          type: attr(phNode, "type") ?? "body",
          idx: null,
        };
        const idxRaw = attr(phNode, "idx");
        const idx = idxRaw === null ? null : Number(idxRaw);
        ph.idx = idx === null || Number.isNaN(idx) ? null : idx;
        const spPr = child(node, "spPr");
        const levels = parseLevelStyles(child(child(node, "txBody"), "lstStyle"), ctx);
        map.set(placeholderKey(ph.type, ph.idx), {
          transform: parseTransform(spPr),
          levels,
        });
      }
    }
    for (const inner of node.children) {
      if (inner.local === "grpSp" || inner.local === "spTree") stack.push(inner);
    }
  }
  return map;
}

function mergeLevels(
  base: (LvlStyle | null)[] | undefined,
  override: (LvlStyle | null)[] | undefined,
): (LvlStyle | null)[] {
  const out: (LvlStyle | null)[] = [];
  for (let level = 0; level < 9; level += 1) {
    const a = base?.[level] ?? null;
    const b = override?.[level] ?? null;
    if (!a && !b) {
      out.push(null);
      continue;
    }
    out.push({
      defRun: { ...(a?.defRun ?? {}), ...clean(b?.defRun) },
      pPr: { ...(a?.pPr ?? {}), ...clean(b?.pPr) },
    });
  }
  return out;
}

function clean<T extends object | undefined>(value: T): Partial<NonNullable<T>> {
  if (!value) return {};
  const out: Record<string, unknown> = {};
  for (const [key, entry] of Object.entries(value)) {
    if (entry !== undefined) out[key] = entry;
  }
  return out as Partial<NonNullable<T>>;
}

/* ------------------------------------------------------------- the parser */

class PresentationParser {
  private readonly diagnostics: Diagnostics;
  private readonly themeCache = new Map<string, ThemeData>();
  private readonly masterCache = new Map<string, MasterInfo>();
  private readonly layoutCache = new Map<string, LayoutInfo>();
  private readonly chartCache = new Map<string, ChartData | null>();
  private readonly media: MediaStore;
  private readonly contentTypes;

  constructor(
    private readonly archive: Archive,
    options: ParseOptions,
  ) {
    this.diagnostics = new Diagnostics(options.maxDiagnostics ?? DEFAULT_MAX_DIAGNOSTICS);
    this.media = createMediaStore(archive);
    let types = { defaults: new Map<string, string>(), overrides: new Map<string, string>() };
    try {
      types = parseContentTypes(parseXml(readText(archive, "[Content_Types].xml") ?? "", {
        part: "[Content_Types].xml",
      }));
    } catch {
      this.warn("content-types.unreadable", "The content-type map could not be read", "[Content_Types].xml");
    }
    this.contentTypes = types;
  }

  private warn(code: string, message: string, part?: string, slideIndex?: number): void {
    const diagnostic: Diagnostic = { severity: "warning", code, message };
    if (part !== undefined) diagnostic.part = part;
    if (slideIndex !== undefined) diagnostic.slideIndex = slideIndex;
    this.diagnostics.add(diagnostic);
  }

  private info(code: string, message: string, part?: string, slideIndex?: number): void {
    const diagnostic: Diagnostic = { severity: "info", code, message };
    if (part !== undefined) diagnostic.part = part;
    if (slideIndex !== undefined) diagnostic.slideIndex = slideIndex;
    this.diagnostics.add(diagnostic);
  }

  /* ------------------------------------------------------------ contexts */

  private drawingFor(part: string, colorMap: ColorMap, theme: ThemeData): DrawingContext {
    return {
      theme,
      colorMap,
      resolveRel: (id: string): RelTarget | null => {
        const rel = relationshipsFor(this.archive, part).get(id);
        if (!rel) return null;
        return { target: rel.target, external: rel.external };
      },
      warn: (code, message) => this.warn(code, message, part),
    };
  }

  private themeFor(part: string): ThemeData {
    const cached = this.themeCache.get(part);
    if (cached) return cached;
    let theme = createEmptyTheme();
    try {
      theme = parseTheme(parseXml(readText(this.archive, part) ?? "", { part }));
    } catch {
      this.warn("theme.unreadable", "The theme part could not be parsed; default Office colours are used", part);
    }
    this.themeCache.set(part, theme);
    return theme;
  }

  private levelStylesFor(obj: XmlNode | null): (LvlStyle | null)[] {
    return parseLevelStyles(obj, this.emptyTextContext());
  }

  private emptyTextContext(): TextContext {
    return {
      theme: createEmptyTheme(),
      colorMap: DEFAULT_COLOR_MAP,
      warn: () => undefined,
    };
  }

  /* ------------------------------------------------------ masters/layouts */

  private loadMaster(masterPart: string): MasterInfo {
    const cached = this.masterCache.get(masterPart);
    if (cached) return cached;

    const root = this.safeParse(masterPart);
    const colorMap = parseColorMap(child(root, "clrMap"));

    const themeRel = findRelationship(this.archive, masterPart, "theme");
    const theme = this.themeFor(themeRel ? themeRel.target : "ppt/theme/theme1.xml");

    const txStyles = child(root, "txStyles");
    const textStyles: TextStyles = {
      title: this.levelStylesFor(child(txStyles, "titleStyle")),
      body: this.levelStylesFor(child(txStyles, "bodyStyle")),
      other: this.levelStylesFor(child(txStyles, "otherStyle")),
    };

    const cSld = child(root, "cSld");
    const spTree = child(cSld, "spTree");
    const placeholders = collectPlaceholders(spTree, { theme, colorMap, warn: () => undefined });
    const resolver = this.resolver({
      part: masterPart,
      colorMap,
      theme,
      masterTextStyles: textStyles,
      layoutPlaceholders: null,
      masterPlaceholders: placeholders,
    });
    const shapes = parseShapeTree(spTree, masterPart, resolver, { layer: "inherited" });
    const background = this.parseBackground(child(cSld, "bg"), theme, colorMap, masterPart);

    const info: MasterInfo = {
      part: masterPart,
      colorMap,
      theme,
      textStyles,
      shapes,
      background,
      placeholders,
      layoutParts: findAllRelationships(this.archive, masterPart, "slideLayout").map(
        (rel) => rel.target,
      ),
    };
    this.masterCache.set(masterPart, info);
    return info;
  }

  private loadLayout(layoutPart: string): LayoutInfo {
    const cached = this.layoutCache.get(layoutPart);
    if (cached) return cached;

    const root = this.safeParse(layoutPart);
    const masterRel = findRelationship(this.archive, layoutPart, "slideMaster");
    const masterPart = masterRel ? masterRel.target : null;
    const master = masterPart ? this.loadMaster(masterPart) : null;
    const theme = master ? master.theme : createEmptyTheme();
    const colorMap = this.layoutColorMap(root, master?.colorMap ?? DEFAULT_COLOR_MAP);

    const cSld = child(root, "cSld");
    const spTree = child(cSld, "spTree");
    const placeholders = collectPlaceholders(spTree, { theme, colorMap, warn: () => undefined });
    const resolver = this.resolver({
      part: layoutPart,
      colorMap,
      theme,
      masterTextStyles: master?.textStyles ?? emptyTextStyles(),
      layoutPlaceholders: placeholders,
      masterPlaceholders: master?.placeholders ?? null,
    });
    const shapes = parseShapeTree(spTree, layoutPart, resolver, { layer: "inherited" });
    const background = this.parseBackground(child(cSld, "bg"), theme, colorMap, layoutPart);

    const info: LayoutInfo = {
      part: layoutPart,
      masterPart,
      showMasterShapes: attr(root, "showMasterSp") !== "0",
      shapes,
      background,
      colorMap: this.hasColorMapOverride(root) ? colorMap : null,
      placeholders,
    };
    this.layoutCache.set(layoutPart, info);
    return info;
  }

  private hasColorMapOverride(root: XmlNode | null): boolean {
    return child(child(root, "clrMapOvr"), "overrideClrMapping") !== null;
  }

  private layoutColorMap(root: XmlNode | null, fallback: ColorMap): ColorMap {
    const override = child(child(root, "clrMapOvr"), "overrideClrMapping");
    if (override) return parseColorMap(override);
    return fallback;
  }

  /* --------------------------------------------------------------- assets */

  private resolver(options: {
    part: string;
    colorMap: ColorMap;
    theme: ThemeData;
    masterTextStyles: TextStyles;
    layoutPlaceholders: Map<string, PlaceholderInfo> | null;
    masterPlaceholders: Map<string, PlaceholderInfo> | null;
  }): PartResolver {
    const { part, colorMap, theme } = options;
    return {
      drawing: (p) => this.drawingFor(p, colorMap, theme),
      text: (p, ph) =>
        this.textContextFor({
          part: p,
          ph,
          colorMap,
          theme,
          masterTextStyles: options.masterTextStyles,
          layoutPlaceholders: options.layoutPlaceholders,
          masterPlaceholders: options.masterPlaceholders,
        }),
      placeholderTransform: (_p, ph) => {
        const fromLayout = lookupPlaceholder(options.layoutPlaceholders ?? new Map(), ph);
        if (fromLayout?.transform) return fromLayout.transform;
        const fromMaster = lookupPlaceholder(options.masterPlaceholders ?? new Map(), ph);
        return fromMaster?.transform ?? null;
      },
      chart: (p) => this.chartFor(p, colorMap, theme),
      mediaKind: (p) => this.media.kind(p),
      warn: (code, message) => this.warn(code, message, part),
    };
  }

  private chartFor(part: string, colorMap: ColorMap, theme: ThemeData): ChartData | null {
    const key = part.toLowerCase();
    if (this.chartCache.has(key)) return this.chartCache.get(key) ?? null;
    let chart: ChartData | null = null;
    try {
      const root = parseXml(readText(this.archive, part) ?? "", { part });
      const ctx: ChartParseContext = {
        theme,
        colorMap,
        resolveRel: (id) => {
          const rel = relationshipsFor(this.archive, part).get(id);
          return rel ? { target: rel.target, external: rel.external } : null;
        },
        warn: (code, message) => this.warn(code, message, part),
        textContext: this.emptyTextContext(),
        palette: accentPalette(theme, colorMap),
      };
      chart = parseChart(root, ctx);
      if (!chart) this.warn("chart.empty", `Chart part "${part}" contains no readable plot`, part);
    } catch {
      this.warn("chart.unreadable", `Chart part "${part}" could not be parsed`, part);
      chart = null;
    }
    this.chartCache.set(key, chart);
    return chart;
  }

  private textContextFor(options: {
    part: string;
    ph: Placeholder | null;
    colorMap: ColorMap;
    theme: ThemeData;
    masterTextStyles: TextStyles;
    layoutPlaceholders: Map<string, PlaceholderInfo> | null;
    masterPlaceholders: Map<string, PlaceholderInfo> | null;
  }): TextContext {
    const { part, ph, colorMap, theme } = options;
    const bucket = textStyleBucket(ph);
    // Cascade: master text styles, then master placeholder, then layout placeholder.
    let levels: (LvlStyle | null)[] | undefined = options.masterTextStyles[bucket];
    const masterPlaceholder = lookupPlaceholder(options.masterPlaceholders ?? new Map(), ph);
    if (masterPlaceholder) levels = mergeLevels(levels, masterPlaceholder.levels);
    const layoutPlaceholder = lookupPlaceholder(options.layoutPlaceholders ?? new Map(), ph);
    if (layoutPlaceholder) levels = mergeLevels(levels, layoutPlaceholder.levels);

    const isTitle = bucket === "title";
    const defaultColor =
      theme.colors[colorMap.tx1 ?? "tx1"] ?? theme.colors.dk1 ?? { r: 0, g: 0, b: 0, a: 1 };

    return {
      theme,
      colorMap,
      resolveRel: (id) => {
        const rel = relationshipsFor(this.archive, part).get(id);
        return rel ? { target: rel.target, external: rel.external } : null;
      },
      resolveHyperlink: (id) => {
        const rel = relationshipsFor(this.archive, part).get(id);
        return rel ? { target: rel.target, external: rel.external } : null;
      },
      warn: (code, message) => this.warn(code, message, part),
      inherited: levels,
      defaultColor,
      defaultFont: (isTitle ? theme.majorFont : theme.minorFont) ?? undefined,
    };
  }

  /* ----------------------------------------------------------- background */

  private parseBackground(
    bg: XmlNode | null,
    theme: ThemeData,
    colorMap: ColorMap,
    part: string,
  ): Fill | null {
    if (!bg) return null;
    const ctx = this.drawingFor(part, colorMap, theme);
    const bgPr = child(bg, "bgPr");
    if (bgPr) {
      const fill = parseFillFromProps(bgPr, ctx);
      return fill;
    }
    const bgRef = child(bg, "bgRef");
    if (bgRef) {
      const idx = Number(attr(bgRef, "idx") ?? "0");
      if (!Number.isFinite(idx) || idx <= 0) return null;
      const fill = resolveBackgroundRef(idx, theme, { theme });
      const colorNode = bgRef.children.find((c) =>
        ["srgbClr", "schemeClr", "sysClr", "prstClr", "scrgbClr", "hslClr"].includes(c.local),
      );
      const color: RGBA | null = colorNode
        ? resolveColorElement(colorNode, { theme, colorMap })
        : null;
      if (fill) {
        if (fill.type === "solid" && color) return { type: "solid", color };
        return fill;
      }
      if (color) return { type: "solid", color };
    }
    return null;
  }

  /* ---------------------------------------------------------------- slide */

  private safeParse(part: string): XmlNode | null {
    try {
      return parseXml(readText(this.archive, part) ?? "", { part });
    } catch (error) {
      this.warn(
        "part.unreadable",
        `Part "${part}" could not be parsed: ${error instanceof Error ? error.message : String(error)}`,
        part,
      );
      return null;
    }
  }

  parse(): ParsedPresentation {
    const officeDocument = findOfficeDocumentPart(this.archive);
    if (officeDocument.fromFallback) {
      this.warn(
        "package.root-rels-missing",
        "The package root relationships were missing; the conventional presentation path was used instead",
      );
    }
    const presentationPart = officeDocument.part;
    const presentation = this.safeParse(presentationPart);
    if (!presentation) {
      throw new DocumentViewerError(
        "MALFORMED_PRESENTATION",
        "The presentation part could not be read from the package.",
        { detail: presentationPart },
      );
    }

    const sldSz = child(presentation, "sldSz");
    const width = Number(attr(sldSz, "cx") ?? DEFAULT_SLIDE_WIDTH) || DEFAULT_SLIDE_WIDTH;
    const height = Number(attr(sldSz, "cy") ?? DEFAULT_SLIDE_HEIGHT) || DEFAULT_SLIDE_HEIGHT;

    // Warm the layout cache before slides so placeholder inheritance resolves.
    for (const masterRel of findAllRelationships(this.archive, presentationPart, "slideMaster")) {
      const master = this.loadMaster(masterRel.target);
      for (const layoutPart of master.layoutParts) this.loadLayout(layoutPart);
    }

    const slideParts: { part: string; hidden: boolean }[] = [];
    const sldIdLst = child(presentation, "sldIdLst");
    for (const sldId of children(sldIdLst, "sldId")) {
      const relId = attr(sldId, "r:id");
      if (!relId) continue;
      const rel = relationshipsFor(this.archive, presentationPart).get(relId);
      if (!rel || rel.external) {
        this.warn(
          "slide.missing-relationship",
          `Slide entry "${relId}" does not resolve to a slide part`,
          presentationPart,
        );
        continue;
      }
      slideParts.push({ part: rel.target, hidden: attr(sldId, "show") === "0" });
    }

    if (slideParts.length === 0) {
      throw new DocumentViewerError(
        "NO_SLIDES",
        "This presentation contains no slides that could be resolved.",
      );
    }

    const slides: SlideModel[] = [];
    for (let index = 0; index < slideParts.length; index += 1) {
      const entry = slideParts[index]!;
      slides.push(this.parseSlide(entry.part, index, entry.hidden));
    }

    const firstMaster = this.masterForSlide(slideParts[0]!.part);
    const theme = firstMaster?.theme ?? createEmptyTheme();

    const model: PresentationModel = {
      width,
      height,
      slides,
      theme,
      colorMap: firstMaster?.colorMap ?? DEFAULT_COLOR_MAP,
      textStyles: firstMaster?.textStyles ?? emptyTextStyles(),
      title: this.coreProperty("title"),
      author: this.coreProperty("creator"),
      partCount: this.archive.entries.size,
      diagnostics: [],
    };
    model.diagnostics = this.diagnostics.list();
    return {
      model,
      media: this.media,
      dispose: () => this.media.dispose(),
    };
  }

  private layoutForSlide(slidePart: string): LayoutInfo | null {
    const rel = findRelationship(this.archive, slidePart, "slideLayout");
    return rel ? this.layoutCache.get(rel.target) ?? null : null;
  }

  private masterForSlide(slidePart: string): MasterInfo | null {
    const layout = this.layoutForSlide(slidePart);
    if (!layout?.masterPart) return null;
    return this.masterCache.get(layout.masterPart) ?? null;
  }

  private coreProperty(name: string): string | null {
    try {
      const root = parseXml(readText(this.archive, "docProps/core.xml") ?? "", {
        part: "docProps/core.xml",
      });
      if (!root) return null;
      const node = child(root, name);
      return node && node.text.trim().length > 0 ? node.text.trim() : null;
    } catch {
      return null;
    }
  }

  private parseSlide(slidePart: string, index: number, hidden: boolean): SlideModel {
    const mark = this.diagnostics.mark();
    const root = this.safeParse(slidePart);
    const cSld = child(root, "cSld");
    const spTree = child(cSld, "spTree");

    const layoutRel = findRelationship(this.archive, slidePart, "slideLayout");
    const layout = layoutRel ? this.loadLayout(layoutRel.target) : null;
    if (!layout) {
      this.warn("slide.no-layout", "The slide has no layout relationship", slidePart, index);
    }
    const master = layout?.masterPart ? this.masterCache.get(layout.masterPart) : undefined;
    const theme = master?.theme ?? createEmptyTheme();

    const slideColorMap: ColorMap = (() => {
      const ovr = child(root, "clrMapOvr");
      const override = child(ovr, "overrideClrMapping");
      if (override) return parseColorMap(override);
      if (layout?.colorMap) return layout.colorMap;
      return master?.colorMap ?? DEFAULT_COLOR_MAP;
    })();

    const resolver = this.resolver({
      part: slidePart,
      colorMap: slideColorMap,
      theme,
      masterTextStyles: master?.textStyles ?? emptyTextStyles(),
      layoutPlaceholders: layout?.placeholders ?? null,
      masterPlaceholders: master?.placeholders ?? null,
    });
    const shapes = parseShapeTree(spTree, slidePart, resolver, { layer: "slide" });
    if (shapes.length === 0) {
      this.info("slide.empty", "The slide contains no renderable shapes", slidePart, index);
    }

    const inheritedShapes: ShapeNodeUnion[] = [];
    if (layout) {
      if (layout.showMasterShapes && master) inheritedShapes.push(...master.shapes);
      inheritedShapes.push(...layout.shapes);
    }

    const background =
      this.parseBackground(child(cSld, "bg"), theme, slideColorMap, slidePart) ??
      layout?.background ??
      master?.background ??
      null;

    const slideDiagnostics = this.diagnostics.addedSince(mark);

    return {
      index,
      partName: slidePart,
      name: attr(cSld, "name"),
      slideshowHidden: hidden,
      shapes,
      inheritedShapes,
      background,
      diagnostics: slideDiagnostics,
    };
  }
}

function emptyTextStyles(): TextStyles {
  return { title: [], body: [], other: [] };
}

function textStyleBucket(ph: Placeholder | null): "title" | "body" | "other" {
  if (!ph) return "other";
  switch (ph.type) {
    case "title":
    case "ctrTitle":
      return "title";
    case "body":
    case "subTitle":
    case "obj":
      return "body";
    default:
      return "other";
  }
}

/* ------------------------------------------------------------------ public */

export function parsePresentationArchive(
  archive: Archive,
  options: ParseOptions = {},
): ParsedPresentation {
  const parser = new PresentationParser(archive, options);
  return parser.parse();
}

export function parsePresentationBytes(
  bytes: Uint8Array,
  options: ParseOptions = {},
): ParsedPresentation {
  const archive = readArchive(bytes, options.limits ?? {});
  return parsePresentationArchive(archive, options);
}

/* --------------------------------------------------------------- utilities */

export function slideSizeInches(model: PresentationModel): { width: number; height: number } {
  return { width: model.width / EMU_PER_INCH, height: model.height / EMU_PER_INCH };
}

export function slideSizePx(model: PresentationModel, pixelsPerInch = 96): { width: number; height: number } {
  const inches = slideSizeInches(model);
  return {
    width: Math.round(inches.width * pixelsPerInch),
    height: Math.round(inches.height * pixelsPerInch),
  };
}

export function hasSlideText(model: PresentationModel, index: number): boolean {
  const slide = model.slides[index];
  if (!slide) return false;
  for (const shape of slide.shapes) {
    if (shape.kind === "shape" && shape.text?.hasText) return true;
    if (shape.kind === "group") {
      const stack = [...shape.children];
      while (stack.length > 0) {
        const node = stack.shift();
        if (!node) break;
        if (node.kind === "shape" && node.text?.hasText) return true;
        if (node.kind === "group") stack.push(...node.children);
      }
    }
  }
  return false;
}

export { emuToPx, contentTypeFor };
export type { TextBody };
