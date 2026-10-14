/**
 * Internal presentation model.
 *
 * This is the boundary between the parser layer (`src/parser/*`) and the renderer
 * layer (`src/renderer/*`, `src/components/*`). Everything in here is already
 * resolved: theme colours are resolved to RGBA, EMUs are kept as EMUs, and
 * relationship ids have been turned into concrete part names.
 *
 * It is intentionally NOT exported from the package entry point: the public API
 * only exposes opaque handles and summary metadata so that internals can change
 * without a breaking release. `src/types/public.ts` documents the public surface.
 */

export type Emu = number;

/** A 2-D transform in slide coordinates (EMU). `rot` is in degrees. */
export interface Transform {
  x: Emu;
  y: Emu;
  cx: Emu;
  cy: Emu;
  rot: number;
  flipH: boolean;
  flipV: boolean;
}

export interface RGBA {
  r: number;
  g: number;
  b: number;
  a: number;
}

export interface GradientStop {
  pos: number;
  color: RGBA;
}

/** `l/t/r/b` are fractions of the source image that are cropped away (0..1). */
export interface Crop {
  l: number;
  t: number;
  r: number;
  b: number;
}

export type Fill =
  | { type: "none" }
  | { type: "solid"; color: RGBA }
  | { type: "gradient"; stops: GradientStop[]; angle: number; scaled: boolean; radial: boolean }
  | { type: "image"; mediaPath: string; mode: "stretch" | "tile"; crop: Crop | null };

export interface LineEnd {
  type: string;
  width?: string;
  length?: string;
}

export interface LineStyle {
  width: Emu | null;
  fill: Fill | null;
  dash: string | null;
  cap: string | null;
  compound: string | null;
  head: LineEnd | null;
  tail: LineEnd | null;
}

export interface Shadow {
  kind: "outer" | "inner";
  blur: Emu;
  dist: Emu;
  /** direction in degrees */
  dir: number;
  color: RGBA;
  align: string | null;
  scaleX: number | null;
  scaleY: number | null;
}

export type PathCommand =
  | { c: "M"; pts: number[] }
  | { c: "L"; pts: number[] }
  | { c: "C"; pts: number[] }
  | { c: "Q"; pts: number[] }
  | { c: "Z"; pts: number[] };

export type Geometry =
  | { type: "preset"; preset: string; adjust: Record<string, number> }
  | { type: "custom"; commands: PathCommand[]; w: number; h: number }
  | { type: "none" };

/* ------------------------------------------------------------------ text */

export interface RunProps {
  /** font size in CSS px (already converted from hundredths of a point) */
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: string | false;
  strike?: string | false;
  color?: RGBA;
  font?: string;
  /** baseline offset as a fraction of font size; negative is subscript */
  baseline?: number;
  /** character spacing in px */
  spacing?: number;
  caps?: "none" | "small" | "all";
  highlight?: RGBA;
}

export interface BulletText {
  char: string;
  font?: string;
  color?: RGBA;
  /** bullet size as a percentage of the text size (100 = same size) */
  sizePct?: number;
  sizePts?: number;
}

export type Bullet =
  | { type: "none" }
  | { type: "char"; text: BulletText }
  | { type: "autoNum"; scheme: string; startAt: number; text: BulletText }
  | { type: "image"; mediaPath: string };

export interface ParaProps {
  align?: "l" | "ctr" | "r" | "just" | "dist" | "justLow" | "thaiDist";
  level?: number;
  lineSpacing?: { type: "pct" | "pt"; value: number };
  /** px */
  spaceBefore?: number;
  /** px */
  spaceAfter?: number;
  /** px */
  indent?: { marL: number; marR: number; indent: number };
  bullet?: Bullet;
  rtl?: boolean;
  tabs?: number[];
}

export interface Hyperlink {
  href: string;
  tooltip?: string;
  /** false when the target scheme was rejected by the sanitiser */
  safe: boolean;
}

export interface Run {
  /** literal text; empty for line-break runs */
  text: string;
  /** true when the run is an `<a:br/>` */
  br?: boolean;
  /** true when the run came from `<a:fld>` (slide number / date placeholder) */
  field?: boolean;
  props: RunProps;
  link?: Hyperlink;
}

export interface Paragraph {
  props: ParaProps;
  runs: Run[];
}

export interface TextBody {
  anchor: "t" | "ctr" | "b" | "just" | "dist";
  wrap: "square" | "none";
  /** px */
  insets: { l: number; t: number; r: number; b: number };
  autofit: { type: "none" | "norm" | "shape"; fontScale: number; lineSpacingReduction: number };
  /** `vert` attribute for the body, e.g. `horz`, `vert270`, `eaVert` */
  vertical: string;
  paragraphs: Paragraph[];
  /** true only when at least one run carries non-whitespace text */
  hasText: boolean;
}

export interface LvlStyle {
  defRun: RunProps;
  pPr: ParaProps;
}

export interface Placeholder {
  type: string;
  idx: number | null;
  orient?: string;
  size?: string;
}

/* ---------------------------------------------------------------- shapes */

export type ShapeKind =
  | "shape"
  | "picture"
  | "graphicFrame"
  | "group"
  | "connector"
  | "table"
  | "chart"
  | "unsupported";

interface ShapeCommon {
  id: string;
  name: string | null;
  xfrm: Transform;
  link?: Hyperlink;
  title?: string;
  description?: string;
  hidden?: boolean;
  placeholder?: Placeholder;
  /** populated on the slide so consumers can jump to the offending element */
  path?: string;
}

export interface ShapeNode extends ShapeCommon {
  kind: "shape";
  geometry: Geometry;
  fill: Fill | null;
  line: LineStyle | null;
  shadow: Shadow | null;
  text: TextBody | null;
  isPlaceholder: boolean;
}

export interface PictureNode extends ShapeCommon {
  kind: "picture";
  /** resolved archive path of the image, or null when the relationship is missing */
  mediaPath: string | null;
  /** content type of the image, when known */
  mediaType: string | null;
  crop: Crop;
  fill: Fill | null;
  line: LineStyle | null;
  geometry: Geometry;
  shadow: Shadow | null;
  text: TextBody | null;
  alt: string | null;
  /** set when the image part exists but the browser cannot decode it (emf/wmf/tiff) */
  unrenderableMedia: string | null;
}

export interface GroupNode extends ShapeCommon {
  kind: "group";
  children: ShapeNodeUnion[];
  chOff: { x: number; y: number };
  chExt: { cx: number; cy: number };
}

export interface ConnectorNode extends ShapeCommon {
  kind: "connector";
  geometry: Geometry;
  line: LineStyle | null;
}

export interface TableBorderEdge {
  width: Emu;
  color: RGBA | null;
  dash: string | null;
}

export interface TableBorders {
  t?: TableBorderEdge | null;
  l?: TableBorderEdge | null;
  b?: TableBorderEdge | null;
  r?: TableBorderEdge | null;
}

export interface TableCell {
  gridSpan: number;
  rowSpan: number;
  hMerge: boolean;
  vMerge: boolean;
  fill: Fill | null;
  borders: TableBorders | null;
  /** px */
  insets: { l: number; t: number; r: number; b: number };
  anchor: "t" | "ctr" | "b" | null;
  text: TextBody | null;
  /** true when the cell is hidden because it is merged away */
  merged: boolean;
}

export interface TableRow {
  height: Emu;
  cells: TableCell[];
}

export interface TableData {
  colWidths: Emu[];
  rows: TableRow[];
  firstRow: boolean;
  bandRow: boolean;
  styleId: string | null;
}

export interface TableNode extends ShapeCommon {
  kind: "table";
  table: TableData;
}

export interface ChartSeries {
  name: string | null;
  values: (number | null)[];
  categories: (string | number)[] | null;
  xValues: (number | null)[] | null;
  color: RGBA | null;
  fill: Fill | null;
  line: LineStyle | null;
}

export interface ChartData {
  /** raw chart type name, e.g. `barChart`, `pieChart` */
  type: string;
  title: string | null;
  series: ChartSeries[];
  categories: (string | number)[];
  showLegend: boolean;
  legendPos: string;
  legendOverlay: boolean;
  /** column vs bar for barChart */
  barDir: "col" | "bar";
  grouping: string | null;
  overlap: number | null;
  gapWidth: number | null;
  holeSize: number;
  showValueLabels: boolean;
  valueAxisTitle: string | null;
  categoryAxisTitle: string | null;
  /** 1..48 Office chart style number, when present */
  style: string | null;
}

export interface ChartNode extends ShapeCommon {
  kind: "chart";
  chart: ChartData;
}

export interface UnsupportedNode extends ShapeCommon {
  kind: "unsupported";
  reason: string;
  label: string;
  /** a short plain-text summary, when one can be recovered */
  text: string | null;
}

export type ShapeNodeUnion =
  | ShapeNode
  | PictureNode
  | GroupNode
  | ConnectorNode
  | TableNode
  | ChartNode
  | UnsupportedNode;

/* --------------------------------------------------------------- package */

export interface Diagnostic {
  severity: "info" | "warning" | "error";
  code: string;
  message: string;
  slideIndex?: number;
  part?: string;
  element?: string;
}

export interface ThemeData {
  colors: Record<string, RGBA>;
  majorFont: string | null;
  minorFont: string | null;
  fillStyles: Fill[];
  bgFillStyles: Fill[];
  lineStyles: LineStyle[];
  effectStyles: (Shadow | null)[];
}

export type ColorMap = Record<string, string>;

export interface TextStyles {
  title: (LvlStyle | null)[];
  body: (LvlStyle | null)[];
  other: (LvlStyle | null)[];
}

export interface SlideModel {
  /** 0-based presentation order */
  index: number;
  partName: string;
  name: string | null;
  slideshowHidden: boolean;
  shapes: ShapeNodeUnion[];
  /** shapes inherited from the layout and master, drawn behind slide shapes */
  inheritedShapes: ShapeNodeUnion[];
  background: Fill | null;
  diagnostics: Diagnostic[];
}

export interface PackageParts {
  /** every archive path we were able to read, used for diagnostics */
  names: string[];
}

export interface PresentationModel {
  width: Emu;
  height: Emu;
  slides: SlideModel[];
  theme: ThemeData;
  colorMap: ColorMap;
  textStyles: TextStyles;
  title: string | null;
  author: string | null;
  /** number of archive entries in the source package */
  partCount: number;
  diagnostics: Diagnostic[];
}
