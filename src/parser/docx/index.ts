/**
 * Word document orchestration (`word/document.xml` -> pages).
 *
 * The pipeline mirrors the presentation parser: read the parts, resolve the
 * relationships, flatten the body into blocks, paginate, and hand back the same
 * `PresentationModel` the slide renderer already understands. A page of a Word
 * document is therefore just a slide whose shapes are paragraphs, tables and
 * pictures, so navigation, thumbnails, zoom and presentation mode all work
 * without knowing what a `.docx` is.
 */

import type { PresentationModel, SlideModel } from "../../model/types";
import { DocumentViewerError } from "../../errors";
import { readText, type Archive, type ArchiveLimits } from "../archive";
import { child, find, parseXml, type XmlNode } from "../xml";
import { parseRelationships, relsPartNameFor, relTypeName, type RelationshipMap } from "../relationships";
import { createMediaStore, type MediaStore } from "../media";
import { DEFAULT_COLOR_MAP, createEmptyTheme } from "../theme";
import { twipsToEmu } from "../units";
import { Diagnostics, DEFAULT_MAX_DIAGNOSTICS } from "../diagnostics";
import { createParagraphMeasurer, type MeasurerHandle } from "../../renderer/measureParagraph";
import { layoutDocument, type PageGeometry } from "./layout";
import type { DocxBlock, ParagraphMeasurer } from "./model";
import { parseTableBlock, type TableParseContext } from "./table";
import { parseParagraphBlocks, type DocxTextContext } from "./text";
import { buildStyles } from "./styles";

/** US Letter, the size Word uses when `w:pgSz` is absent. */
const DEFAULT_PAGE_TWIPS = { width: 12_240, height: 15_840 };
const DEFAULT_MARGIN_TWIPS = 1_440;

export interface DocxParseOptions {
  limits?: Partial<ArchiveLimits>;
  maxDiagnostics?: number;
  /** overrides height measurement; the DOM measurer is used by default */
  measure?: ParagraphMeasurer;
}

export interface ParsedDocx {
  model: PresentationModel;
  media: MediaStore;
  dispose(): void;
}

type Warn = (code: string, message: string, part?: string) => void;

/** `w:pgSz`/`w:pgMar` attributes are written with the `w:` prefix by Word. */
function wAttr(node: XmlNode | null, name: string): string | null {
  if (!node) return null;
  const direct = node.attrs[name];
  if (direct !== undefined) return direct;
  const qualified = node.attrs[`w:${name}`];
  if (qualified !== undefined) return qualified;
  for (const [key, value] of Object.entries(node.attrs)) {
    if (key.endsWith(`:${name}`)) return value;
  }
  return null;
}

function locateMainPart(archive: Archive): string {
  const rels = parseRelationships(readText(archive, "_rels/.rels"), "");
  for (const rel of rels.values()) {
    if (relTypeName(rel.type) === "officeDocument" && !rel.external) return rel.target;
  }
  return "word/document.xml";
}

/**
 * Page setup comes from the last `w:sectPr` in the body: that is the section
 * whose properties Word applies to the final pages, and using it for the whole
 * document is a deliberate approximation for multi-section files.
 */
function sectionGeometry(body: XmlNode | null): { geometry: PageGeometry; sections: number } {
  const sections: XmlNode[] = [];
  for (const item of body?.children ?? []) {
    if (item.local === "sectPr") sections.push(item);
    else if (item.local === "p") {
      const sectPr = child(child(item, "pPr"), "sectPr");
      if (sectPr) sections.push(sectPr);
    }
  }
  const sectPr = sections.length > 0 ? (sections[sections.length - 1] as XmlNode) : null;

  const pgSz = child(sectPr, "pgSz");
  const widthTwips = Number(wAttr(pgSz, "w")) || DEFAULT_PAGE_TWIPS.width;
  const heightTwips = Number(wAttr(pgSz, "h")) || DEFAULT_PAGE_TWIPS.height;
  const landscape = (wAttr(pgSz, "orient") ?? "portrait").toLowerCase() === "landscape";

  const pgMar = child(sectPr, "pgMar");
  const margin = (name: string): number => {
    const raw = wAttr(pgMar, name);
    const value = raw === null || raw === "" ? DEFAULT_MARGIN_TWIPS : Number(raw);
    return twipsToEmu(Number.isFinite(value) ? value : DEFAULT_MARGIN_TWIPS);
  };

  return {
    geometry: {
      // A landscape section still declares portrait dimensions plus `orient`, but
      // a custom page size is taken exactly as written.
      widthEmu: twipsToEmu(landscape ? Math.max(widthTwips, heightTwips) : widthTwips),
      heightEmu: twipsToEmu(landscape ? Math.min(widthTwips, heightTwips) : heightTwips),
      margins: { top: margin("top"), right: margin("right"), bottom: margin("bottom"), left: margin("left") },
    },
    sections: sections.length,
  };
}

function readCoreProperties(archive: Archive): { title: string | null; author: string | null } {
  const xml = readText(archive, "docProps/core.xml");
  if (!xml) return { title: null, author: null };
  let root: XmlNode | null = null;
  try {
    root = parseXml(xml, { part: "docProps/core.xml" });
  } catch {
    return { title: null, author: null };
  }
  const text = (local: string): string | null => {
    const value = find(root, local)?.text.trim() ?? "";
    return value.length > 0 ? value : null;
  };
  return { title: text("title"), author: text("creator") };
}

function emptyPage(index: number, part: string): SlideModel {
  return {
    index,
    partName: `${part}#page${index + 1}`,
    name: null,
    slideshowHidden: false,
    shapes: [],
    inheritedShapes: [],
    background: { type: "solid", color: { r: 255, g: 255, b: 255, a: 1 } },
    diagnostics: [],
  };
}

function tryParse(archive: Archive, partName: string): XmlNode | null {
  try {
    const xml = readText(archive, partName);
    return xml === null ? null : parseXml(xml, { part: partName });
  } catch {
    // Styles and numbering are optional: a damaged one only costs formatting.
    return null;
  }
}

export function parseDocxArchive(archive: Archive, options: DocxParseOptions = {}): ParsedDocx {
  const diagnostics = new Diagnostics(options.maxDiagnostics ?? DEFAULT_MAX_DIAGNOSTICS);
  const warn: Warn = (code, message, part) => {
    diagnostics.add({ severity: "warning", code, message, ...(part ? { part } : {}) });
  };

  const mainPart = locateMainPart(archive);
  const documentXml = readText(archive, mainPart);
  if (documentXml === null) {
    throw new DocumentViewerError("MALFORMED_DOCUMENT", "The package does not contain a Word document part.", {
      detail: `${mainPart} is missing`,
    });
  }
  const root = parseXml(documentXml, { part: mainPart });
  const body = child(root, "body");
  if (!body) {
    throw new DocumentViewerError("MALFORMED_DOCUMENT", "The Word document has no body.", { detail: mainPart });
  }

  const styles = buildStyles(tryParse(archive, "word/styles.xml"), tryParse(archive, "word/numbering.xml"));

  let relationships: RelationshipMap;
  try {
    relationships = parseRelationships(readText(archive, relsPartNameFor(mainPart)), mainPart);
  } catch {
    // A damaged relationship part must not lose the document itself.
    relationships = new Map();
    warn("docx.relationships", "The document's relationships could not be read; pictures may be missing.", mainPart);
  }

  const resolveRel = (id: string): { target: string; external: boolean } | null => {
    const rel = relationships.get(id);
    if (!rel) return null;
    // `parseRelationships` already resolved internal targets against the source
    // part, so the target is used as-is here.
    return { target: rel.target, external: rel.external };
  };

  const { geometry, sections } = sectionGeometry(body);
  if (sections > 1) {
    warn(
      "docx.section-break",
      "The document has multiple sections; section breaks were rendered as page breaks and the last section's page setup was applied to all pages.",
      mainPart,
    );
  }

  const measurer: MeasurerHandle = options.measure
    ? { measure: options.measure, dispose: () => undefined, kind: "layout" }
    : createParagraphMeasurer();
  if (!options.measure && measurer.kind === "estimate") {
    diagnostics.add({
      severity: "info",
      code: "docx.layout.estimated",
      message: "Page breaks were estimated because text could not be measured in this environment.",
    });
  }

  const textContext: DocxTextContext = {
    styles,
    part: mainPart,
    resolveRel,
    warn: (code, message) => warn(code, message, mainPart),
  };
  const tableContext: TableParseContext = {
    ...textContext,
    measure: measurer.measure,
    contentWidthEmu: Math.max(geometry.widthEmu - geometry.margins.left - geometry.margins.right, 1),
  };

  const blocks: DocxBlock[] = [];

  const walk = (container: XmlNode): void => {
    for (const item of container.children) {
      switch (item.local) {
        case "p": {
          blocks.push(...parseParagraphBlocks(item, textContext));
          // A paragraph that carries section properties ends its section, which
          // is a page break as far as a single-section viewer is concerned.
          if (child(child(item, "pPr"), "sectPr")) blocks.push({ kind: "break", part: mainPart });
          break;
        }
        case "tbl": {
          const table = parseTableBlock(item, tableContext);
          if (table) blocks.push(table);
          break;
        }
        case "sectPr":
          // The body's own section properties describe the pages, not a break.
          break;
        case "sdt":
        case "customXml":
        case "ins":
        case "moveTo":
          walk(child(item, "sdtContent") ?? item);
          break;
        case "del":
        case "moveFrom":
          break;
        case "altChunk":
          warn("docx.altChunk", "Embedded (altChunk) content is not supported and was skipped.", mainPart);
          break;
        default:
          break;
      }
    }
  };
  walk(body);

  const hasPart = (pattern: RegExp): boolean => archive.names.some((name) => pattern.test(name.toLowerCase()));
  if (hasPart(/^word\/header\d*\.xml$/) || hasPart(/^word\/footer\d*\.xml$/)) {
    warn("docx.header-footer", "Headers and footers are not rendered.", mainPart);
  }
  if (hasPart(/^word\/footnotes\.xml$/) || hasPart(/^word\/endnotes\.xml$/)) {
    warn("docx.notes", "Footnotes and endnotes are not rendered.", mainPart);
  }
  if (hasPart(/^word\/comments\d*\.xml$/)) {
    warn("docx.comments", "Comments are not rendered.", mainPart);
  }

  const { pages, overflowed } = layoutDocument(blocks, {
    geometry,
    measure: measurer.measure,
    part: mainPart,
    pagePart: (index) => `${mainPart}#page${index + 1}`,
  });
  measurer.dispose();

  if (overflowed > 0) {
    warn("docx.page.overflow", `${overflowed} block(s) are taller than the page and were clipped.`, mainPart);
  }

  const { title, author } = readCoreProperties(archive);
  const media = createMediaStore(archive);
  const slides = pages.length > 0 ? pages : [emptyPage(0, mainPart)];

  const model: PresentationModel = {
    width: geometry.widthEmu,
    height: geometry.heightEmu,
    slides,
    theme: createEmptyTheme(),
    colorMap: DEFAULT_COLOR_MAP,
    textStyles: { title: [], body: [], other: [] },
    title,
    author,
    partCount: archive.names.length,
    diagnostics: diagnostics.list(),
  };

  return { model, media, dispose: () => media.dispose() };
}
