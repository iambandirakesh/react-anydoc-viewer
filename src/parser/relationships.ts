/**
 * OPC relationship resolution.
 *
 * Slides, layouts, masters and charts are located by following relationships
 * from the package root. Filenames are never used to infer identity or order,
 * because they are neither sequential nor guaranteed.
 */

import { attr, children, parseXml, type XmlNode } from "./xml";
import { readText, type Archive } from "./archive";
import { XmlError } from "./xml";

export interface Relationship {
  id: string;
  type: string;
  /** absolute part name for internal targets, or the URL for external targets */
  target: string;
  external: boolean;
}

export type RelationshipMap = Map<string, Relationship>;

export const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export function relTypeName(type: string): string {
  const slash = type.lastIndexOf("/");
  return slash === -1 ? type : type.slice(slash + 1);
}

/** Resolve a relationship target relative to the part that declared it. */
export function resolvePartName(sourcePart: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const sourceDir = sourcePart.includes("/") ? sourcePart.slice(0, sourcePart.lastIndexOf("/")) : "";
  const segments = (sourceDir.length > 0 ? `${sourceDir}/${target}` : target).split("/");
  const out: string[] = [];
  for (const segment of segments) {
    if (segment === "" || segment === ".") continue;
    if (segment === "..") {
      out.pop();
      continue;
    }
    out.push(segment);
  }
  return out.join("/");
}

export function relsPartNameFor(partName: string): string {
  const slash = partName.lastIndexOf("/");
  const dir = slash === -1 ? "" : partName.slice(0, slash + 1);
  const base = slash === -1 ? partName : partName.slice(slash + 1);
  return `${dir}_rels/${base}.rels`;
}

export function emptyRelationships(): RelationshipMap {
  return new Map();
}

export function parseRelationships(xml: string | null, sourcePart: string): RelationshipMap {
  const map: RelationshipMap = new Map();
  if (!xml) return map;
  const root = parseXml(xml, { part: sourcePart });
  if (!root) return map;
  for (const node of root.children) {
    if (node.local !== "Relationship") continue;
    const id = attr(node, "Id");
    const type = attr(node, "Type");
    const target = attr(node, "Target");
    if (!id || !type || !target) continue;
    const external = (attr(node, "TargetMode") ?? "Internal") === "External";
    map.set(id, {
      id,
      type,
      target: external ? target : resolvePartName(sourcePart, target),
      external,
    });
  }
  return map;
}

const relCache = new WeakMap<Archive, Map<string, RelationshipMap>>();

/** Read and cache the relationships declared by a part. */
export function relationshipsFor(archive: Archive, partName: string): RelationshipMap {
  let perArchive = relCache.get(archive);
  if (!perArchive) {
    perArchive = new Map();
    relCache.set(archive, perArchive);
  }
  const key = partName.toLowerCase();
  const cached = perArchive.get(key);
  if (cached) return cached;
  const relsPart = relsPartNameFor(partName);
  let map: RelationshipMap;
  try {
    map = parseRelationships(readText(archive, relsPart), partName);
  } catch (error) {
    if (error instanceof XmlError) {
      // A malformed relationship part must not take the whole deck down: the
      // affected slide will report a diagnostic and render what it can.
      map = new Map();
    } else {
      throw error;
    }
  }
  perArchive.set(key, map);
  return map;
}

export function findRelationship(
  archive: Archive,
  partName: string,
  typeSuffix: string,
): Relationship | null {
  const map = relationshipsFor(archive, partName);
  for (const rel of map.values()) {
    if (relTypeName(rel.type) === typeSuffix) return rel;
  }
  return null;
}

export function findAllRelationships(
  archive: Archive,
  partName: string,
  typeSuffix: string,
): Relationship[] {
  const out: Relationship[] = [];
  for (const rel of relationshipsFor(archive, partName).values()) {
    if (relTypeName(rel.type) === typeSuffix) out.push(rel);
  }
  return out;
}

/**
 * The package root relationship that identifies the main presentation part.
 * Falls back to the conventional name when the root rels are missing, so that
 * slightly malformed but otherwise readable decks still open.
 */
export function findOfficeDocumentPart(archive: Archive): { part: string; fromFallback: boolean } {
  const rels = parseRelationships(readText(archive, "_rels/.rels"), "");
  for (const rel of rels.values()) {
    if (relTypeName(rel.type) === "officeDocument" && !rel.external) {
      return { part: rel.target, fromFallback: false };
    }
  }
  return { part: "ppt/presentation.xml", fromFallback: true };
}

/** Elements that carry a relationship id alongside their own namespace. */
export function relationshipIdOf(node: XmlNode | null): string | null {
  if (!node) return null;
  return attr(node, `${"r"}:id`) ?? attr(node, "r:id") ?? attr(node, "id");
}

export function hyperlinkTarget(
  archive: Archive,
  sourcePart: string,
  id: string,
): Relationship | null {
  return relationshipsFor(archive, sourcePart).get(id) ?? null;
}

export function relationshipChildren(node: XmlNode | null): XmlNode[] {
  return children(node, "Relationship");
}
