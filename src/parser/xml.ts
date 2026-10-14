/**
 * Small XML layer built on `fast-xml-parser`'s `preserveOrder` mode.
 *
 * `preserveOrder` matters here: the order of children inside a `p:spTree` is the
 * z-order of a slide's shapes, and object-key ordering in a normal parse would
 * silently regroup shapes by element name.
 *
 * Security notes:
 *  - Entity processing is disabled (`processEntities: false`) so crafted DTD /
 *    entity expansion cannot be used for a memory blow-up. The five predefined
 *    entities and numeric character references are decoded here instead, with a
 *    bounded output size.
 *  - Input containing a DOCTYPE declaration with an internal subset (the shape
 *    required for entity expansion tricks) is rejected outright. PresentationML
 *    never needs one.
 *  - Callers must enforce a maximum document size before calling `parseXml`.
 */

import { XMLParser } from "fast-xml-parser";

export interface XmlNode {
  /** qualified name, e.g. `p:sp` */
  name: string;
  /** local name, e.g. `sp` */
  local: string;
  attrs: Record<string, string>;
  children: XmlNode[];
  /** concatenated direct text content */
  text: string;
}

export class XmlError extends Error {
  constructor(
    message: string,
    readonly detail?: string,
  ) {
    super(message);
    this.name = "XmlError";
  }
}

const MAX_TEXT_LENGTH = 4_000_000;

const PREDEFINED: Record<string, string> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
};

/** Decode XML entities without ever allocating unbounded output. */
export function decodeXmlText(input: string): string {
  if (input.indexOf("&") === -1) return input;
  let out = "";
  let i = 0;
  while (i < input.length) {
    const amp = input.indexOf("&", i);
    if (amp === -1) {
      out += input.slice(i);
      break;
    }
    out += input.slice(i, amp);
    const semi = input.indexOf(";", amp + 1);
    if (semi === -1 || semi - amp > 12) {
      // Not a valid reference; keep the literal ampersand.
      out += "&";
      i = amp + 1;
      continue;
    }
    const body = input.slice(amp + 1, semi);
    let decoded: string | null = null;
    if (body.charCodeAt(0) === 35 /* # */) {
      const isHex = body[1] === "x" || body[1] === "X";
      const digits = isHex ? body.slice(2) : body.slice(1);
      if (digits.length > 0 && /^[0-9a-fA-F]+$/.test(digits)) {
        const code = parseInt(digits, isHex ? 16 : 10);
        if (Number.isFinite(code) && code > 0 && code <= 0x10ffff) {
          try {
            decoded = String.fromCodePoint(code);
          } catch {
            decoded = null;
          }
        }
      }
    } else if (Object.prototype.hasOwnProperty.call(PREDEFINED, body)) {
      decoded = PREDEFINED[body] as string;
    }
    if (decoded === null) {
      out += "&";
      i = amp + 1;
    } else {
      out += decoded;
      i = semi + 1;
      if (out.length > MAX_TEXT_LENGTH) {
        throw new XmlError("XML text content exceeded the supported size");
      }
    }
  }
  return out;
}

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "",
  preserveOrder: true,
  trimValues: false,
  parseAttributeValue: false,
  parseTagValue: false,
  processEntities: false,
  allowBooleanAttributes: false,
  alwaysCreateTextNode: false,
});

type RawNode = Record<string, unknown>;

/**
 * Element names only. `preserveOrder` also emits keys for the XML declaration
 * (`?xml`), comments and doctypes, all of which are skipped by this test.
 */
const ELEMENT_NAME = /^[A-Za-z_][\w.:-]*$/;

function localNameOf(name: string): string {
  const colon = name.indexOf(":");
  return colon === -1 ? name : name.slice(colon + 1);
}

function convert(raw: unknown): XmlNode | null {
  if (typeof raw !== "object" || raw === null) return null;
  const record = raw as RawNode;
  const keys = Object.keys(record);
  let name: string | null = null;
  for (const key of keys) {
    if (key === ":@") continue;
    if (!ELEMENT_NAME.test(key)) return null; // declaration, comment, doctype, text
    name = key;
    break;
  }
  if (name === null) return null;

  const node: XmlNode = {
    name,
    local: localNameOf(name),
    attrs: {},
    children: [],
    text: "",
  };

  const rawAttrs = record[":@"];
  if (typeof rawAttrs === "object" && rawAttrs !== null) {
    for (const [key, value] of Object.entries(rawAttrs as Record<string, unknown>)) {
      if (key === "xmlns" || key.startsWith("xmlns:")) continue;
      node.attrs[key] = typeof value === "string" ? value : String(value);
    }
  }

  const body = record[name];
  if (Array.isArray(body)) {
    for (const child of body) {
      if (typeof child !== "object" || child === null) continue;
      const childKeys = Object.keys(child as RawNode);
      const first = childKeys[0];
      if (first === undefined) continue;
      if (first === "#text") {
        const raw = (child as RawNode)["#text"];
        if (typeof raw === "string") node.text += decodeXmlText(raw);
        continue;
      }
      if (!ELEMENT_NAME.test(first)) continue; // comments, cdata, doctype fragments
      const converted = convert(child);
      if (converted) node.children.push(converted);
    }
  }
  return node;
}

export interface ParseXmlOptions {
  /** part name used in error messages */
  part?: string;
}

/**
 * Parse an XML document into an ordered {@link XmlNode} tree.
 * Returns `null` for an empty document; throws {@link XmlError} for malformed or
 * hostile input.
 */
export function parseXml(source: string, options: ParseXmlOptions = {}): XmlNode | null {
  const where = options.part ? ` in ${options.part}` : "";
  if (source.length === 0) return null;
  const head = source.slice(0, 4096);
  if (/<!doctype/i.test(head) || /<!entity/i.test(head)) {
    throw new XmlError(`Refused to parse XML with a DOCTYPE/ENTITY declaration${where}`);
  }
  let raw: unknown;
  try {
    raw = parser.parse(source);
  } catch (error) {
    throw new XmlError(
      `Malformed XML${where}`,
      error instanceof Error ? error.message : String(error),
    );
  }
  if (!Array.isArray(raw)) return null;
  for (const item of raw) {
    if (typeof item !== "object" || item === null) continue;
    const converted = convert(item);
    if (converted) return converted;
  }
  return null;
}

/* ------------------------------------------------------------- traversal */

export function attr(node: XmlNode | null | undefined, name: string): string | null {
  if (!node) return null;
  const value = node.attrs[name];
  return value === undefined ? null : value;
}

/**
 * Read a relationship attribute regardless of how the producer prefixed it.
 * Office writes `r:embed`, but hand-written and converted decks sometimes use a
 * different prefix or none at all.
 */
export function relAttr(node: XmlNode | null | undefined, name: string): string | null {
  if (!node) return null;
  const direct = node.attrs[name];
  if (direct !== undefined) return direct;
  const qualified = `r:${name}`;
  if (node.attrs[qualified] !== undefined) return node.attrs[qualified] as string;
  const suffix = `:${name}`;
  for (const [key, value] of Object.entries(node.attrs)) {
    if (key.endsWith(suffix)) return value;
  }
  return null;
}

export function attrNum(node: XmlNode | null | undefined, name: string): number | null {
  const raw = attr(node, name);
  if (raw === null || raw.length === 0) return null;
  const value = Number(raw);
  return Number.isFinite(value) ? value : null;
}

export function attrBool(node: XmlNode | null | undefined, name: string): boolean | null {
  const raw = attr(node, name);
  if (raw === null) return null;
  return raw === "1" || raw.toLowerCase() === "true";
}

/** First child matching a local name. */
export function child(node: XmlNode | null | undefined, local: string): XmlNode | null {
  if (!node) return null;
  for (const candidate of node.children) {
    if (candidate.local === local) return candidate;
  }
  return null;
}

/** All children matching a local name, in document order. */
export function children(node: XmlNode | null | undefined, local: string): XmlNode[] {
  if (!node) return [];
  const out: XmlNode[] = [];
  for (const candidate of node.children) {
    if (candidate.local === local) out.push(candidate);
  }
  return out;
}

export function hasChild(node: XmlNode | null | undefined, local: string): boolean {
  return child(node, local) !== null;
}

/** Depth-first search for the first descendant (excluding `node` itself). */
export function find(node: XmlNode | null | undefined, local: string): XmlNode | null {
  if (!node) return null;
  const stack: XmlNode[] = [...node.children];
  while (stack.length > 0) {
    const current = stack.shift() as XmlNode;
    if (current.local === local) return current;
    for (const candidate of current.children) stack.push(candidate);
  }
  return null;
}

export function findAll(node: XmlNode | null | undefined, local: string): XmlNode[] {
  const out: XmlNode[] = [];
  if (!node) return out;
  const stack: XmlNode[] = [...node.children];
  while (stack.length > 0) {
    const current = stack.shift() as XmlNode;
    if (current.local === local) out.push(current);
    for (const candidate of current.children) stack.push(candidate);
  }
  return out;
}

/**
 * Unwrap `mc:AlternateContent`.
 *
 * The renderer implements a deliberately small feature set, so the `Fallback`
 * branch is preferred whenever it exists: it is by definition the representation
 * that pre-Markup-Compatibility clients understand. The `Choice` branch is used
 * only when no fallback is present.
 */
export function unwrapAlternate(node: XmlNode, onChoice: () => void): XmlNode[] {
  if (node.local !== "AlternateContent") return [node];
  const fallback = child(node, "Fallback");
  if (fallback) {
    return fallback.children.filter((c) => !c.name.startsWith("#"));
  }
  onChoice();
  const choice = child(node, "Choice");
  if (choice) return choice.children.filter((c) => !c.name.startsWith("#"));
  return [];
}
