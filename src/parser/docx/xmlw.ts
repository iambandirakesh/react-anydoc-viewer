/**
 * WordprocessingML attribute helpers.
 *
 * Element names are fine as-is (the XML layer matches on local names), but
 * attributes are not: Word writes most of them with the `w:` prefix
 * (`<w:jc w:val="center"/>`) while other producers leave it off, and the XML
 * layer keeps attribute keys verbatim. Every attribute read therefore falls back
 * to a prefixed lookup, so both spellings work.
 */

import { type XmlNode } from "../xml";

export function wAttr(node: XmlNode | null | undefined, name: string): string | null {
  if (!node) return null;
  const direct = node.attrs[name];
  if (direct !== undefined) return direct;
  const qualified = node.attrs[`w:${name}`];
  if (qualified !== undefined) return qualified;
  const suffix = `:${name}`;
  for (const [key, value] of Object.entries(node.attrs)) {
    if (key.endsWith(suffix)) return value;
  }
  return null;
}

export function wNum(node: XmlNode | null | undefined, name: string): number | null {
  const raw = wAttr(node, name);
  if (raw === null || raw.length === 0) return null;
  const value = Number(raw);
  return Number.isFinite(value) ? value : null;
}

/**
 * Word's on/off attribute: `<w:b/>` and `<w:b w:val="1"/>` are on, and
 * `<w:b w:val="0"/>`, `"false"` and `"off"` are off.
 */
export function wToggle(node: XmlNode | null | undefined): boolean | null {
  if (!node) return null;
  const raw = wAttr(node, "val");
  if (raw === null) return true;
  const value = raw.trim().toLowerCase();
  if (value === "0" || value === "false" || value === "off" || value === "no") return false;
  return true;
}

/** `w:val`, `w:styleId` and friends, tolerating an unprefixed attribute. */
export function wVal(node: XmlNode | null | undefined): string | null {
  return wAttr(node, "val");
}

/** True when the attribute is present (Word uses presence for flags). */
export function wHas(node: XmlNode | null | undefined, name: string): boolean {
  return wAttr(node, name) !== null;
}

/** `xml:space="preserve"` — the only case where a Word run keeps its spaces. */
export function preservesSpace(node: XmlNode | null | undefined): boolean {
  if (!node) return false;
  const raw = wAttr(node, "space") ?? node.attrs["xml:space"] ?? null;
  return raw === "preserve";
}
