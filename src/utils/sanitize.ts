/**
 * Sanitisation helpers.
 *
 * Presentation content is untrusted input: a deck can contain hyperlinks whose
 * target is `javascript:` or `data:`. Everything that ends up in an `href` goes
 * through {@link sanitizeUrl} first.
 */

export const SAFE_PROTOCOLS = ["http:", "https:", "mailto:", "tel:"] as const;

export interface SanitizeResult {
  safe: boolean;
  href: string;
  reason?: string;
}

/**
 * Validate a hyperlink target. Relative URLs are rejected because the viewer has
 * no meaningful base to resolve them against in a document-relative sense.
 */
export function sanitizeUrl(raw: string, options: { allowRelative?: boolean } = {}): SanitizeResult {
  const value = raw.trim();
  if (value.length === 0) return { safe: false, href: "", reason: "empty" };
  if (value.length > 2083) return { safe: false, href: "", reason: "too-long" };

  let parsed: URL;
  try {
    parsed = new URL(value, "https://invalid.invalid/");
  } catch {
    return { safe: false, href: "", reason: "unparseable" };
  }

  const protocol = parsed.protocol.toLowerCase();
  if (protocol === "https:" && parsed.origin === "https://invalid.invalid") {
    // Actually a bare relative path.
    if (options.allowRelative) return { safe: true, href: value };
    return { safe: false, href: "", reason: "relative" };
  }
  if (!(SAFE_PROTOCOLS as readonly string[]).includes(protocol)) {
    return { safe: false, href: "", reason: `blocked-protocol:${protocol}` };
  }
  // Strip credentials: a deck should never be able to smuggle `user:pass@host`.
  if (parsed.username || parsed.password) {
    return { safe: false, href: "", reason: "credentials" };
  }
  return { safe: true, href: parsed.toString() };
}

/** Escape a string for use inside a CSS `content:` value or attribute. */
export function cssSafeText(value: string): string {
  return value.replace(/["\\]/g, "\\$&");
}
