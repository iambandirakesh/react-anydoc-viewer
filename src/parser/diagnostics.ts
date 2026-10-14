/**
 * Diagnostic collection shared by the parsers.
 *
 * Parsers are expected to explain what they could not render instead of
 * silently dropping it. Duplicates are collapsed (the same construct usually
 * appears on many slides) and the list is capped so a pathological file cannot
 * grow it without bound.
 */

import type { Diagnostic } from "../model/types";

export const DEFAULT_MAX_DIAGNOSTICS = 200;

export class Diagnostics {
  private readonly seen = new Set<string>();
  private readonly all: Diagnostic[] = [];
  private truncated = 0;

  constructor(private readonly max: number = DEFAULT_MAX_DIAGNOSTICS) {}

  add(diagnostic: Diagnostic): void {
    const key = `${diagnostic.severity}|${diagnostic.code}|${diagnostic.message}|${diagnostic.part ?? ""}|${diagnostic.slideIndex ?? ""}`;
    if (this.seen.has(key)) return;
    this.seen.add(key);
    if (this.all.length >= this.max) {
      this.truncated += 1;
      return;
    }
    this.all.push(diagnostic);
  }

  /** Index to pass to {@link addedSince} to capture a slice of diagnostics. */
  mark(): number {
    return this.all.length;
  }

  addedSince(mark: number): Diagnostic[] {
    return this.all.slice(mark);
  }

  list(): Diagnostic[] {
    const out = [...this.all];
    if (this.truncated > 0) {
      out.push({
        severity: "info",
        code: "diagnostics.truncated",
        message: `${this.truncated} further diagnostic message(s) were suppressed.`,
      });
    }
    return out;
  }
}
