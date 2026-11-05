import { useCallback, useEffect, useRef, useState, type MutableRefObject } from "react";
import type { Diagnostic } from "../model/types";

const MAX_RENDER_ISSUES = 40;

export interface RenderIssues {
  /** issues reported while drawing, deduplicated per page */
  issues: Diagnostic[];
  /** safe to call during a child's render */
  report: (code: string, message: string) => void;
}

/**
 * Problems the renderer finds while drawing a page.
 *
 * Issues are reported from inside the renderer, i.e. while a child component is
 * rendering. Updating state there trips React's "cannot update a component while
 * rendering a different component" warning (and concurrent rendering may drop
 * the update), so reports are buffered and flushed right after the commit.
 */
export function useRenderIssues(pageIndexRef: MutableRefObject<number>, resetKey: number): RenderIssues {
  const [issues, setIssues] = useState<Diagnostic[]>([]);
  const issueKeys = useRef<Set<string>>(new Set());
  const pending = useRef<Diagnostic[]>([]);

  const report = useCallback(
    (code: string, message: string) => {
      const slideIndex = pageIndexRef.current;
      const key = `${slideIndex}|${code}|${message}`;
      if (issueKeys.current.has(key) || issueKeys.current.size >= MAX_RENDER_ISSUES) return;
      issueKeys.current.add(key);
      pending.current.push({ severity: "warning", code, message, slideIndex });
    },
    [pageIndexRef],
  );

  useEffect(() => {
    if (pending.current.length === 0) return;
    const queued = pending.current;
    pending.current = [];
    setIssues((previous) => [...previous, ...queued]);
  });

  // A newly loaded document starts with a clean list.
  useEffect(() => {
    if (resetKey === 0) return;
    issueKeys.current = new Set();
    pending.current = [];
    setIssues([]);
  }, [resetKey]);

  return { issues, report };
}
