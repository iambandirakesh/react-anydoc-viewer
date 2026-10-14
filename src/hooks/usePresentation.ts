import { useCallback, useEffect, useRef, useState } from "react";
import type { Diagnostic, PresentationModel } from "../model/types";
import type { MediaStore } from "../parser/media";
import { DocumentViewerError, type DocumentFormat } from "../errors";
import {
  isViewerSource,
  normalizeError,
  type LoadOptions,
  type ViewerSource,
} from "../parser/loadPresentation";
import { loadDocument } from "../parser/document";
import type { PageRenderer } from "../renderer/pageRenderer";

export type PresentationStatus = "idle" | "loading" | "ready" | "error";

export interface PresentationState {
  status: PresentationStatus;
  model: PresentationModel | null;
  media: MediaStore | null;
  error: DocumentViewerError | null;
  diagnostics: Diagnostic[];
  /** what the loaded bytes turned out to be */
  format: DocumentFormat | null;
  /** set when the format draws its pages itself instead of using shapes (PDF) */
  renderPage: PageRenderer | null;
  /** increments once per successful load */
  loadId: number;
}

export interface UsePresentationOptions extends LoadOptions {
  /** called with the parsed model and its format; the viewer maps this to public info */
  onLoad?: (model: PresentationModel, format: DocumentFormat) => void;
  onError?: (error: DocumentViewerError) => void;
  onDiagnostics?: (diagnostics: Diagnostic[]) => void;
}

export interface UsePresentationResult extends PresentationState {
  /** re-fetch and re-parse the current source */
  reload: () => void;
}

const EMPTY: Omit<PresentationState, "loadId"> = {
  status: "idle",
  model: null,
  media: null,
  error: null,
  diagnostics: [],
  format: null,
  renderPage: null,
};

/**
 * Load and parse a presentation.
 *
 * Cancellation is enforced with an abort signal *and* a generation counter, so
 * a result from an earlier source can never be committed after a newer one has
 * started. Object URLs live in a ref and are revoked from effects (never from a
 * state updater, which React may invoke more than once).
 */
export function usePresentation(
  source: ViewerSource | null | undefined,
  options: UsePresentationOptions = {},
): UsePresentationResult {
  const [state, setState] = useState<PresentationState>({ ...EMPTY, loadId: 0 });
  const [reloadToken, setReloadToken] = useState(0);

  const generation = useRef(0);
  const mediaRef = useRef<MediaStore | null>(null);
  const optionsRef = useRef(options);
  optionsRef.current = options;

  const reload = useCallback(() => {
    setReloadToken((token) => token + 1);
  }, []);

  useEffect(() => {
    const currentGeneration = generation.current + 1;
    generation.current = currentGeneration;
    const controller = new AbortController();

    // Release the previous run's object URLs before starting a new one.
    mediaRef.current?.dispose();
    mediaRef.current = null;

    if (source === null || source === undefined) {
      setState((previous) =>
        previous.status === "idle" && previous.loadId === 0 ? previous : { ...EMPTY, loadId: previous.loadId },
      );
      return () => controller.abort();
    }

    if (!isViewerSource(source)) {
      const error = new DocumentViewerError(
        "UNSUPPORTED_INPUT",
        "The source must be a File, Blob, ArrayBuffer, Uint8Array, URL or URL string.",
      );
      setState((previous) => ({ ...EMPTY, status: "error", error, loadId: previous.loadId }));
      optionsRef.current.onError?.(error);
      return () => controller.abort();
    }

    let cancelled = false;
    setState((previous) => ({ ...EMPTY, status: "loading", loadId: previous.loadId }));

    const run = async () => {
      try {
        // `loadDocument` fetches, sniffs the format and parses in one step; the
        // synchronous parsers resolve on the next microtask, which is also what
        // lets the loading state paint before the work starts.
        const parsed = await loadDocument(source, {
          ...optionsRef.current,
          signal: controller.signal,
        });
        if (cancelled || generation.current !== currentGeneration) {
          parsed.dispose();
          return;
        }
        mediaRef.current = parsed.media;
        setState((previous) => ({
          status: "ready",
          model: parsed.model,
          media: parsed.media,
          error: null,
          diagnostics: parsed.model.diagnostics,
          format: parsed.format,
          renderPage: parsed.renderPage ?? null,
          loadId: previous.loadId + 1,
        }));
        optionsRef.current.onLoad?.(parsed.model, parsed.format);
        if (parsed.model.diagnostics.length > 0) {
          optionsRef.current.onDiagnostics?.(parsed.model.diagnostics);
        }
      } catch (error) {
        if (cancelled || generation.current !== currentGeneration) return;
        const normalized = normalizeError(error);
        if (normalized.code === "ABORTED") return;
        setState((previous) => ({ ...EMPTY, status: "error", error: normalized, loadId: previous.loadId }));
        optionsRef.current.onError?.(normalized);
      }
    };

    void run();

    return () => {
      cancelled = true;
      controller.abort();
    };
  }, [source, reloadToken]);

  useEffect(
    () => () => {
      generation.current += 1;
      mediaRef.current?.dispose();
      mediaRef.current = null;
    },
    [],
  );

  return { ...state, reload };
}
