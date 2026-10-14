/**
 * `pdfjs-dist` publishes its workers as ES modules with no type declaration.
 *
 * Importing one is an implementation detail of the main-thread fallback: the
 * namespace object is handed straight to pdf.js, which looks for
 * `globalThis.pdfjsWorker` and uses it instead of spawning a Worker. Both builds
 * are declared because the one that can run depends on the environment (older
 * engines need the `legacy` build; see `detectPdfEngine`).
 */
declare module "pdfjs-dist/build/pdf.worker.min.mjs" {
  /** pdf.js inspects this export when the module is used in-process. */
  export const WorkerMessageHandler: unknown;
  const worker: Record<string, unknown>;
  export default worker;
}

declare module "pdfjs-dist/legacy/build/pdf.worker.min.mjs" {
  /** pdf.js inspects this export when the module is used in-process. */
  export const WorkerMessageHandler: unknown;
  const worker: Record<string, unknown>;
  export default worker;
}
