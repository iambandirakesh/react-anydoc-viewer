import type { ViewerLabels } from "../types/public";

export const DEFAULT_LABELS: ViewerLabels = {
  previous: "Previous slide",
  next: "Next slide",
  first: "First slide",
  last: "Last slide",
  zoomIn: "Zoom in",
  zoomOut: "Zoom out",
  fitWidth: "Fit to width",
  fitPage: "Fit to page",
  actualSize: "Actual size (100%)",
  fullscreen: "Enter fullscreen",
  exitFullscreen: "Exit fullscreen",
  present: "Present from this slide",
  exitPresentation: "Exit presentation",
  presentationHint: "Space to advance · Esc to exit",
  toggleThumbnails: "Toggle page thumbnails",
  thumbnails: "Page thumbnails",
  slideOf: (current, total) => `${current} / ${total}`,
  loading: "Loading document…",
  errorTitle: "This document could not be displayed",
  retry: "Try again",
  diagnostics: "Diagnostics",
  diagnosticsClose: "Close diagnostics",
  diagnosticsEmpty: "No issues were reported for this document.",
  noMeasure: "Waiting for the container to have a usable size…",
  emptySlide: "This page is empty",
  renderError: "This page failed to render",
  viewerRegion: "Document viewer",
};

export function mergeLabels(overrides: Partial<ViewerLabels> | undefined): ViewerLabels {
  return overrides ? { ...DEFAULT_LABELS, ...overrides } : DEFAULT_LABELS;
}
