import type { ReactNode } from "react";
import type { DocumentDiagnostic } from "../types/public";
import type { ViewerLabels } from "../types/public";

export interface DiagnosticsPanelProps {
  diagnostics: DocumentDiagnostic[];
  labels: ViewerLabels;
  onClose?: () => void;
}

const SEVERITY_ORDER: Record<DocumentDiagnostic["severity"], number> = {
  error: 0,
  warning: 1,
  info: 2,
};

export function DiagnosticsPanel({ diagnostics, labels, onClose }: DiagnosticsPanelProps): ReactNode {
  const sorted = [...diagnostics].sort(
    (a, b) => SEVERITY_ORDER[a.severity] - SEVERITY_ORDER[b.severity],
  );

  return (
    <div className="radv-diagnostics" role="log" aria-label={labels.diagnostics} data-testid="rv-diagnostics">
      <div className="radv-diagnostics-header">
        <strong>{labels.diagnostics}</strong>
        <span className="radv-diagnostics-count">{sorted.length}</span>
        {onClose ? (
          <button type="button" className="radv-btn radv-btn-small" onClick={onClose} aria-label={labels.diagnosticsClose}>
            <span aria-hidden="true">×</span>
          </button>
        ) : null}
      </div>
      {sorted.length === 0 ? (
        <p className="radv-diagnostics-empty">{labels.diagnosticsEmpty}</p>
      ) : (
        <ul className="radv-diagnostics-list">
          {sorted.map((diagnostic, index) => (
            <li
              key={`${diagnostic.code}-${index}`}
              className={`radv-diagnostic radv-diagnostic-${diagnostic.severity}`}
              data-code={diagnostic.code}
            >
              <span className="radv-diagnostic-code">{diagnostic.code}</span>
              <span className="radv-diagnostic-message">{diagnostic.message}</span>
              {diagnostic.slideIndex !== undefined ? (
                <span className="radv-diagnostic-slide">
                  {labels.slideOf(diagnostic.slideIndex + 1, diagnostic.slideIndex + 1)}
                </span>
              ) : null}
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
