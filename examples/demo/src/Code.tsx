/**
 * A code block with light syntax colouring and a copy button. The tokeniser only
 * knows what the demo's own TSX and shell snippets need.
 */

import { useEffect, useState, type ReactNode } from "react";

const TOKEN =
  /(\/\/[^\n]*|\{\/\*[\s\S]*?\*\/\}|"[^"\n]*"|`[^`]*`|<\/?[A-Za-z][\w.]*|\/?>|\b(?:import|from|export|function|return|const|default)\b|\b[a-zA-Z]+(?==)|\b(?:npm|pnpm|yarn)\b)/g;

function kind(token: string): string {
  if (token.startsWith("//") || token.startsWith("{/*")) return "c";
  if (token.startsWith('"') || token.startsWith("`")) return "s";
  if (token.startsWith("<") || token.endsWith(">")) return "t";
  if (/^(npm|pnpm|yarn)$/.test(token)) return "k";
  if (/^(import|from|export|function|return|const|default)$/.test(token)) return "k";
  return "a";
}

function highlight(code: string): ReactNode[] {
  const out: ReactNode[] = [];
  let last = 0;
  for (const match of code.matchAll(TOKEN)) {
    const index = match.index ?? 0;
    if (index > last) out.push(code.slice(last, index));
    out.push(
      <span key={index} className={`tok-${kind(match[0])}`}>
        {match[0]}
      </span>,
    );
    last = index + match[0].length;
  }
  if (last < code.length) out.push(code.slice(last));
  return out;
}

export function useCopy(): [boolean, (text: string) => void] {
  const [copied, setCopied] = useState(false);
  useEffect(() => {
    if (!copied) return undefined;
    const timer = setTimeout(() => setCopied(false), 1600);
    return () => clearTimeout(timer);
  }, [copied]);
  const copy = (text: string) => {
    try {
      void navigator.clipboard?.writeText(text).then(
        () => setCopied(true),
        () => undefined,
      );
    } catch {
      // Clipboard access can be refused; the text stays selectable.
    }
  };
  return [copied, copy];
}

export function Code({ code, title, compact }: { code: string; title?: string; compact?: boolean }) {
  const [copied, copy] = useCopy();
  return (
    <div className={`code${compact ? " code--compact" : ""}`}>
      {title || !compact ? (
        <div className="code__head">
          <span>{title}</span>
          <button type="button" className="code__copy" onClick={() => copy(code)}>
            {copied ? "Copied" : "Copy"}
          </button>
        </div>
      ) : null}
      <pre>
        <code>{highlight(code)}</code>
      </pre>
    </div>
  );
}
