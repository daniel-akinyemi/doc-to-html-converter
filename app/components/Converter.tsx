"use client";

import { useCallback, useState } from "react";
import {
  loadMammoth,
  MAMMOTH_STYLE_MAP,
  plainTextToHtml,
  sanitizeWordHtml,
} from "../lib/docx";
import { UploadZone } from "./UploadZone";
import { ViewToggle, type View } from "./ViewToggle";

type Status =
  | { kind: "idle" }
  | { kind: "loading"; label: string }
  | { kind: "error"; message: string }
  | {
      kind: "ready";
      fileName: string;
      html: string;
      warnings: string[];
      sizeKb: number;
    };

const ACCEPTED_EXT = ".docx";

export function Converter() {
  const [status, setStatus] = useState<Status>({ kind: "idle" });
  const [view, setView] = useState<View>("preview");
  const [copied, setCopied] = useState(false);

  const handleFile = useCallback(async (file: File) => {
    if (!file.name.toLowerCase().endsWith(ACCEPTED_EXT)) {
      setStatus({
        kind: "error",
        message:
          "That doesn’t look like a .docx file. Try exporting from Word or Google Docs as “Microsoft Word (.docx)”.",
      });
      return;
    }
    setStatus({ kind: "loading", label: file.name });
    try {
      const [mammoth, arrayBuffer] = await Promise.all([
        loadMammoth(),
        file.arrayBuffer(),
      ]);
      const result = await mammoth.convertToHtml(
        { arrayBuffer },
        { styleMap: MAMMOTH_STYLE_MAP },
      );
      setStatus({
        kind: "ready",
        fileName: file.name,
        html: result.value,
        warnings: result.messages.map((m: { message: string }) => m.message),
        sizeKb: Math.max(1, Math.round(file.size / 1024)),
      });
      setView("preview");
    } catch (err) {
      setStatus({
        kind: "error",
        message:
          err instanceof Error
            ? err.message
            : "Something went wrong converting that document.",
      });
    }
  }, []);

  const handlePaste = useCallback(
    (payload: { kind: "html" | "text"; value: string }) => {
      const html =
        payload.kind === "html"
          ? sanitizeWordHtml(payload.value)
          : plainTextToHtml(payload.value);
      if (!html.trim()) {
        setStatus({
          kind: "error",
          message:
            "Nothing convertible was found on the clipboard. Try copying the document content again.",
        });
        return;
      }
      setStatus({
        kind: "ready",
        fileName: "Pasted content",
        html,
        warnings: [],
        sizeKb: Math.max(1, Math.round(new Blob([html]).size / 1024)),
      });
      setView("source");
    },
    [],
  );

  const handleClipboardError = useCallback(
    (message: string) => setStatus({ kind: "error", message }),
    [],
  );

  const onCopy = async () => {
    if (status.kind !== "ready") return;
    try {
      await navigator.clipboard.writeText(status.html);
      setCopied(true);
      setTimeout(() => setCopied(false), 1800);
    } catch {
      // clipboard may be unavailable on insecure contexts; ignore silently
    }
  };

  const reset = () => {
    setStatus({ kind: "idle" });
    setView("preview");
  };

  if (status.kind !== "ready") {
    return (
      <UploadZone
        status={status}
        onFile={handleFile}
        onPaste={handlePaste}
        onClipboardError={handleClipboardError}
      />
    );
  }

  return (
    <div className="border border-border bg-background">
      <header className="flex flex-col gap-4 px-5 sm:px-6 py-4 border-b border-border sm:flex-row sm:items-center sm:justify-between">
        <div className="flex items-center gap-3 min-w-0">
          <span className="font-mono text-[10px] tracking-[0.25em] uppercase text-muted-foreground shrink-0">
            § 02
          </span>
          <span className="font-mono text-xs truncate" title={status.fileName}>
            {status.fileName}
          </span>
          <span className="font-mono text-[10px] text-muted-foreground shrink-0">
            {status.sizeKb} KB
          </span>
        </div>
        <div className="flex items-center gap-2">
          <ViewToggle view={view} setView={setView} />
          <button
            type="button"
            onClick={onCopy}
            className="font-mono text-[10px] tracking-[0.25em] uppercase px-3 py-2 border border-foreground bg-foreground text-background hover:bg-foreground/85 transition-colors"
          >
            {copied ? "Copied ✓" : "Copy HTML"}
          </button>
          <button
            type="button"
            onClick={reset}
            className="font-mono text-[10px] tracking-[0.25em] uppercase px-3 py-2 border border-border hover:bg-subtle transition-colors"
          >
            New
          </button>
        </div>
      </header>

      {status.warnings.length > 0 && (
        <details className="border-b border-border bg-subtle/60">
          <summary className="px-6 py-3 cursor-pointer font-mono text-[11px] uppercase tracking-[0.2em] text-muted-foreground select-none">
            {status.warnings.length} note
            {status.warnings.length === 1 ? "" : "s"} from the converter
          </summary>
          <ul className="px-6 pb-4 space-y-1 font-mono text-[11px] text-muted-foreground">
            {status.warnings.slice(0, 12).map((w, i) => (
              <li key={i} className="leading-relaxed">
                — {w}
              </li>
            ))}
            {status.warnings.length > 12 && (
              <li className="leading-relaxed opacity-70">
                + {status.warnings.length - 12} more
              </li>
            )}
          </ul>
        </details>
      )}

      <div>
        {view === "preview" ? (
          <article
            className="prose mx-auto max-w-2xl px-6 sm:px-12 py-12 sm:py-16"
            dangerouslySetInnerHTML={{ __html: status.html }}
          />
        ) : (
          <pre className="px-6 sm:px-12 py-8 text-xs leading-relaxed overflow-auto font-mono text-foreground/90 max-h-[70vh] whitespace-pre-wrap wrap-break-word">
            <code>{status.html}</code>
          </pre>
        )}
      </div>
    </div>
  );
}
