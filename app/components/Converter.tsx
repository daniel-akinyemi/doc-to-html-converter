"use client";

import { useCallback, useRef, useState } from "react";

type Status =
  | { kind: "idle" }
  | { kind: "loading"; fileName: string }
  | { kind: "error"; message: string }
  | {
      kind: "ready";
      fileName: string;
      html: string;
      warnings: string[];
      sizeKb: number;
    };

type View = "preview" | "source";

const ACCEPTED_EXT = ".docx";

export function Converter() {
  const [status, setStatus] = useState<Status>({ kind: "idle" });
  const [view, setView] = useState<View>("preview");
  const [copied, setCopied] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const handleFile = useCallback(async (file: File) => {
    if (!file.name.toLowerCase().endsWith(ACCEPTED_EXT)) {
      setStatus({
        kind: "error",
        message:
          "That doesn’t look like a .docx file. Try exporting from Word or Google Docs as “Microsoft Word (.docx)”.",
      });
      return;
    }
    setStatus({ kind: "loading", fileName: file.name });
    try {
      const [{ default: mammoth }, arrayBuffer] = await Promise.all([
        import("mammoth"),
        file.arrayBuffer(),
      ]);
      const result = await mammoth.convertToHtml({ arrayBuffer });
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

  const onInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) handleFile(file);
    e.target.value = "";
  };

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files?.[0];
    if (file) handleFile(file);
  };

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
    const isLoading = status.kind === "loading";
    const isError = status.kind === "error";
    return (
      <div
        onDragOver={(e) => {
          e.preventDefault();
          if (!isLoading) setDragOver(true);
        }}
        onDragLeave={() => setDragOver(false)}
        onDrop={onDrop}
        className={`group relative border border-dashed transition-colors ${
          dragOver
            ? "border-foreground bg-subtle"
            : "border-border hover:border-foreground/40"
        }`}
      >
        <input
          ref={inputRef}
          type="file"
          accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          className="hidden"
          onChange={onInputChange}
        />
        <button
          type="button"
          onClick={() => inputRef.current?.click()}
          disabled={isLoading}
          className="block w-full px-6 py-20 sm:py-24 text-center disabled:cursor-progress"
        >
          <div className="font-mono text-[10px] tracking-[0.25em] uppercase text-muted-foreground mb-8">
            § 01 — Upload
          </div>
          <div className="font-serif text-4xl sm:text-5xl leading-tight tracking-tight">
            {isLoading ? (
              <em>Converting…</em>
            ) : (
              <>
                Drop a <em>.docx</em> here
              </>
            )}
          </div>
          <div className="mt-4 text-sm text-muted-foreground">
            {isLoading ? (
              <span className="font-mono text-xs">{status.fileName}</span>
            ) : (
              <>
                or{" "}
                <span className="underline underline-offset-4 decoration-foreground/30 group-hover:decoration-foreground transition-colors">
                  choose from your computer
                </span>
              </>
            )}
          </div>
          {isError && (
            <div className="mx-auto mt-8 max-w-md text-xs font-mono text-red-700 dark:text-red-400 leading-relaxed">
              {status.message}
            </div>
          )}
        </button>
      </div>
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
          <pre className="px-6 sm:px-12 py-8 text-xs leading-relaxed overflow-auto font-mono text-foreground/90 max-h-[70vh] whitespace-pre-wrap break-words">
            <code>{status.html}</code>
          </pre>
        )}
      </div>
    </div>
  );
}

function ViewToggle({
  view,
  setView,
}: {
  view: View;
  setView: (v: View) => void;
}) {
  const items: { id: View; label: string }[] = [
    { id: "preview", label: "Preview" },
    { id: "source", label: "HTML" },
  ];
  return (
    <div className="flex border border-border">
      {items.map((it) => (
        <button
          key={it.id}
          type="button"
          onClick={() => setView(it.id)}
          aria-pressed={view === it.id}
          className={`font-mono text-[10px] tracking-[0.25em] uppercase px-3 py-2 transition-colors ${
            view === it.id
              ? "bg-foreground text-background"
              : "hover:bg-subtle"
          }`}
        >
          {it.label}
        </button>
      ))}
    </div>
  );
}
