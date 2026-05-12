"use client";

import { useCallback, useEffect, useRef, useState } from "react";

export type UploadZoneStatus =
  | { kind: "idle" }
  | { kind: "loading"; label: string }
  | { kind: "error"; message: string };

type PastePayload = { kind: "html" | "text"; value: string };

type Props = {
  status: UploadZoneStatus;
  onFile: (file: File) => void;
  onPaste: (payload: PastePayload) => void;
  onClipboardError: (message: string) => void;
};

// Drop zone + "paste from clipboard" card + a document-level ⌘V listener.
// Shared by the converter and the splitter.
export function UploadZone({ status, onFile, onPaste, onClipboardError }: Props) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragOver, setDragOver] = useState(false);
  const isLoading = status.kind === "loading";

  const onInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) onFile(file);
    e.target.value = "";
  };

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files?.[0];
    if (file) onFile(file);
  };

  const readClipboard = useCallback(async () => {
    try {
      if (navigator.clipboard && "read" in navigator.clipboard) {
        const items = await navigator.clipboard.read();
        for (const item of items) {
          if (item.types.includes("text/html")) {
            const blob = await item.getType("text/html");
            onPaste({ kind: "html", value: await blob.text() });
            return;
          }
        }
        for (const item of items) {
          if (item.types.includes("text/plain")) {
            const blob = await item.getType("text/plain");
            onPaste({ kind: "text", value: await blob.text() });
            return;
          }
        }
      }
      const text = await navigator.clipboard.readText();
      if (text) onPaste({ kind: "text", value: text });
      else
        onClipboardError(
          "Your clipboard is empty. Copy something from your document, then try again.",
        );
    } catch {
      onClipboardError(
        "Couldn’t read your clipboard. Try pressing ⌘V (Ctrl+V) anywhere on the page instead.",
      );
    }
  }, [onPaste, onClipboardError]);

  useEffect(() => {
    if (status.kind === "loading") return;
    const handler = (e: ClipboardEvent) => {
      const target = e.target as HTMLElement | null;
      if (
        target &&
        (target.tagName === "INPUT" ||
          target.tagName === "TEXTAREA" ||
          target.isContentEditable)
      ) {
        return;
      }
      const html = e.clipboardData?.getData("text/html");
      const text = e.clipboardData?.getData("text/plain");
      if (html) {
        e.preventDefault();
        onPaste({ kind: "html", value: html });
      } else if (text) {
        e.preventDefault();
        onPaste({ kind: "text", value: text });
      }
    };
    document.addEventListener("paste", handler);
    return () => document.removeEventListener("paste", handler);
  }, [status.kind, onPaste]);

  return (
    <div className="space-y-3">
      <div className="grid gap-3 sm:grid-cols-2">
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
            className="flex h-full w-full flex-col items-center justify-center px-6 py-16 sm:py-20 text-center disabled:cursor-progress"
          >
            <div className="font-mono text-[10px] tracking-[0.25em] uppercase text-muted-foreground mb-6">
              § File
            </div>
            <div className="font-serif text-3xl sm:text-4xl leading-tight tracking-tight">
              Drop a <em>.docx</em>
            </div>
            <div className="mt-3 text-sm text-muted-foreground">
              or{" "}
              <span className="underline underline-offset-4 decoration-foreground/30 group-hover:decoration-foreground transition-colors">
                choose from your computer
              </span>
            </div>
          </button>
        </div>

        <div className="group relative border border-dashed border-border hover:border-foreground/40 transition-colors">
          <button
            type="button"
            onClick={readClipboard}
            disabled={isLoading}
            className="flex h-full w-full flex-col items-center justify-center px-6 py-16 sm:py-20 text-center disabled:cursor-progress"
          >
            <div className="font-mono text-[10px] tracking-[0.25em] uppercase text-muted-foreground mb-6">
              § Paste
            </div>
            <div className="font-serif text-3xl sm:text-4xl leading-tight tracking-tight">
              Paste <em>content</em>
            </div>
            <div className="mt-3 text-sm text-muted-foreground">
              <span className="underline underline-offset-4 decoration-foreground/30 group-hover:decoration-foreground transition-colors">
                from your clipboard
              </span>{" "}
              — or just ⌘V
            </div>
          </button>
        </div>
      </div>

      {status.kind === "loading" && (
        <div
          role="status"
          className="border border-border bg-subtle px-5 py-3 font-mono text-[11px] tracking-[0.2em] uppercase text-muted-foreground"
        >
          Working… <span className="text-foreground">{status.label}</span>
        </div>
      )}

      {status.kind === "error" && (
        <div
          role="alert"
          className="border border-red-500/40 bg-red-50 dark:bg-red-950/30 px-5 py-3 font-mono text-[11px] leading-relaxed text-red-700 dark:text-red-400"
        >
          {status.message}
        </div>
      )}
    </div>
  );
}
