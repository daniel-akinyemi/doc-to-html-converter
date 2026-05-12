"use client";

import { useCallback, useRef, useState } from "react";
import {
  loadMammoth,
  MAMMOTH_STYLE_MAP,
  plainTextToHtml,
  sanitizeWordHtml,
  SECTION_META,
  splitDocument,
  type SectionId,
  type SplitProduct,
  type SplitResult,
} from "../lib/docx";
import { exportToMatrixify, type ExportReport } from "../lib/matrixify";
import { UploadZone } from "./UploadZone";
import { ViewToggle, type View } from "./ViewToggle";

type Status =
  | { kind: "idle" }
  | { kind: "loading"; label: string }
  | { kind: "error"; message: string }
  | { kind: "ready"; result: SplitResult; sourceLabel: string };

const ACCEPTED_EXT = ".docx";

export function Splitter() {
  const [status, setStatus] = useState<Status>({ kind: "idle" });

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
      const baseName = file.name.replace(/\.docx$/i, "");
      setStatus({
        kind: "ready",
        result: splitDocument(result.value, baseName),
        sourceLabel: file.name,
      });
    } catch (err) {
      setStatus({
        kind: "error",
        message:
          err instanceof Error
            ? err.message
            : "Something went wrong processing that document.",
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
        result: splitDocument(html, "Pasted content"),
        sourceLabel: "Pasted content",
      });
    },
    [],
  );

  const handleClipboardError = useCallback(
    (message: string) => setStatus({ kind: "error", message }),
    [],
  );

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
    <SplitView
      result={status.result}
      sourceLabel={status.sourceLabel}
      onReset={() => setStatus({ kind: "idle" })}
    />
  );
}

type XlsxStatus =
  | { kind: "idle" }
  | { kind: "working"; fileName: string }
  | { kind: "done"; report: ExportReport; fileName: string; url: string }
  | { kind: "error"; message: string };

function downloadBlob(blob: Blob, fileName: string): string {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  return url;
}

function SplitView({
  result,
  sourceLabel,
  onReset,
}: {
  result: SplitResult;
  sourceLabel: string;
  onReset: () => void;
}) {
  const [productIdx, setProductIdx] = useState(0);
  const [xlsx, setXlsx] = useState<XlsxStatus>({ kind: "idle" });
  const xlsxInputRef = useRef<HTMLInputElement>(null);
  const multi = result.products.length > 1;
  const product = result.products[productIdx] ?? result.products[0];

  const runExport = async (file: File) => {
    if (xlsx.kind === "done") URL.revokeObjectURL(xlsx.url);
    setXlsx({ kind: "working", fileName: file.name });
    try {
      const bytes = await file.arrayBuffer();
      const res = await exportToMatrixify(bytes, result.products, file.name);
      const url = downloadBlob(res.blob, res.fileName);
      setXlsx({ kind: "done", report: res.report, fileName: res.fileName, url });
    } catch (err) {
      setXlsx({
        kind: "error",
        message:
          err instanceof Error
            ? err.message
            : "Couldn’t produce the .xlsx file.",
      });
    }
  };

  return (
    <div className="border border-border bg-background">
      <header className="flex flex-col gap-3 px-5 sm:px-6 py-4 border-b border-border sm:flex-row sm:items-center sm:justify-between">
        <div className="flex items-center gap-3 min-w-0">
          <span className="font-mono text-[10px] tracking-[0.25em] uppercase text-muted-foreground shrink-0">
            § 02
          </span>
          <span className="font-mono text-xs truncate" title={sourceLabel}>
            {sourceLabel}
          </span>
          <span className="font-mono text-[10px] text-muted-foreground shrink-0">
            {multi
              ? `${result.products.length} products`
              : "1 document"}
          </span>
        </div>
        <div className="flex items-center gap-2 self-start">
          <input
            ref={xlsxInputRef}
            type="file"
            accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            className="hidden"
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) runExport(f);
              e.target.value = "";
            }}
          />
          <button
            type="button"
            onClick={() => xlsxInputRef.current?.click()}
            disabled={xlsx.kind === "working"}
            title="Fill a customer's Matrixify .xlsx with these sections, matched by Variant SKU"
            className="font-mono text-[10px] tracking-[0.25em] uppercase px-3 py-2 border border-foreground bg-foreground text-background hover:bg-foreground/85 transition-colors disabled:cursor-progress"
          >
            {xlsx.kind === "working" ? "Working…" : "Export .xlsx"}
          </button>
          <button
            type="button"
            onClick={onReset}
            className="font-mono text-[10px] tracking-[0.25em] uppercase px-3 py-2 border border-border hover:bg-subtle transition-colors"
          >
            New
          </button>
        </div>
      </header>

      {xlsx.kind !== "idle" && <ExportPanel status={xlsx} />}

      {multi && (
        <div className="border-b border-border px-5 sm:px-6 py-3">
          <div className="font-mono text-[10px] tracking-[0.25em] uppercase text-muted-foreground mb-2">
            Product · {result.products.length}
          </div>
          {result.products.length <= 8 ? (
            <div className="flex flex-wrap gap-2">
              {result.products.map((p, i) => (
                <button
                  key={i}
                  type="button"
                  onClick={() => setProductIdx(i)}
                  aria-pressed={i === productIdx}
                  className={`font-mono text-[11px] px-3 py-1.5 border transition-colors max-w-full truncate ${
                    i === productIdx
                      ? "border-foreground bg-foreground text-background"
                      : "border-border hover:bg-subtle"
                  }`}
                  title={p.title}
                >
                  {p.title}
                </button>
              ))}
            </div>
          ) : (
            <select
              value={productIdx}
              onChange={(e) => setProductIdx(Number(e.target.value))}
              className="w-full max-w-lg font-mono text-xs border border-border bg-background px-3 py-2"
            >
              {result.products.map((p, i) => (
                <option key={i} value={i}>
                  {p.title}
                </option>
              ))}
            </select>
          )}
        </div>
      )}

      <SectionTabs key={productIdx} product={product} />
    </div>
  );
}

function SectionTabs({ product }: { product: SplitProduct }) {
  const firstNonEmpty =
    SECTION_META.find((s) => product.sections[s.id]?.trim())?.id ?? "overview";
  const [active, setActive] = useState<SectionId>(firstNonEmpty);
  const [view, setView] = useState<View>("preview");
  const [copied, setCopied] = useState(false);
  const html = product.sections[active] || "";
  const activeLabel = SECTION_META.find((s) => s.id === active)?.label ?? "";

  const onCopy = async () => {
    if (!html) return;
    try {
      await navigator.clipboard.writeText(html);
      setCopied(true);
      setTimeout(() => setCopied(false), 1800);
    } catch {
      // ignore
    }
  };

  return (
    <div>
      <div className="px-5 sm:px-6 pt-5">
        <div className="font-mono text-[10px] tracking-[0.25em] uppercase text-muted-foreground">
          Document
        </div>
        <div className="font-serif text-2xl sm:text-3xl leading-tight tracking-tight mt-1">
          {product.title}
        </div>
      </div>

      <div className="mt-5 px-5 sm:px-6 flex flex-wrap gap-x-1 border-b border-border">
        {SECTION_META.map(({ id, label }) => {
          const empty = !product.sections[id]?.trim();
          return (
            <button
              key={id}
              type="button"
              onClick={() => {
                setActive(id);
                setView("preview");
                setCopied(false);
              }}
              aria-pressed={id === active}
              className={`font-mono text-[10px] tracking-[0.2em] uppercase px-3 py-2.5 border-b-2 -mb-px transition-colors ${
                id === active
                  ? "border-foreground text-foreground"
                  : "border-transparent text-muted-foreground hover:text-foreground"
              } ${empty ? "opacity-40" : ""}`}
            >
              {label}
            </button>
          );
        })}
      </div>

      <div className="flex items-center justify-between gap-2 px-5 sm:px-6 py-3 border-b border-border bg-subtle/40">
        <span className="font-mono text-[10px] tracking-[0.25em] uppercase text-muted-foreground">
          {activeLabel}
        </span>
        <div className="flex items-center gap-2">
          <ViewToggle view={view} setView={setView} />
          <button
            type="button"
            onClick={onCopy}
            disabled={!html}
            className="font-mono text-[10px] tracking-[0.25em] uppercase px-3 py-2 border border-foreground bg-foreground text-background hover:bg-foreground/85 transition-colors disabled:opacity-40 disabled:cursor-not-allowed"
          >
            {copied ? "Copied ✓" : "Copy HTML"}
          </button>
        </div>
      </div>

      <div>
        {!html ? (
          <div className="px-6 sm:px-12 py-16 sm:py-20 text-center font-mono text-xs text-muted-foreground">
            Not found in this document.
          </div>
        ) : view === "preview" ? (
          <article
            className="prose mx-auto max-w-2xl px-6 sm:px-12 py-12 sm:py-16"
            dangerouslySetInnerHTML={{ __html: html }}
          />
        ) : (
          <pre className="px-6 sm:px-12 py-8 text-xs leading-relaxed overflow-auto font-mono text-foreground/90 max-h-[70vh] whitespace-pre-wrap wrap-break-word">
            <code>{html}</code>
          </pre>
        )}
      </div>
    </div>
  );
}

function ExportPanel({
  status,
}: {
  status: Exclude<XlsxStatus, { kind: "idle" }>;
}) {
  if (status.kind === "working") {
    return (
      <div className="border-b border-border bg-subtle px-5 sm:px-6 py-3 font-mono text-[11px] tracking-[0.2em] uppercase text-muted-foreground">
        Updating <span className="text-foreground">{status.fileName}</span> …
      </div>
    );
  }
  if (status.kind === "error") {
    return (
      <div className="border-b border-border bg-red-50 dark:bg-red-950/30 px-5 sm:px-6 py-3 font-mono text-[11px] leading-relaxed text-red-700 dark:text-red-400">
        {status.message}
      </div>
    );
  }

  const r = status.report;
  return (
    <div className="border-b border-border bg-subtle/60 px-5 sm:px-6 py-4 space-y-2 font-mono text-[11px] leading-relaxed text-muted-foreground">
      <div className="text-foreground tracking-[0.2em] uppercase">
        ✓ {status.fileName} downloaded
      </div>
      <div>
        Sheet “{r.sheetName}” · {r.rowsMatched}/{r.rowsScanned} SKU row
        {r.rowsScanned === 1 ? "" : "s"} matched · {r.cellsWritten} cell
        {r.cellsWritten === 1 ? "" : "s"} written
      </div>
      <div>
        Columns filled: {r.mappedColumns.map((m) => m.section).join(", ") || "—"}
        {r.missingColumns.length > 0 && (
          <span className="text-red-600 dark:text-red-400">
            {" "}
            · not in the sheet: {r.missingColumns.join(", ")}
          </span>
        )}
      </div>
      {r.unmatchedProducts.length > 0 && (
        <details>
          <summary className="cursor-pointer select-none">
            {r.unmatchedProducts.length} product
            {r.unmatchedProducts.length === 1 ? "" : "s"} with no matching row
          </summary>
          <ul className="mt-1 space-y-0.5">
            {r.unmatchedProducts.slice(0, 30).map((p, i) => (
              <li key={i}>
                — {p.title} <span className="opacity-70">[{p.key}]</span>
              </li>
            ))}
            {r.unmatchedProducts.length > 30 && (
              <li className="opacity-70">
                + {r.unmatchedProducts.length - 30} more
              </li>
            )}
          </ul>
        </details>
      )}
      {r.productsWithoutKey.length > 0 && (
        <details>
          <summary className="cursor-pointer select-none">
            {r.productsWithoutKey.length} product
            {r.productsWithoutKey.length === 1 ? "" : "s"} with no “ID …” / “SKU
            …” line in the doc
          </summary>
          <ul className="mt-1 space-y-0.5">
            {r.productsWithoutKey.slice(0, 30).map((t, i) => (
              <li key={i}>— {t}</li>
            ))}
            {r.productsWithoutKey.length > 30 && (
              <li className="opacity-70">
                + {r.productsWithoutKey.length - 30} more
              </li>
            )}
          </ul>
        </details>
      )}
      <a
        href={status.url}
        download={status.fileName}
        className="inline-block text-foreground underline underline-offset-4 decoration-foreground/30 hover:decoration-foreground"
      >
        Download again
      </a>
    </div>
  );
}
