"use client";

import { useCallback, useEffect, useRef, useState } from "react";

type MammothOptions = {
  styleMap?: string | string[];
};

type MammothLib = {
  convertToHtml: (
    input: { arrayBuffer: ArrayBuffer },
    options?: MammothOptions,
  ) => Promise<{ value: string; messages: { message: string }[] }>;
};

declare global {
  interface Window {
    mammoth?: MammothLib;
  }
}

let mammothPromise: Promise<MammothLib> | null = null;

function loadMammoth(): Promise<MammothLib> {
  if (typeof window === "undefined") {
    return Promise.reject(
      new Error("Converter is only available in the browser."),
    );
  }
  if (window.mammoth) return Promise.resolve(window.mammoth);
  if (mammothPromise) return mammothPromise;
  mammothPromise = new Promise<MammothLib>((resolve, reject) => {
    const existing = document.querySelector<HTMLScriptElement>(
      'script[data-lib="mammoth"]',
    );
    const onReady = () => {
      if (window.mammoth) resolve(window.mammoth);
      else reject(new Error("Converter library loaded but is unavailable."));
    };
    if (existing) {
      existing.addEventListener("load", onReady, { once: true });
      existing.addEventListener(
        "error",
        () => reject(new Error("Failed to load converter library.")),
        { once: true },
      );
      return;
    }
    const script = document.createElement("script");
    script.src = "/mammoth.browser.min.js";
    script.async = true;
    script.dataset.lib = "mammoth";
    script.addEventListener("load", onReady, { once: true });
    script.addEventListener(
      "error",
      () => {
        mammothPromise = null;
        reject(new Error("Failed to load converter library."));
      },
      { once: true },
    );
    document.head.appendChild(script);
  });
  return mammothPromise;
}

// Extend mammoth's defaults so character formatting that it ignores by default
// (underline, strikethrough, etc.) survives into the HTML output. h1/h2/Title
// are clamped to h3 — h3 is the largest heading we let through.
const MAMMOTH_STYLE_MAP: string[] = [
  "p[style-name='Heading 1'] => h3:fresh",
  "p[style-name='Heading 2'] => h3:fresh",
  "p[style-name='Title'] => h3:fresh",
  "u => u",
  "strike => s",
  "r[style-name='Strong'] => strong",
  "r[style-name='Emphasis'] => em",
  "r[style-name='Subtle Emphasis'] => em",
  "r[style-name='Intense Emphasis'] => strong > em",
  "r[style-name='Book Title'] => em",
  "r[style-name='Code'] => code",
  "p[style-name='Quote'] => blockquote > p:fresh",
  "p[style-name='Intense Quote'] => blockquote > p:fresh",
  "p[style-name='Caption'] => p.caption:fresh",
];

// ---------- HTML sanitiser used for clipboard paste content ----------

const BLOCK_TAGS = new Set([
  "p",
  "h1",
  "h2",
  "h3",
  "h4",
  "h5",
  "h6",
  "li",
  "td",
  "th",
  "div",
  "blockquote",
  "ul",
  "ol",
  "table",
  "tr",
]);

const STYLE_ALLOWED = new Set([
  "p",
  "h1",
  "h2",
  "h3",
  "h4",
  "h5",
  "h6",
  "li",
  "td",
  "th",
  "div",
  "blockquote",
  "ul",
  "ol",
  "table",
  "tr",
  "span",
  "strong",
  "em",
  "u",
  "s",
  "mark",
  "sup",
  "sub",
  "a",
  "code",
]);

const KEEP_ATTRS: Record<string, Set<string>> = {
  a: new Set(["href"]),
  img: new Set(["src", "alt"]),
  td: new Set(["colspan", "rowspan"]),
  th: new Set(["colspan", "rowspan"]),
};

function unwrap(el: Element) {
  const parent = el.parentNode;
  if (!parent) return;
  while (el.firstChild) parent.insertBefore(el.firstChild, el);
  parent.removeChild(el);
}

function parseInlineStyle(styleAttr: string): Record<string, string> {
  const map: Record<string, string> = {};
  for (const decl of styleAttr.split(";")) {
    const idx = decl.indexOf(":");
    if (idx < 0) continue;
    const prop = decl.slice(0, idx).trim().toLowerCase();
    const val = decl.slice(idx + 1).trim();
    if (prop && val) map[prop] = val;
  }
  return map;
}

function isBoldWeight(value: string): boolean {
  const v = value.trim().toLowerCase();
  if (v === "bold" || v === "bolder") return true;
  const n = parseInt(v, 10);
  return Number.isFinite(n) && n >= 600;
}

function isDefaultColor(value: string): boolean {
  const v = value.replace(/\s+/g, "").toLowerCase();
  return (
    v === "" ||
    v === "inherit" ||
    v === "initial" ||
    v === "currentcolor" ||
    v === "transparent" ||
    v === "#000" ||
    v === "#000000" ||
    v === "rgb(0,0,0)" ||
    v === "rgba(0,0,0,1)" ||
    v === "black" ||
    v === "windowtext"
  );
}

function isDefaultBg(value: string): boolean {
  const v = value.replace(/\s+/g, "").toLowerCase();
  return (
    v === "" ||
    v === "inherit" ||
    v === "initial" ||
    v === "transparent" ||
    v === "rgba(0,0,0,0)" ||
    v === "#fff" ||
    v === "#ffffff" ||
    v === "rgb(255,255,255)" ||
    v === "rgba(255,255,255,1)" ||
    v === "white"
  );
}

function hasBlockChildren(el: Element): boolean {
  for (const child of Array.from(el.children)) {
    if (BLOCK_TAGS.has(child.tagName.toLowerCase())) return true;
  }
  return false;
}

// Convert messy Word / Google Docs paste HTML into clean semantic markup,
// preserving meaningful formatting (bold, italic, underline, strikethrough,
// super/subscript, alignment, text/background colour).
function sanitizeWordHtml(rawHtml: string): string {
  if (typeof window === "undefined") return rawHtml;
  const cleaned = rawHtml.replace(/<!--\[if[\s\S]*?<!\[endif\]-->/g, "");
  const doc = new DOMParser().parseFromString(cleaned, "text/html");
  const root = doc.body;

  // 1. Promote MsoHeading / MsoTitle paragraphs to real headings (clamped to h3).
  root.querySelectorAll("p[class]").forEach((p) => {
    const cls = p.getAttribute("class") || "";
    const m = /\bMsoHeading(\d)\b/i.exec(cls);
    if (m) {
      const level = Math.min(6, Math.max(3, parseInt(m[1], 10)));
      const h = doc.createElement(`h${level}`);
      const style = p.getAttribute("style");
      if (style) h.setAttribute("style", style);
      while (p.firstChild) h.appendChild(p.firstChild);
      p.replaceWith(h);
    } else if (/\bMsoTitle\b/i.test(cls)) {
      const h = doc.createElement("h3");
      const style = p.getAttribute("style");
      if (style) h.setAttribute("style", style);
      while (p.firstChild) h.appendChild(p.firstChild);
      p.replaceWith(h);
    }
  });

  // 2. Unwrap Google Docs' <b style="font-weight:normal"> document wrapper.
  root.querySelectorAll("b[style]").forEach((b) => {
    const style = (b.getAttribute("style") || "")
      .replace(/\s+/g, "")
      .toLowerCase();
    if (
      style.includes("font-weight:normal") ||
      style.includes("font-weight:400")
    ) {
      unwrap(b);
    }
  });

  // 3. Drop non-content elements outright.
  root
    .querySelectorAll("style, script, meta, link, title, head, noscript")
    .forEach((el) => el.remove());

  // 4. Convert legacy presentational tags before inspecting styles.
  for (const el of Array.from(root.querySelectorAll("*"))) {
    const tag = el.tagName.toLowerCase();
    if (tag.includes(":") || tag === "xml") {
      el.remove();
      continue;
    }
    if (tag === "b") {
      const strong = doc.createElement("strong");
      const style = el.getAttribute("style");
      if (style) strong.setAttribute("style", style);
      while (el.firstChild) strong.appendChild(el.firstChild);
      el.replaceWith(strong);
      continue;
    }
    if (tag === "i") {
      const em = doc.createElement("em");
      const style = el.getAttribute("style");
      if (style) em.setAttribute("style", style);
      while (el.firstChild) em.appendChild(el.firstChild);
      el.replaceWith(em);
      continue;
    }
    if (tag === "font") {
      const color = el.getAttribute("color");
      if (color) {
        const span = doc.createElement("span");
        span.setAttribute("style", `color: ${color}`);
        while (el.firstChild) span.appendChild(el.firstChild);
        el.replaceWith(span);
      } else {
        unwrap(el);
      }
      continue;
    }
    if (tag === "center") {
      unwrap(el);
      continue;
    }
  }

  // 5. For every element, derive semantic wrappers + preserved styles from CSS.
  for (const el of Array.from(root.querySelectorAll("*"))) {
    const tag = el.tagName.toLowerCase();
    const styleAttr = el.getAttribute("style");
    const styles = styleAttr ? parseInlineStyle(styleAttr) : {};

    const wrappers: string[] = [];
    if (styleAttr) {
      const fw = styles["font-weight"];
      if (fw && isBoldWeight(fw) && tag !== "strong") wrappers.push("strong");

      const fs = styles["font-style"];
      if ((fs === "italic" || fs === "oblique") && tag !== "em")
        wrappers.push("em");

      const td = styles["text-decoration"] || styles["text-decoration-line"];
      if (td) {
        if (td.includes("underline") && tag !== "u") wrappers.push("u");
        if (td.includes("line-through") && tag !== "s") wrappers.push("s");
      }

      const va = styles["vertical-align"];
      if (va === "super" && tag !== "sup") wrappers.push("sup");
      if (va === "sub" && tag !== "sub") wrappers.push("sub");
    }

    const preserved: string[] = [];
    if (styleAttr) {
      const color = styles["color"];
      if (color && !isDefaultColor(color)) preserved.push(`color: ${color}`);
      const bg = styles["background-color"];
      if (bg && !isDefaultBg(bg)) preserved.push(`background-color: ${bg}`);
    }

    // Strip all attributes; we'll re-add what we want to keep.
    const baseKeep = KEEP_ATTRS[tag] || new Set<string>();
    for (const attr of Array.from(el.attributes)) {
      if (!baseKeep.has(attr.name)) el.removeAttribute(attr.name);
    }
    if (preserved.length > 0 && STYLE_ALLOWED.has(tag)) {
      el.setAttribute("style", preserved.join("; "));
    }

    // Apply semantic wrappers (innermost first) when safe.
    if (wrappers.length > 0 && !hasBlockChildren(el) && el.childNodes.length) {
      let current: Node[] = Array.from(el.childNodes);
      for (let i = wrappers.length - 1; i >= 0; i--) {
        const wrap = doc.createElement(wrappers[i]);
        for (const node of current) wrap.appendChild(node);
        current = [wrap];
      }
      for (const node of current) el.appendChild(node);
    }
  }

  // 6. Unwrap span / div with no remaining attributes (just clutter).
  root.querySelectorAll("span, div").forEach((el) => {
    if (el.attributes.length === 0) unwrap(el);
  });

  // 7. Drop empty paragraphs.
  root.querySelectorAll("p").forEach((p) => {
    if (!p.textContent?.trim() && !p.querySelector("img")) p.remove();
  });

  // 8. Clamp headings to h3 maximum.
  root.querySelectorAll("h1, h2").forEach((h) => {
    const replacement = doc.createElement("h3");
    for (const attr of Array.from(h.attributes)) {
      replacement.setAttribute(attr.name, attr.value);
    }
    while (h.firstChild) replacement.appendChild(h.firstChild);
    h.replaceWith(replacement);
  });

  return root.innerHTML.trim();
}

function plainTextToHtml(text: string): string {
  const escape = (s: string) =>
    s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  return text
    .replace(/\r\n/g, "\n")
    .split(/\n{2,}/)
    .map((b) => b.trim())
    .filter(Boolean)
    .map((b) => `<p>${escape(b).replace(/\n/g, "<br />")}</p>`)
    .join("\n");
}

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

  const acceptPasted = useCallback(
    (
      payload: { kind: "html" | "text"; value: string },
      sourceLabel = "Pasted content",
    ) => {
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
        fileName: sourceLabel,
        html,
        warnings: [],
        sizeKb: Math.max(1, Math.round(new Blob([html]).size / 1024)),
      });
      setView("source");
    },
    [],
  );

  const handleClipboardRead = useCallback(async () => {
    setStatus({ kind: "loading", label: "Reading clipboard…" });
    try {
      if (navigator.clipboard && "read" in navigator.clipboard) {
        const items = await navigator.clipboard.read();
        for (const item of items) {
          if (item.types.includes("text/html")) {
            const blob = await item.getType("text/html");
            acceptPasted({ kind: "html", value: await blob.text() });
            return;
          }
        }
        for (const item of items) {
          if (item.types.includes("text/plain")) {
            const blob = await item.getType("text/plain");
            acceptPasted({ kind: "text", value: await blob.text() });
            return;
          }
        }
      }
      const text = await navigator.clipboard.readText();
      if (text) acceptPasted({ kind: "text", value: text });
      else
        setStatus({
          kind: "error",
          message:
            "Your clipboard is empty. Copy something from your document, then try again.",
        });
    } catch {
      setStatus({
        kind: "error",
        message:
          "Couldn’t read your clipboard. Try pressing ⌘V (Ctrl+V) anywhere on the page instead.",
      });
    }
  }, [acceptPasted]);

  useEffect(() => {
    if (status.kind !== "idle" && status.kind !== "error") return;
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
        acceptPasted({ kind: "html", value: html });
      } else if (text) {
        e.preventDefault();
        acceptPasted({ kind: "text", value: text });
      }
    };
    document.addEventListener("paste", handler);
    return () => document.removeEventListener("paste", handler);
  }, [status.kind, acceptPasted]);

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
              onClick={handleClipboardRead}
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

        {isLoading && (
          <div
            role="status"
            className="border border-border bg-subtle px-5 py-3 font-mono text-[11px] tracking-[0.2em] uppercase text-muted-foreground"
          >
            Converting… <span className="text-foreground">{status.label}</span>
          </div>
        )}

        {isError && (
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
