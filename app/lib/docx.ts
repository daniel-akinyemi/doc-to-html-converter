// Shared .docx -> semantic-HTML conversion, clipboard sanitising, and
// document splitting. Everything here is safe to import anywhere; the
// browser-only functions guard on `typeof window` so they no-op during SSR.

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

export function loadMammoth(): Promise<MammothLib> {
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

// Extend mammoth's defaults so character formatting it ignores by default
// (underline, strikethrough, etc.) survives. h1/h2/Title are clamped to h3 —
// h3 is the largest heading we let through.
export const MAMMOTH_STYLE_MAP: string[] = [
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
  "p", "h1", "h2", "h3", "h4", "h5", "h6", "li", "td", "th", "div",
  "blockquote", "ul", "ol", "table", "tr",
]);

const STYLE_ALLOWED = new Set([
  "p", "h1", "h2", "h3", "h4", "h5", "h6", "li", "td", "th", "div",
  "blockquote", "ul", "ol", "table", "tr", "span", "strong", "em", "u",
  "s", "mark", "sup", "sub", "a", "code",
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
    v === "" || v === "inherit" || v === "initial" || v === "currentcolor" ||
    v === "transparent" || v === "#000" || v === "#000000" ||
    v === "rgb(0,0,0)" || v === "rgba(0,0,0,1)" || v === "black" ||
    v === "windowtext"
  );
}

function isDefaultBg(value: string): boolean {
  const v = value.replace(/\s+/g, "").toLowerCase();
  return (
    v === "" || v === "inherit" || v === "initial" || v === "transparent" ||
    v === "rgba(0,0,0,0)" || v === "#fff" || v === "#ffffff" ||
    v === "rgb(255,255,255)" || v === "rgba(255,255,255,1)" || v === "white"
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
// super/subscript, text/background colour). Headings are clamped to h3.
export function sanitizeWordHtml(rawHtml: string): string {
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

    const baseKeep = KEEP_ATTRS[tag] || new Set<string>();
    for (const attr of Array.from(el.attributes)) {
      if (!baseKeep.has(attr.name)) el.removeAttribute(attr.name);
    }
    if (preserved.length > 0 && STYLE_ALLOWED.has(tag)) {
      el.setAttribute("style", preserved.join("; "));
    }

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

export function plainTextToHtml(text: string): string {
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

// ---------- document splitting ----------

export type SectionId =
  | "overview"
  | "specifications"
  | "sizing"
  | "care"
  | "faq";

export const SECTION_META: { id: SectionId; label: string }[] = [
  { id: "overview", label: "Product overview" },
  { id: "specifications", label: "Specifications" },
  { id: "sizing", label: "Sizing & selection" },
  { id: "care", label: "Care & storage" },
  { id: "faq", label: "Frequently asked questions" },
];

export type SplitProduct = {
  title: string;
  sections: Record<SectionId, string>;
};

export type SplitResult = {
  products: SplitProduct[];
};

function emptySections(): Record<SectionId, string> {
  return { overview: "", specifications: "", sizing: "", care: "", faq: "" };
}

const BUYERS_GUIDE_RE = /complete buyer.?s guide/i;

function normalizeHeading(s: string): string {
  return s.replace(/\s+/g, " ").trim().toLowerCase();
}

// Map a heading's text to one of the 5 sections, "drop" (the marketing CTAs),
// or null (not a recognised section heading).
function bucketForHeading(rawText: string): SectionId | "drop" | null {
  const t = normalizeHeading(rawText);
  if (!t || t.length > 80) return null;
  if (/^ready to upgrade your supply/.test(t)) return "drop";
  if (/^what makes this product stand out/.test(t)) return "overview";
  if (/^what is this product made of/.test(t)) return "overview";
  if (/^what is it used for/.test(t)) return "overview";
  if (/^why you.?ll appreciate this product/.test(t)) return "overview";
  if (/^specs? and compliance/.test(t)) return "specifications";
  if (/^how to choose the right size/.test(t)) return "sizing";
  if (/^explore other options/.test(t)) return "sizing";
  if (/^care,? storage,? and disposal/.test(t)) return "care";
  if (/^frequently asked questions/.test(t)) return "faq";
  return null;
}

function isHeadingLike(el: Element): boolean {
  return /^(h[1-6]|p)$/i.test(el.tagName);
}

function isSectionHeading(el: Element): boolean {
  return isHeadingLike(el) && bucketForHeading(el.textContent || "") !== null;
}

function isProductTitle(el: Element): boolean {
  if (!isHeadingLike(el)) return false;
  const t = (el.textContent || "").trim();
  return t.length <= 200 && BUYERS_GUIDE_RE.test(t);
}

function isMetadataLine(el: Element): boolean {
  if (isSectionHeading(el) || isProductTitle(el)) return false;
  const t = (el.textContent || "").trim();
  if (!t) return true;
  return t.length <= 200 && /^(sku\b|id\b)/i.test(t);
}

function cleanProductTitle(raw: string): string {
  const t = raw.replace(/\s+/g, " ").trim();
  const stripped = t.replace(
    /\s*[-–—:]\s*complete buyer.?s guide\s*$/i,
    "",
  ).trim();
  return stripped || t;
}

function clampHeadingsIn(root: Element, doc: Document) {
  root.querySelectorAll("h1, h2").forEach((h) => {
    const r = doc.createElement("h3");
    for (const a of Array.from(h.attributes)) r.setAttribute(a.name, a.value);
    while (h.firstChild) r.appendChild(h.firstChild);
    h.replaceWith(r);
  });
}

function bucketChunk(
  nodes: Element[],
  doc: Document,
): Record<SectionId, string> {
  const buckets: Record<SectionId, Element[]> = {
    overview: [], specifications: [], sizing: [], care: [], faq: [],
  };
  let current: SectionId | "drop" = "overview"; // preamble -> overview
  for (const el of nodes) {
    if (isProductTitle(el)) continue; // safety; titles are excluded upstream
    if (isSectionHeading(el)) {
      const bucket = bucketForHeading(el.textContent || "");
      if (!bucket) continue;
      current = bucket;
      if (bucket === "drop") continue;
      const h = doc.createElement("h3");
      while (el.firstChild) h.appendChild(el.firstChild);
      buckets[bucket].push(h);
      continue;
    }
    if (current === "drop") continue;
    buckets[current].push(el);
  }
  const out = emptySections();
  for (const { id } of SECTION_META) {
    if (buckets[id].length === 0) continue;
    const wrapper = doc.createElement("div");
    for (const el of buckets[id]) wrapper.appendChild(el);
    wrapper.querySelectorAll("p").forEach((p) => {
      if (!p.textContent?.trim() && !p.querySelector("img")) p.remove();
    });
    clampHeadingsIn(wrapper, doc);
    out[id] = wrapper.innerHTML.trim();
  }
  return out;
}

// Split a converted/sanitised HTML document into 5 named sections. If the
// document contains several "... – Complete Buyer's Guide" titles, each one
// becomes its own product; otherwise it's treated as a single product.
export function splitDocument(html: string, fallbackTitle: string): SplitResult {
  const fallback = fallbackTitle.trim() || "Document";
  if (typeof window === "undefined") {
    return { products: [{ title: fallback, sections: emptySections() }] };
  }
  const doc = new DOMParser().parseFromString(html, "text/html");
  if (!doc.body) {
    return { products: [{ title: fallback, sections: emptySections() }] };
  }
  const nodes = Array.from(doc.body.children);
  if (nodes.length === 0) {
    return { products: [{ title: fallback, sections: emptySections() }] };
  }

  const titleIdxs: number[] = [];
  nodes.forEach((el, i) => {
    if (isProductTitle(el)) titleIdxs.push(i);
  });

  type Chunk = { title: string; nodes: Element[] };
  const chunks: Chunk[] = [];

  if (titleIdxs.length <= 1) {
    const titleEl = titleIdxs.length === 1 ? nodes[titleIdxs[0]] : null;
    const title = titleEl
      ? cleanProductTitle(titleEl.textContent || "")
      : fallback;
    chunks.push({ title, nodes: nodes.filter((el) => el !== titleEl) });
  } else {
    const boundaries: number[] = [0];
    for (let k = 1; k < titleIdxs.length; k++) {
      let b = titleIdxs[k];
      const lower = boundaries[boundaries.length - 1];
      while (b - 1 > lower && isMetadataLine(nodes[b - 1])) b--;
      boundaries.push(b);
    }
    boundaries.push(nodes.length);
    for (let k = 0; k < titleIdxs.length; k++) {
      const titleEl = nodes[titleIdxs[k]];
      const slice = nodes.slice(boundaries[k], boundaries[k + 1]);
      chunks.push({
        title: cleanProductTitle(titleEl.textContent || ""),
        nodes: slice.filter((el) => el !== titleEl),
      });
    }
  }

  const products: SplitProduct[] = chunks.map((c, i) => ({
    title:
      c.title?.trim() ||
      `${fallback}${chunks.length > 1 ? ` (${i + 1})` : ""}`,
    sections: bucketChunk(c.nodes, doc),
  }));

  return { products };
}
