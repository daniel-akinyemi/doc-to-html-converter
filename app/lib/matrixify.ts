// Fill a customer's Matrixify (Shopify) Excel export with the split sections,
// matched per "Variant SKU". The npm "xlsx" package is used for types only;
// the actual code runs against the standalone build loaded from /public.
import type * as XLSXModule from "xlsx";
import { SECTION_META, type SectionId, type SplitProduct } from "./docx";

declare global {
  interface Window {
    XLSX?: typeof XLSXModule;
  }
}

let xlsxPromise: Promise<typeof XLSXModule> | null = null;

function loadXlsx(): Promise<typeof XLSXModule> {
  if (typeof window === "undefined") {
    return Promise.reject(new Error("Export is only available in the browser."));
  }
  if (window.XLSX) return Promise.resolve(window.XLSX);
  if (xlsxPromise) return xlsxPromise;
  xlsxPromise = new Promise<typeof XLSXModule>((resolve, reject) => {
    const existing = document.querySelector<HTMLScriptElement>(
      'script[data-lib="xlsx"]',
    );
    const onReady = () => {
      if (window.XLSX) resolve(window.XLSX);
      else reject(new Error("Spreadsheet library loaded but is unavailable."));
    };
    if (existing) {
      existing.addEventListener("load", onReady, { once: true });
      existing.addEventListener(
        "error",
        () => reject(new Error("Failed to load the spreadsheet library.")),
        { once: true },
      );
      return;
    }
    const script = document.createElement("script");
    script.src = "/xlsx.full.min.js";
    script.async = true;
    script.dataset.lib = "xlsx";
    script.addEventListener("load", onReady, { once: true });
    script.addEventListener(
      "error",
      () => {
        xlsxPromise = null;
        reject(new Error("Failed to load the spreadsheet library."));
      },
      { once: true },
    );
    document.head.appendChild(script);
  });
  return xlsxPromise;
}

// Split section -> Matrixify metafield column. Matched by substring on the
// header text (so "Metafield: beyond.overview [rich_text_field]" works).
const SECTION_COLUMN: { id: SectionId; re: RegExp }[] = [
  { id: "overview", re: /beyond\.overview\b/i },
  { id: "specifications", re: /beyond\.specs\b/i },
  { id: "sizing", re: /beyond\.sizing_guide\b/i },
  { id: "care", re: /beyond\.care_storage\b/i },
  { id: "faq", re: /beyond\.faqs\b/i },
];

export type ExportReport = {
  sheetName: string;
  originalRows: number; // data rows in the uploaded sheet
  originalProducts: number; // distinct product blocks (by Handle) in the upload
  keptRows: number; // data rows remaining in the exported sheet
  droppedRows: number; // originalRows - keptRows
  rowsMatched: number; // rows we wrote section HTML into
  cellsWritten: number;
  mappedColumns: { section: string; header: string }[];
  missingColumns: string[]; // sections whose column wasn't found in the sheet
  matchedProducts: { title: string; key: string }[];
  unmatchedProducts: { title: string; key: string }[]; // had a key but no row
  productsWithoutKey: string[]; // doc had no "ID ..." / "SKU ..." line
};

export type ExportResult = {
  blob: Blob;
  fileName: string;
  report: ExportReport;
};

function cellText(value: unknown): string {
  if (value === undefined || value === null) return "";
  return String(value).trim();
}

export async function exportToMatrixify(
  fileBytes: ArrayBuffer,
  products: SplitProduct[],
  originalFileName: string,
): Promise<ExportResult> {
  const XLSX = await loadXlsx();
  let wb: XLSXModule.WorkBook;
  try {
    wb = XLSX.read(fileBytes, { type: "array" });
  } catch {
    throw new Error("That file isn’t a readable .xlsx workbook.");
  }

  // Find the sheet with a "Variant SKU" header.
  let sheetName = "";
  let ws: XLSXModule.WorkSheet | null = null;
  let headers: string[] = [];
  for (const name of wb.SheetNames) {
    const candidate = wb.Sheets[name];
    if (!candidate || !candidate["!ref"]) continue;
    const range = XLSX.utils.decode_range(candidate["!ref"]);
    const row: string[] = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = candidate[XLSX.utils.encode_cell({ r: range.s.r, c })];
      row[c] = cell ? cellText(cell.v) : "";
    }
    if (row.some((h) => h.toLowerCase() === "variant sku")) {
      sheetName = name;
      ws = candidate;
      headers = row;
      break;
    }
  }
  if (!ws) {
    throw new Error(
      'Couldn’t find a sheet with a "Variant SKU" column in that workbook.',
    );
  }

  const range = XLSX.utils.decode_range(ws["!ref"] as string);
  const skuCol = headers.findIndex((h) => h.toLowerCase() === "variant sku");

  const colFor = new Map<SectionId, number>();
  const mappedColumns: { section: string; header: string }[] = [];
  const missingColumns: string[] = [];
  for (const { id, re } of SECTION_COLUMN) {
    const label = SECTION_META.find((s) => s.id === id)?.label ?? id;
    const idx = headers.findIndex((h) => re.test(h));
    if (idx >= 0) {
      colFor.set(id, idx);
      mappedColumns.push({ section: label, header: headers[idx] });
    } else {
      missingColumns.push(label);
    }
  }
  if (colFor.size === 0) {
    throw new Error(
      "Couldn’t find any of the beyond.* metafield columns " +
        "(overview / specs / sizing_guide / care_storage / faqs) in that workbook.",
    );
  }

  // Index products by every key they carry — the "ID xxx" code (what the
  // Variant SKU column actually holds) and the supplier "SKU xxx" as a fallback.
  const byKey = new Map<string, SplitProduct>();
  const primaryKey = (p: SplitProduct) => p.id.trim() || p.sku.trim();
  const productsWithoutKey: string[] = [];
  for (const p of products) {
    const id = p.id.trim();
    const sku = p.sku.trim();
    if (id) byKey.set(id, p);
    if (sku && !byKey.has(sku)) byKey.set(sku, p);
    if (!id && !sku) productsWithoutKey.push(p.title);
  }

  const handleCol = headers.findIndex((h) => h.toLowerCase() === "handle");

  // Walk data rows, grouping them into product blocks by Handle (Matrixify
  // continuation rows repeat or blank the Handle and carry no Variant SKU).
  // A block is "matched" if any of its rows' Variant SKU maps to a product.
  type RowMeta = { r: number; group: string; product: SplitProduct | null };
  const rowMetas: RowMeta[] = [];
  const groupsSeen = new Set<string>();
  let currentGroup = "";
  for (let r = range.s.r + 1; r <= range.e.r; r++) {
    const handle =
      handleCol >= 0
        ? cellText(ws[XLSX.utils.encode_cell({ r, c: handleCol })]?.v)
        : "";
    const variantSku = cellText(
      ws[XLSX.utils.encode_cell({ r, c: skuCol })]?.v,
    );
    if (handle) currentGroup = handle;
    const group =
      handle || currentGroup || (variantSku ? `sku:${variantSku}` : `row:${r}`);
    groupsSeen.add(group);
    rowMetas.push({
      r,
      group,
      product: variantSku ? byKey.get(variantSku) ?? null : null,
    });
  }

  const matchedGroups = new Set<string>();
  for (const m of rowMetas) if (m.product) matchedGroups.add(m.group);

  // Write section HTML into the rows that matched a split product.
  const matchedProductSet = new Set<SplitProduct>();
  let rowsMatched = 0;
  let cellsWritten = 0;
  for (const m of rowMetas) {
    if (!m.product) continue;
    rowsMatched++;
    matchedProductSet.add(m.product);
    for (const [id, c] of colFor) {
      const html = m.product.sections[id] || "";
      if (!html.trim()) continue; // leave the cell as-is when we have nothing
      ws[XLSX.utils.encode_cell({ r: m.r, c })] = { t: "s", v: html };
      cellsWritten++;
    }
  }

  // Rebuild the sheet keeping the header + only rows in matched blocks.
  const keptMetas = rowMetas.filter((m) => matchedGroups.has(m.group));
  const newWs: XLSXModule.WorkSheet = {};
  const outOriginRows = [range.s.r, ...keptMetas.map((m) => m.r)];
  outOriginRows.forEach((origR, outIdx) => {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const src = ws[XLSX.utils.encode_cell({ r: origR, c })];
      if (src !== undefined) {
        newWs[XLSX.utils.encode_cell({ r: outIdx, c })] = src;
      }
    }
  });
  newWs["!ref"] = XLSX.utils.encode_range({
    s: { r: 0, c: range.s.c },
    e: { r: outOriginRows.length - 1, c: range.e.c },
  });
  if (ws["!cols"]) newWs["!cols"] = ws["!cols"];
  wb.Sheets[sheetName] = newWs;

  const matchedProducts: { title: string; key: string }[] = [];
  const unmatchedProducts: { title: string; key: string }[] = [];
  for (const p of products) {
    const key = primaryKey(p);
    if (!key) continue;
    (matchedProductSet.has(p) ? matchedProducts : unmatchedProducts).push({
      title: p.title,
      key,
    });
  }

  const originalRows = rowMetas.length;
  const keptRows = keptMetas.length;

  const out = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  const blob = new Blob([out], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const base = originalFileName.replace(/\.xlsx$/i, "") || "matrixify";
  const fileName = `${base} — matched.xlsx`;

  return {
    blob,
    fileName,
    report: {
      sheetName,
      originalRows,
      originalProducts: groupsSeen.size,
      keptRows,
      droppedRows: originalRows - keptRows,
      rowsMatched,
      cellsWritten,
      mappedColumns,
      missingColumns,
      matchedProducts,
      unmatchedProducts,
      productsWithoutKey,
    },
  };
}
