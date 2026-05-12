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
  rowsScanned: number; // data rows with a non-empty Variant SKU
  rowsMatched: number; // ...of those, ones whose SKU matched a split product
  cellsWritten: number;
  mappedColumns: { section: string; header: string }[];
  missingColumns: string[]; // sections whose column wasn't found in the sheet
  matchedProducts: { title: string; sku: string }[];
  unmatchedProducts: { title: string; sku: string }[]; // had a SKU but no row
  productsWithoutSku: string[]; // doc had no "SKU ..." line
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

  const bySku = new Map<string, SplitProduct>();
  const productsWithoutSku: string[] = [];
  for (const p of products) {
    const key = p.sku.trim();
    if (key) bySku.set(key, p);
    else productsWithoutSku.push(p.title);
  }

  const matchedSkus = new Set<string>();
  let rowsScanned = 0;
  let rowsMatched = 0;
  let cellsWritten = 0;

  for (let r = range.s.r + 1; r <= range.e.r; r++) {
    const skuCell = ws[XLSX.utils.encode_cell({ r, c: skuCol })];
    const sku = skuCell ? cellText(skuCell.v) : "";
    if (!sku) continue;
    rowsScanned++;
    const product = bySku.get(sku);
    if (!product) continue;
    rowsMatched++;
    matchedSkus.add(sku);
    for (const [id, c] of colFor) {
      const html = product.sections[id] || "";
      if (!html.trim()) continue; // leave the cell as-is when we have nothing
      ws[XLSX.utils.encode_cell({ r, c })] = { t: "s", v: html };
      cellsWritten++;
    }
  }

  const matchedProducts: { title: string; sku: string }[] = [];
  const unmatchedProducts: { title: string; sku: string }[] = [];
  for (const p of products) {
    const key = p.sku.trim();
    if (!key) continue;
    (matchedSkus.has(key) ? matchedProducts : unmatchedProducts).push({
      title: p.title,
      sku: key,
    });
  }

  const out = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  const blob = new Blob([out], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const base = originalFileName.replace(/\.xlsx$/i, "") || "matrixify";
  const fileName = `${base} — updated.xlsx`;

  return {
    blob,
    fileName,
    report: {
      sheetName,
      rowsScanned,
      rowsMatched,
      cellsWritten,
      mappedColumns,
      missingColumns,
      matchedProducts,
      unmatchedProducts,
      productsWithoutSku,
    },
  };
}
