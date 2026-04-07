/**
 * Bridges ExcelJS cells to HyperFormula so .xlsx formulas recalculate in the browser.
 * Uses GPLv3 HyperFormula — ensure your product complies or use a commercial HF license.
 */

import type { Cell } from "exceljs";
import { ValueType } from "exceljs";
import type { CellValue } from "hyperformula";
import { DetailedCellError, HyperFormula } from "hyperformula";
import type { RawCellContent } from "hyperformula";

export function formatHyperFormulaDisplay(v: CellValue): string {
  if (v === null || v === undefined) return "";
  if (v instanceof DetailedCellError) return v.value;
  if (typeof v === "number" && Number.isFinite(v)) {
    return String(v);
  }
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  const asUnknown = v as unknown;
  if (asUnknown instanceof Date) return asUnknown.toLocaleString();
  return String(v);
}

/** Excel sheet name safe for HyperFormula keys (avoid empty / invalid). */
export function hfSheetKey(name: string): string {
  const s = name.trim() || "Sheet1";
  return s.length > 31 ? s.slice(0, 31) : s;
}

export function cellToHFContent(cell: Cell): { content: RawCellContent; isFormula: boolean } {
  if (cell.isMerged && cell.address !== cell.master.address) {
    return { content: null, isFormula: false };
  }
  const c = cell.isMerged ? cell.master : cell;
  const raw = c.value as unknown;

  // ExcelJS can expose formulas as object values (including shared formulas)
  // even when the cell type is not strictly ValueType.Formula.
  if (raw && typeof raw === "object") {
    const maybe = raw as { formula?: unknown; sharedFormula?: unknown };
    const f = typeof maybe.formula === "string" ? maybe.formula : undefined;
    const sf = typeof maybe.sharedFormula === "string" ? maybe.sharedFormula : undefined;
    const expr = f ?? sf;
    if (expr && expr.trim() !== "") {
      const fs = expr.trim();
      return { content: fs.startsWith("=") ? fs : `=${fs}`, isFormula: true };
    }
  }

  if (c.type === ValueType.Formula) {
    let f = c.formula;
    if (f != null && f !== "") {
      const fs = String(f);
      return { content: fs.startsWith("=") ? fs : `=${fs}`, isFormula: true };
    }
    return { content: null, isFormula: false };
  }

  if (c.type === ValueType.Null || c.value === null || c.value === undefined) {
    return { content: null, isFormula: false };
  }

  if (c.type === ValueType.Number) {
    const n = c.value as number;
    return { content: Number.isFinite(n) ? n : null, isFormula: false };
  }
  if (c.type === ValueType.String || c.type === ValueType.SharedString) {
    return { content: String(c.value ?? ""), isFormula: false };
  }
  if (c.type === ValueType.Boolean) {
    return { content: Boolean(c.value), isFormula: false };
  }
  if (c.type === ValueType.Date) {
    return { content: c.value instanceof Date ? c.value : new Date(String(c.value)), isFormula: false };
  }
  if (c.type === ValueType.Error) {
    return { content: c.text || "#ERROR!", isFormula: false };
  }
  if (c.type === ValueType.RichText) {
    return { content: c.text || "", isFormula: false };
  }
  if (c.type === ValueType.Hyperlink) {
    return { content: c.text || "", isFormula: false };
  }
  if (c.type === ValueType.Merge) {
    return { content: null, isFormula: false };
  }

  const t = c.text;
  if (t != null && t !== "") return { content: t, isFormula: false };
  return { content: null, isFormula: false };
}

export function buildHyperFormulaEngine(sheets: Record<string, RawCellContent[][]>): HyperFormula {
  return HyperFormula.buildFromSheets(sheets, {
    licenseKey: "gpl-v3",
  });
}

/** Parse user-typed cell text for HyperFormula (constants only). */
export function parseCellInputForHF(text: string): RawCellContent {
  const t = text.trim();
  if (t === "") return null;
  if (/^true$/i.test(t)) return true;
  if (/^false$/i.test(t)) return false;
  if (/^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(t)) {
    const n = Number(t);
    if (Number.isFinite(n)) return n;
  }
  return t;
}

export function sliceGridFromHFValues(
  vals: CellValue[][],
  top: number,
  left: number,
  numRows: number,
  numCols: number,
): string[][] {
  const grid: string[][] = [];
  for (let gr = 0; gr < numRows; gr++) {
    const row: string[] = [];
    for (let gc = 0; gc < numCols; gc++) {
      const er = top + gr;
      const ec = left + gc;
      const v = vals[er - 1]?.[ec - 1] as CellValue | undefined;
      row.push(formatHyperFormulaDisplay(v ?? null));
    }
    grid.push(row);
  }
  return grid;
}
