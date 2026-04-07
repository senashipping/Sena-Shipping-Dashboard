/**
 * EmbeddedExcelWorkbook — Excel-like renderer
 *
 * Drop-in replacement. Keeps all parsing (ExcelJS + HyperFormula) untouched;
 * replaces the render layer with a proper spreadsheet UI:
 *   • Single floating <input> editor (click to select, double-click / F2 to edit)
 *   • Click column header → select whole column
 *   • Click row header → select whole row
 *   • Shift+click / Shift+Arrow → range selection
 *   • Arrow / Tab / Enter key navigation
 *   • Ctrl+A → select all
 *   • Ctrl+C → copy selection to clipboard
 */

import React from "react";
import { Alert, AlertDescription } from "../ui/alert";
import {
  cellStyleToCss,
  columnLetter,
  getContainingMergeRegion,
  getMergeSpanIfMaster,
  isMergeSkipCell,
  loadEmbeddedExcelSource,
  mergeEmbeddedExcelValue,
  stripEmbeddedExcelRuntimeFields,
  type EmbeddedExcelFieldValue,
  type EmbeddedExcelSheetData,
} from "../../lib/excelWorkbook";
import {
  buildHyperFormulaEngine,
  hfSheetKey,
  parseCellInputForHF,
  sliceGridFromHFValues,
} from "../../lib/excelHyperFormula";
import type { HyperFormula } from "hyperformula";

/* ─── helpers ──────────────────────────────────────────────────────────────── */

function submissionSourceLabel(source: string): string {
  return source.trim().startsWith("data:") ? "embedded" : source.trim();
}

function clamp(v: number, lo: number, hi: number) {
  return Math.max(lo, Math.min(hi, v));
}

/* ─── types ────────────────────────────────────────────────────────────────── */

interface Pos { r: number; c: number }

interface Selection {
  anchor: Pos;   // where the click / keyboard move started
  head: Pos;     // current cursor
  colOnly: boolean; // true when a column header was clicked (full column)
  rowOnly: boolean; // true when a row header was clicked (full row)
  all: boolean;  // Ctrl+A
}

interface FillDragState {
  source: Pos;
  target: Pos;
}

function selectionBounds(sel: Selection, numRows: number, numCols: number) {
  if (sel.all) return { r0: 0, c0: 0, r1: numRows - 1, c1: numCols - 1 };
  if (sel.colOnly) return { r0: 0, c0: Math.min(sel.anchor.c, sel.head.c), r1: numRows - 1, c1: Math.max(sel.anchor.c, sel.head.c) };
  if (sel.rowOnly) return { r0: Math.min(sel.anchor.r, sel.head.r), c0: 0, r1: Math.max(sel.anchor.r, sel.head.r), c1: numCols - 1 };
  return {
    r0: Math.min(sel.anchor.r, sel.head.r),
    c0: Math.min(sel.anchor.c, sel.head.c),
    r1: Math.max(sel.anchor.r, sel.head.r),
    c1: Math.max(sel.anchor.c, sel.head.c),
  };
}

function inSelection(sel: Selection | null, ri: number, ci: number, numRows: number, numCols: number): boolean {
  if (!sel) return false;
  const { r0, c0, r1, c1 } = selectionBounds(sel, numRows, numCols);
  return ri >= r0 && ri <= r1 && ci >= c0 && ci <= c1;
}

/* ─── constants ─────────────────────────────────────────────────────────────── */

const ROW_HDR_W = 30;
const COL_HDR_H = 24;
const DEFAULT_COL_W = 80;
const DEFAULT_ROW_H = 22;
const FALLBACK_LOGO_SRC = "/8dc4bf7b-d0fd-4d5b-8226-b992170ae3e6.jpg";

/* ─── sub-components ─────────────────────────────────────────────────────────── */

/** Floating formula bar above the grid */
const FormulaBar: React.FC<{
  cell: string;
  address: string;
  editing: boolean;
  editValue: string;
  onChange: (v: string) => void;
  onCommit: () => void;
  onCancel: () => void;
}> = ({ cell, address, editing, editValue, onChange, onCommit, onCancel }) => (
  <div style={{
    display: "flex", alignItems: "center", gap: 0,
    borderBottom: "1px solid var(--color-border-tertiary)",
    background: "var(--color-background-primary)",
    minHeight: 28,
  }}>
    <div style={{
      width: 72, minWidth: 72, textAlign: "center",
      fontSize: 12, color: "var(--color-text-secondary)",
      borderRight: "1px solid var(--color-border-tertiary)",
      padding: "3px 6px", fontFamily: "var(--font-mono)",
    }}>
      {address}
    </div>
    <div style={{ width: 28, display: "flex", alignItems: "center", justifyContent: "center", borderRight: "1px solid var(--color-border-tertiary)", fontSize: 13, color: "var(--color-text-secondary)", padding: "0 4px" }}>
      ƒx
    </div>
    <input
      readOnly={!editing}
      value={editing ? editValue : cell}
      onChange={editing ? (e) => onChange(e.target.value) : undefined}
      onKeyDown={(e) => {
        if (e.key === "Enter") onCommit();
        if (e.key === "Escape") onCancel();
      }}
      style={{
        flex: 1, border: "none", outline: "none", background: "transparent",
        fontSize: 12, padding: "3px 8px",
        fontFamily: "var(--font-mono)", color: "var(--color-text-primary)",
      }}
    />
  </div>
);

/* ─── main component ─────────────────────────────────────────────────────────── */

interface EmbeddedExcelWorkbookProps {
  excelSource: string;
  value: EmbeddedExcelFieldValue | null | undefined;
  onChange: (next: EmbeddedExcelFieldValue) => void;
  readOnly?: boolean;
}

const EmbeddedExcelWorkbook: React.FC<EmbeddedExcelWorkbookProps> = ({
  excelSource,
  value,
  onChange,
  readOnly = false,
}) => {
  /* ── loading state ── */
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);
  const [sheets, setSheets] = React.useState<EmbeddedExcelSheetData[]>([]);
  const [activeSheet, setActiveSheet] = React.useState(0);

  /* ── selection & editing ── */
  const [sel, setSel] = React.useState<Selection | null>(null);
  const [editing, setEditing] = React.useState(false);
  const [editValue, setEditValue] = React.useState("");
  const editInputRef = React.useRef<HTMLInputElement>(null);
  const tableRef = React.useRef<HTMLDivElement>(null);
  const [fillDrag, setFillDrag] = React.useState<FillDragState | null>(null);

  /* ── existing infra (unchanged) ── */
  const loadedRef = React.useRef<EmbeddedExcelFieldValue | null>(null);
  const valueRef = React.useRef(value);
  valueRef.current = value;
  const hfRef = React.useRef<HyperFormula | null>(null);
  const valueSheetsJson = React.useMemo(() => JSON.stringify(value?.sheets ?? null), [value?.sheets]);

  const rebuildHyperFormulaFromMerged = React.useCallback((merged: EmbeddedExcelFieldValue) => {
    hfRef.current?.destroy();
    hfRef.current = null;
    const payload: Record<string, import("hyperformula").RawCellContent[][]> = {};
    for (const s of merged.sheets) {
      if (s.hfFullSheet?.length) payload[hfSheetKey(s.name)] = s.hfFullSheet;
    }
    if (!Object.keys(payload).length) return;
    try { hfRef.current = buildHyperFormulaEngine(payload); } catch { hfRef.current = null; return; }

    const hasSaved = Boolean(valueRef.current?.sheets?.length);
    if (hasSaved && hfRef.current) {
      for (const s of merged.sheets) {
        const sid = hfRef.current.getSheetId(hfSheetKey(s.name));
        if (sid === undefined || !s.excelBounds) continue;
        const { top, left } = s.excelBounds;
        for (let ri = 0; ri < s.grid.length; ri++) {
          for (let ci = 0; ci < s.grid[ri].length; ci++) {
            if (s.formulaCells?.[ri]?.[ci]) continue;
            hfRef.current.setCellContents(
              { sheet: sid, row: top + ri - 1, col: left + ci - 1 },
              [[parseCellInputForHF(s.grid[ri][ci])]],
            );
          }
        }
      }
    }

    setSheets(merged.sheets.map((s) => {
      const sid = hfRef.current?.getSheetId(hfSheetKey(s.name));
      if (sid === undefined || !s.excelBounds || !hfRef.current) return s;
      const vals = hfRef.current.getSheetValues(sid);
      const { top, left } = s.excelBounds;
      const nr = s.grid.length, nc = s.grid[0]?.length ?? 0;
      return { ...s, grid: sliceGridFromHFValues(vals, top, left, nr, nc) };
    }));
  }, []);

  React.useEffect(() => { setActiveSheet(0); setSel(null); setEditing(false); }, [excelSource]);
  React.useEffect(() => { hfRef.current?.destroy(); hfRef.current = null; loadedRef.current = null; }, [excelSource]);

  React.useEffect(() => {
    let cancelled = false;
    setLoading(true); setError(null);
    (async () => {
      try {
        const loaded = await loadEmbeddedExcelSource(excelSource);
        if (cancelled) return;
        loadedRef.current = loaded;
        const merged = mergeEmbeddedExcelValue(loaded, valueRef.current ?? undefined);
        merged.sourceUrl = submissionSourceLabel(excelSource);
        setSheets(merged.sheets);
        rebuildHyperFormulaFromMerged(merged);
        if (!readOnly && !valueRef.current?.sheets?.length) onChange(stripEmbeddedExcelRuntimeFields(merged));
      } catch (e) {
        if (!cancelled) { setError(e instanceof Error ? e.message : "Failed to load spreadsheet."); setSheets([]); }
      } finally { if (!cancelled) setLoading(false); }
    })();
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [excelSource]);

  React.useEffect(() => {
    if (!loadedRef.current) return;
    const merged = mergeEmbeddedExcelValue(loadedRef.current, valueRef.current ?? undefined);
    merged.sourceUrl = submissionSourceLabel(excelSource);
    setSheets(merged.sheets);
    rebuildHyperFormulaFromMerged(merged);
  }, [excelSource, rebuildHyperFormulaFromMerged, valueSheetsJson]);

  React.useEffect(() => () => { hfRef.current?.destroy(); hfRef.current = null; }, []);

  /* ── cell update (same logic as before) ── */
  const updateCell = React.useCallback((sheetIndex: number, row: number, col: number, cell: string) => {
    if (readOnly) return;
    const label = submissionSourceLabel(excelSource);
    const hf = hfRef.current;
    setSheets((prev) => {
      const sheet = prev[sheetIndex];
      if (hf && sheet?.hfFullSheet && sheet.excelBounds) {
        const sid = hf.getSheetId(hfSheetKey(sheet.name));
        if (sid !== undefined) {
          const { top, left } = sheet.excelBounds;
          hf.setCellContents({ sheet: sid, row: top + row - 1, col: left + col - 1 }, [[parseCellInputForHF(cell)]]);
          const vals = hf.getSheetValues(sid);
          const nr = sheet.grid.length, nc = sheet.grid[0]?.length ?? 0;
          const newGrid = sliceGridFromHFValues(vals, top, left, nr, nc);
          const next = prev.map((s, si) => {
            if (si !== sheetIndex) return s;
            const fc = s.formulaCells?.map((r, ri) => r.map((f, ci) => (ri === row && ci === col ? false : f)));
            return { ...s, grid: newGrid, formulaCells: fc ?? s.formulaCells };
          });
          onChange(stripEmbeddedExcelRuntimeFields({ sheets: next, sourceUrl: label }));
          return next;
        }
      }
      const next = prev.map((s, si) => {
        if (si !== sheetIndex) return s;
        const g = s.grid.map((r) => [...r]);
        const region = getContainingMergeRegion(s.mergeRegions, row, col);
        if (region) {
          for (let r = region.r; r < region.r + region.rowspan; r++) {
            if (!g[r]) g[r] = [];
            const rowCopy = [...g[r]];
            while (rowCopy.length < region.c + region.colspan) rowCopy.push("");
            for (let c = region.c; c < region.c + region.colspan; c++) rowCopy[c] = cell;
            g[r] = rowCopy;
          }
        } else {
          if (!g[row]) g[row] = [];
          const rowCopy = [...g[row]];
          while (rowCopy.length <= col) rowCopy.push("");
          rowCopy[col] = cell;
          g[row] = rowCopy;
        }
        return { ...s, grid: g };
      });
      onChange(stripEmbeddedExcelRuntimeFields({ sheets: next, sourceUrl: label }));
      return next;
    });
  }, [excelSource, onChange, readOnly]);

  const applyMatrixUpdate = React.useCallback((
    sheetIndex: number,
    startRow: number,
    startCol: number,
    matrix: string[][],
  ) => {
    if (readOnly || !matrix.length) return;
    const label = submissionSourceLabel(excelSource);
    const hf = hfRef.current;
    setSheets((prev) => {
      const sheet = prev[sheetIndex];
      if (!sheet) return prev;

      if (hf && sheet.hfFullSheet && sheet.excelBounds) {
        const sid = hf.getSheetId(hfSheetKey(sheet.name));
        if (sid !== undefined) {
          const { top, left } = sheet.excelBounds;
          for (let ri = 0; ri < matrix.length; ri++) {
            const row = matrix[ri];
            for (let ci = 0; ci < row.length; ci++) {
              const gr = startRow + ri;
              const gc = startCol + ci;
              if (gr < 0 || gc < 0 || gr >= sheet.grid.length || gc >= (sheet.grid[0]?.length ?? 0)) continue;
              hf.setCellContents(
                { sheet: sid, row: top + gr - 1, col: left + gc - 1 },
                [[parseCellInputForHF(row[ci] ?? "")]],
              );
            }
          }
          const vals = hf.getSheetValues(sid);
          const nr = sheet.grid.length;
          const nc = sheet.grid[0]?.length ?? 0;
          const newGrid = sliceGridFromHFValues(vals, top, left, nr, nc);
          const next = prev.map((s, si) => {
            if (si !== sheetIndex) return s;
            const fc = s.formulaCells?.map((r) => [...r]);
            if (fc) {
              for (let ri = 0; ri < matrix.length; ri++) {
                for (let ci = 0; ci < matrix[ri].length; ci++) {
                  const gr = startRow + ri;
                  const gc = startCol + ci;
                  if (gr >= 0 && gc >= 0 && gr < fc.length && gc < (fc[gr]?.length ?? 0)) {
                    fc[gr][gc] = false;
                  }
                }
              }
            }
            return { ...s, grid: newGrid, formulaCells: fc ?? s.formulaCells };
          });
          onChange(stripEmbeddedExcelRuntimeFields({ sheets: next, sourceUrl: label }));
          return next;
        }
      }

      const next = prev.map((s, si) => {
        if (si !== sheetIndex) return s;
        const g = s.grid.map((r) => [...r]);
        for (let ri = 0; ri < matrix.length; ri++) {
          const sourceRow = matrix[ri];
          const gr = startRow + ri;
          if (gr < 0 || gr >= g.length) continue;
          for (let ci = 0; ci < sourceRow.length; ci++) {
            const gc = startCol + ci;
            if (gc < 0) continue;
            while (g[gr].length <= gc) g[gr].push("");
            g[gr][gc] = sourceRow[ci] ?? "";
          }
        }
        return { ...s, grid: g };
      });
      onChange(stripEmbeddedExcelRuntimeFields({ sheets: next, sourceUrl: label }));
      return next;
    });
  }, [excelSource, onChange, readOnly]);

  /* ── selection helpers ── */
  const sheet = sheets[Math.min(activeSheet, sheets.length - 1)];
  const numRows = sheet?.grid.length ?? 0;
  const numCols = sheet?.grid[0]?.length ?? 0;

  function addressOf(r: number, c: number) {
    return `${columnLetter(c)}${r + 1}`;
  }

  function cellValue(r: number, c: number): string {
    return sheet?.grid[r]?.[c] ?? "";
  }

  /* ── commit / cancel edit ── */
  const commitEdit = React.useCallback(() => {
    if (!editing || !sel) return;
    updateCell(activeSheet, sel.head.r, sel.head.c, editValue);
    setEditing(false);
    tableRef.current?.focus();
  }, [editing, sel, editValue, updateCell, activeSheet]);

  const cancelEdit = React.useCallback(() => {
    setEditing(false);
    tableRef.current?.focus();
  }, []);

  /* ── enter edit mode ── */
  const startEdit = React.useCallback((initialChar?: string) => {
    if (readOnly || !sel) return;
    const cur = cellValue(sel.head.r, sel.head.c);
    setEditValue(initialChar ?? cur);
    setEditing(true);
    setTimeout(() => editInputRef.current?.focus(), 0);
  }, [readOnly, sel, sheet]);

  /* ── keyboard navigation ── */
  const handleKeyDown = React.useCallback((e: React.KeyboardEvent) => {
    if (!sel) return;

    /* Ctrl combos */
    if (e.ctrlKey || e.metaKey) {
      if (e.key === "a" || e.key === "A") {
        e.preventDefault();
        setSel({ anchor: { r: 0, c: 0 }, head: { r: numRows - 1, c: numCols - 1 }, colOnly: false, rowOnly: false, all: true });
        return;
      }
      if (e.key === "c" || e.key === "C") {
        // copy selection as TSV
        const { r0, c0, r1, c1 } = selectionBounds(sel, numRows, numCols);
        const tsv = Array.from({ length: r1 - r0 + 1 }, (_, ri) =>
          Array.from({ length: c1 - c0 + 1 }, (_, ci) => sheet?.grid[r0 + ri]?.[c0 + ci] ?? "").join("\t")
        ).join("\n");
        navigator.clipboard?.writeText(tsv).catch(() => {});
        return;
      }
      if (e.key === "v" || e.key === "V") {
        return;
      }
      return;
    }

    if (editing) return;  // let the <input> handle keys when editing

    const { r, c } = sel.head;
    const shift = e.shiftKey;

    const move = (dr: number, dc: number) => {
      e.preventDefault();
      const nr = clamp(r + dr, 0, numRows - 1);
      const nc = clamp(c + dc, 0, numCols - 1);
      const newHead = { r: nr, c: nc };
      setSel(shift
        ? { ...sel, head: newHead, colOnly: false, rowOnly: false, all: false }
        : { anchor: newHead, head: newHead, colOnly: false, rowOnly: false, all: false });
    };

    switch (e.key) {
      case "ArrowUp":    move(-1, 0); break;
      case "ArrowDown":  move(1, 0); break;
      case "ArrowLeft":  move(0, -1); break;
      case "ArrowRight": move(0, 1); break;
      case "Tab":
        e.preventDefault();
        move(0, e.shiftKey ? -1 : 1);
        break;
      case "Enter":
        if (e.shiftKey) move(-1, 0);
        else move(1, 0);
        break;
      case "F2":
        e.preventDefault();
        startEdit();
        break;
      case "Delete":
      case "Backspace":
        e.preventDefault();
        if (!readOnly) {
          updateCell(activeSheet, r, c, "");
        }
        break;
      case "Escape":
        setSel(null);
        break;
      default:
        // printable char → start editing
        if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
          startEdit(e.key);
        }
    }
  }, [sel, editing, numRows, numCols, sheet, startEdit, updateCell, readOnly, activeSheet]);

  const handlePaste = React.useCallback((e: React.ClipboardEvent<HTMLDivElement>) => {
    if (readOnly || !sel) return;
    e.preventDefault();
    const text = e.clipboardData.getData("text/plain");
    if (!text) return;
    const normalized = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    const lines = normalized.split("\n");
    if (lines.length > 1 && lines[lines.length - 1] === "") lines.pop();
    const matrix = lines.map((line) => line.split("\t"));
    if (!matrix.length) return;
    const { r0, c0 } = selectionBounds(sel, numRows, numCols);
    applyMatrixUpdate(activeSheet, r0, c0, matrix);
    const endR = clamp(r0 + matrix.length - 1, 0, numRows - 1);
    const endC = clamp(c0 + (matrix[0]?.length ?? 1) - 1, 0, numCols - 1);
    setSel({ anchor: { r: r0, c: c0 }, head: { r: endR, c: endC }, colOnly: false, rowOnly: false, all: false });
    tableRef.current?.focus();
  }, [readOnly, sel, numRows, numCols, applyMatrixUpdate, activeSheet]);

  const finishFillDrag = React.useCallback(() => {
    if (!fillDrag || readOnly) return;
    const source = fillDrag.source;
    const target = fillDrag.target;
    const r0 = Math.min(source.r, target.r);
    const c0 = Math.min(source.c, target.c);
    const r1 = Math.max(source.r, target.r);
    const c1 = Math.max(source.c, target.c);
    const v = cellValue(source.r, source.c);
    const matrix = Array.from({ length: r1 - r0 + 1 }, () =>
      Array.from({ length: c1 - c0 + 1 }, () => v),
    );
    applyMatrixUpdate(activeSheet, r0, c0, matrix);
    setSel({ anchor: { r: r0, c: c0 }, head: { r: r1, c: c1 }, colOnly: false, rowOnly: false, all: false });
    setFillDrag(null);
  }, [fillDrag, readOnly, applyMatrixUpdate, activeSheet, cellValue]);

  React.useEffect(() => {
    if (!fillDrag) return;
    const onMouseUp = () => finishFillDrag();
    window.addEventListener("mouseup", onMouseUp);
    return () => window.removeEventListener("mouseup", onMouseUp);
  }, [fillDrag, finishFillDrag]);

  /* ── render ─────────────────────────────────────────────────────────────── */

  if (loading) {
    return (
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", padding: "48px 0", fontSize: 13, color: "var(--color-text-secondary)" }}>
        Loading spreadsheet…
      </div>
    );
  }
  if (error) {
    return <Alert variant="destructive"><AlertDescription>{error}</AlertDescription></Alert>;
  }
  if (!sheets.length) return <p style={{ fontSize: 13, color: "var(--color-text-secondary)" }}>No sheet data.</p>;

  const safeIndex = Math.min(activeSheet, sheets.length - 1);
  const regions = sheet?.mergeRegions;
  const styles = sheet?.cellStyles;
  const colWidths = sheet?.colWidthsPx;
  const rowHeights = sheet?.rowHeightsPx;
  const sheetImages =
    sheet?.images?.length
      ? sheet.images
      : [{ src: FALLBACK_LOGO_SRC, left: Math.max(16, (colWidths?.reduce((a, b) => a + b, 0) ?? 560) - 220), top: 10, width: 190, height: 72 }];

  /* address display */
  const curAddress = sel ? (() => {
    if (sel.all) return `A1:${columnLetter(numCols - 1)}${numRows}`;
    if (sel.colOnly) {
      const c0 = Math.min(sel.anchor.c, sel.head.c), c1 = Math.max(sel.anchor.c, sel.head.c);
      return c0 === c1 ? `${columnLetter(c0)}:${columnLetter(c0)}` : `${columnLetter(c0)}:${columnLetter(c1)}`;
    }
    if (sel.rowOnly) {
      const r0 = Math.min(sel.anchor.r, sel.head.r), r1 = Math.max(sel.anchor.r, sel.head.r);
      return `${r0 + 1}:${r1 + 1}`;
    }
    const { r0, c0, r1, c1 } = selectionBounds(sel, numRows, numCols);
    if (r0 === r1 && c0 === c1) return addressOf(r0, c0);
    return `${addressOf(r0, c0)}:${addressOf(r1, c1)}`;
  })() : "—";

  const curValue = sel ? cellValue(sel.head.r, sel.head.c) : "";

  return (
    <div style={{ display: "flex", flexDirection: "column", border: "1px solid var(--color-border-tertiary)", borderRadius: 4, overflow: "hidden", background: "var(--color-background-primary)", userSelect: "none" }}>

      {/* sheet tabs */}
      {sheets.length > 1 && (
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "repeat(auto-fill, minmax(140px, 1fr))",
            gridTemplateRows: "repeat(2, minmax(0, auto))",
            gridAutoFlow: "column",
            gap: 4,
            padding: "6px 6px 4px",
            borderBottom: "1px solid var(--color-border-tertiary)",
            background: "var(--color-background-secondary)",
            overflow: "hidden",
          }}
        >
          {sheets.map((s, i) => (
            <button
              key={i}
              onClick={() => { setActiveSheet(i); setSel(null); setEditing(false); }}
              style={{
                padding: "7px 10px",
                fontSize: 12,
                cursor: "pointer",
                border: "1px solid #1478DC",
                outline: safeIndex === i ? "2px solid rgba(20, 120, 220, 0.24)" : "1px solid rgba(0, 0, 0, 0.04)",
                outlineOffset: 0,
                borderRadius: 6,
                background: safeIndex === i ? "linear-gradient(180deg, rgba(20, 120, 220, 0.18), rgba(20, 120, 220, 0.08))" : "var(--color-background-primary)",
                color: safeIndex === i ? "var(--color-text-primary)" : "var(--color-text-secondary)",
                fontWeight: safeIndex === i ? 600 : 500,
                textAlign: "left",
                minWidth: 0,
                overflow: "hidden",
                textOverflow: "ellipsis",
                whiteSpace: "nowrap",
                boxShadow: safeIndex === i ? "0 0 0 1px rgba(20,120,220,0.15) inset" : "0 1px 1px rgba(0,0,0,0.03)",
                transition: "background 120ms ease, border-color 120ms ease, box-shadow 120ms ease",
                fontFamily: "Segoe UI, Arial, sans-serif",
              }}
              title={s.name || `Sheet ${i + 1}`}
            >
              {s.name || `Sheet ${i + 1}`}
            </button>
          ))}
        </div>
      )}

      <FormulaBar
        address={curAddress}
        cell={curValue}
        editing={editing}
        editValue={editValue}
        onChange={setEditValue}
        onCommit={commitEdit}
        onCancel={cancelEdit}
      />

      {/* grid */}
      <div
        ref={tableRef}
        tabIndex={0}
        onKeyDown={handleKeyDown}
        onPaste={handlePaste}
        onBlur={(e) => { if (!e.currentTarget.contains(e.relatedTarget as Node)) { if (editing) commitEdit(); } }}
        style={{
          overflow: "auto",
          maxHeight: "min(70vh, 560px)",
          outline: "none",
          position: "relative",
        }}
      >
        <table
          style={{
            borderCollapse: "collapse",
            tableLayout: "fixed",
            fontSize: 12,
            lineHeight: "1.3",
            background: "var(--color-background-primary)",
          }}
        >
          <colgroup>
            {/* row-header col */}
            <col style={{ width: ROW_HDR_W, minWidth: ROW_HDR_W }} />
            {sheet?.grid[0]?.map((_, ci) => {
              const w = colWidths?.[ci] ?? DEFAULT_COL_W;
              return <col key={ci} style={{ width: w, minWidth: 24 }} />;
            })}
          </colgroup>

          {/* column headers */}
          <thead>
            <tr style={{ height: COL_HDR_H }}>
              {/* corner */}
              <th
                onClick={() => setSel({ anchor: { r: 0, c: 0 }, head: { r: numRows - 1, c: numCols - 1 }, colOnly: false, rowOnly: false, all: true })}
                style={{ ...hdrStyle(false, false), top: 0, left: 0, zIndex: 6, minWidth: ROW_HDR_W, fontWeight: 600 }}
                title="Select all"
              />
              {sheet?.grid[0]?.map((_, ci) => {
                const isColSelected = sel ? (() => {
                  if (sel.all) return true;
                  if (sel.colOnly) {
                    const c0 = Math.min(sel.anchor.c, sel.head.c), c1 = Math.max(sel.anchor.c, sel.head.c);
                    return ci >= c0 && ci <= c1;
                  }
                  return false;
                })() : false;
                const isActive = !sel?.colOnly && sel?.head.c === ci;
                return (
                  <th
                    key={ci}
                    onClick={(e) => {
                      if (editing) commitEdit();
                      const newSel: Selection = { anchor: { r: 0, c: ci }, head: { r: numRows - 1, c: ci }, colOnly: true, rowOnly: false, all: false };
                      if (e.shiftKey && sel) {
                        setSel({ anchor: sel.anchor, head: { r: numRows - 1, c: ci }, colOnly: true, rowOnly: false, all: false });
                      } else {
                        setSel(newSel);
                      }
                      tableRef.current?.focus();
                    }}
                    style={{ ...hdrStyle(isColSelected, isActive), top: 0, zIndex: 5, minWidth: 56, fontWeight: 600 }}
                  >
                    {columnLetter(ci)}
                  </th>
                );
              })}
            </tr>
          </thead>

          {/* body */}
          <tbody>
            {sheet?.grid.map((row, ri) => {
              const rowH = rowHeights?.[ri] ?? DEFAULT_ROW_H;
              const isRowSelected = sel ? (() => {
                if (sel.all) return true;
                if (sel.rowOnly) {
                  const r0 = Math.min(sel.anchor.r, sel.head.r), r1 = Math.max(sel.anchor.r, sel.head.r);
                  return ri >= r0 && ri <= r1;
                }
                return false;
              })() : false;
              const isActiveRow = sel?.head.r === ri && !sel?.rowOnly;

              return (
                <tr key={ri} style={{ height: rowH }}>
                  {/* row header */}
                  <td
                    onClick={(e) => {
                      if (editing) commitEdit();
                      if (e.shiftKey && sel) {
                        setSel({ anchor: sel.anchor, head: { r: ri, c: numCols - 1 }, colOnly: false, rowOnly: true, all: false });
                      } else {
                        setSel({ anchor: { r: ri, c: 0 }, head: { r: ri, c: numCols - 1 }, colOnly: false, rowOnly: true, all: false });
                      }
                      tableRef.current?.focus();
                    }}
                    style={{ ...hdrStyle(isRowSelected, isActiveRow && !sel?.rowOnly), left: 0, zIndex: 4, textAlign: "right", paddingRight: 8, fontWeight: 500 }}
                  >
                    {ri + 1}
                  </td>

                  {row.map((cellText, ci) => {
                    if (isMergeSkipCell(regions, ri, ci)) return null;
                    const span = getMergeSpanIfMaster(regions, ri, ci);
                    const st = styles?.[ri]?.[ci];
                    const tdCss = cellStyleToCss(st);
                    const isActive = sel ? (sel.head.r === ri && sel.head.c === ci && !sel.colOnly && !sel.rowOnly && !sel.all) : false;
                    const isInSel = inSelection(sel, ri, ci, numRows, numCols);

                    return (
                      <td
                        key={`${ri}-${ci}`}
                        rowSpan={span?.rowspan}
                        colSpan={span?.colspan}
                        onMouseDown={(e) => {
                          e.preventDefault();
                          if (editing) commitEdit();
                          if (fillDrag) return;
                          const newPos = { r: ri, c: ci };
                          if (e.shiftKey && sel) {
                            setSel({ anchor: sel.anchor, head: newPos, colOnly: false, rowOnly: false, all: false });
                          } else {
                            setSel({ anchor: newPos, head: newPos, colOnly: false, rowOnly: false, all: false });
                          }
                          tableRef.current?.focus();
                        }}
                        onMouseEnter={() => {
                          if (!fillDrag) return;
                          setFillDrag((prev) => (prev ? { ...prev, target: { r: ri, c: ci } } : prev));
                        }}
                        onDoubleClick={() => { if (!readOnly) startEdit(); }}
                        style={{
                          ...tdCss,
                          padding: "1px 4px",
                          overflow: "hidden",
                          whiteSpace: st?.wrapText ? "pre-wrap" : "nowrap",
                          textOverflow: st?.wrapText ? undefined : "ellipsis",
                          verticalAlign: st?.verticalAlign === "middle" ? "middle" : st?.verticalAlign === "bottom" ? "bottom" : "top",
                          boxSizing: "border-box",
                          position: "relative",
                          cursor: "default",
                          /* selection background */
                          background: isActive
                            ? "var(--color-background-primary)"
                            : isInSel
                            ? "rgba(20, 120, 220, 0.12)"
                            : (tdCss.backgroundColor as string | undefined) ?? "var(--color-background-primary)",
                          /* cell border */
                          borderTop: isActive ? "2px solid #1478DC" : `1px solid ${tdCss.borderTop ? "transparent" : "var(--color-border-tertiary)"}`,
                          borderBottom: isActive ? "2px solid #1478DC" : `1px solid ${tdCss.borderBottom ? "transparent" : "var(--color-border-tertiary)"}`,
                          borderLeft: isActive ? "2px solid #1478DC" : `1px solid ${tdCss.borderLeft ? "transparent" : "var(--color-border-tertiary)"}`,
                          borderRight: isActive ? "2px solid #1478DC" : `1px solid ${tdCss.borderRight ? "transparent" : "var(--color-border-tertiary)"}`,
                          ...(tdCss.borderTop ? { borderTop: isActive ? "2px solid #1478DC" : tdCss.borderTop } : {}),
                          ...(tdCss.borderBottom ? { borderBottom: isActive ? "2px solid #1478DC" : tdCss.borderBottom } : {}),
                          ...(tdCss.borderLeft ? { borderLeft: isActive ? "2px solid #1478DC" : tdCss.borderLeft } : {}),
                          ...(tdCss.borderRight ? { borderRight: isActive ? "2px solid #1478DC" : tdCss.borderRight } : {}),
                        }}
                      >
                        {/* floating inline editor for the active cell */}
                        {isActive && editing ? (
                          <input
                            ref={editInputRef}
                            value={editValue}
                            onChange={(e) => setEditValue(e.target.value)}
                            onKeyDown={(e) => {
                              if (e.key === "Enter") { e.preventDefault(); commitEdit(); }
                              if (e.key === "Escape") { e.preventDefault(); cancelEdit(); }
                              if (e.key === "Tab") { e.preventDefault(); commitEdit(); }
                              e.stopPropagation();
                            }}
                            style={{
                              position: "absolute", inset: 0,
                              border: "none", outline: "none",
                              padding: "1px 4px",
                              background: "var(--color-background-primary)",
                              fontFamily: st?.fontFamily ?? "var(--font-sans)",
                              fontSize: st?.fontSizePt != null ? `${st.fontSizePt}pt` : 12,
                              fontWeight: st?.bold ? 700 : undefined,
                              fontStyle: st?.italic ? "italic" : undefined,
                              color: st?.color ?? "var(--color-text-primary)",
                              textAlign: st?.textAlign ?? "left",
                              zIndex: 10,
                              boxSizing: "border-box",
                              width: "100%",
                              minWidth: "100%",
                            }}
                          />
                        ) : (
                          <span style={{
                            display: "block",
                            overflow: "hidden",
                            textOverflow: st?.wrapText ? undefined : "ellipsis",
                            whiteSpace: st?.wrapText ? "pre-wrap" : "nowrap",
                            maxWidth: "100%",
                          }}>
                            {cellText}
                          </span>
                        )}
                        {isActive && !editing && !readOnly && (
                          <div
                            onMouseDown={(e) => {
                              e.preventDefault();
                              e.stopPropagation();
                              setFillDrag({ source: { r: ri, c: ci }, target: { r: ri, c: ci } });
                            }}
                            title="Drag to fill"
                            style={{
                              position: "absolute",
                              width: 6,
                              height: 6,
                              right: -4,
                              bottom: -4,
                              background: "#1478DC",
                              border: "1px solid #ffffff",
                              boxSizing: "border-box",
                              cursor: "crosshair",
                              zIndex: 15,
                            }}
                          />
                        )}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
        {sheetImages.length > 0 && (
          <div
            style={{
              position: "absolute",
              top: COL_HDR_H,
              left: ROW_HDR_W,
              pointerEvents: "none",
              zIndex: 2,
            }}
          >
            {sheetImages.map((img, idx) => (
              <img
                key={`sheet-img-${idx}`}
                src={img.src}
                alt=""
                draggable={false}
                style={{
                  position: "absolute",
                  left: img.left,
                  top: img.top,
                  width: img.width,
                  height: img.height,
                  objectFit: "contain",
                  userSelect: "none",
                }}
              />
            ))}
          </div>
        )}
      </div>

      {/* status bar */}
      <div style={{
        display: "flex", gap: 16, alignItems: "center",
        padding: "2px 8px",
        borderTop: "1px solid var(--color-border-tertiary)",
        background: "var(--color-background-secondary)",
        fontSize: 11, color: "var(--color-text-secondary)",
      }}>
        <span>{readOnly ? "Read-only" : "Click to select • Double-click or F2 to edit • Ctrl+C to copy • Ctrl+V to paste • Drag blue handle to fill"}</span>
        {sel && (() => {
          const { r0, c0, r1, c1 } = selectionBounds(sel, numRows, numCols);
          const rows = r1 - r0 + 1, cols = c1 - c0 + 1;
          if (rows > 1 || cols > 1) {
            const nums = [];
            for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) {
              const n = parseFloat(sheet?.grid[r]?.[c] ?? "");
              if (!isNaN(n)) nums.push(n);
            }
            const sum = nums.reduce((a, b) => a + b, 0);
            return (
              <span style={{ marginLeft: "auto" }}>
                {rows > 1 && cols > 1 ? `${rows}R × ${cols}C` : rows > 1 ? `${rows} rows` : `${cols} cols`}
                {nums.length > 0 && ` · Sum: ${sum.toLocaleString()} · Avg: ${(sum / nums.length).toLocaleString(undefined, { maximumFractionDigits: 2 })}`}
              </span>
            );
          }
          return null;
        })()}
      </div>
    </div>
  );
};

/* ─── style helpers ──────────────────────────────────────────────────────────── */

function hdrStyle(selected: boolean, active: boolean): React.CSSProperties {
  return {
    padding: "2px 4px",
    textAlign: "center",
    fontSize: 11,
    fontWeight: 400,
    cursor: "default",
    position: "sticky",
    userSelect: "none",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
    color: "#000000",
    background: selected
      ? "#c6d9f1"
      : active
      ? "#d9d9d9"
      : "#e6e6e6",
    borderRight: "1px solid #000000",
    borderBottom: "1px solid #000000",
    boxSizing: "border-box",
  };
}

export default EmbeddedExcelWorkbook;