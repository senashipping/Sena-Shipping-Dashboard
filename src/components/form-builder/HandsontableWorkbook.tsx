import React from "react";
import { HotTable } from "@handsontable/react";
import "handsontable/styles/handsontable.css";
import "handsontable/styles/ht-theme-main.css";
import { registerAllModules } from "handsontable/registry";
import { Button } from "../ui/button";
import { HyperFormula } from "hyperformula";
import ExcelJS from "exceljs";

registerAllModules();

export type SheetData = {
  name: string;
  grid: string[][];
  mergeCells?: Array<{
    row: number;
    col: number;
    rowspan: number;
    colspan: number;
  }>;
  cellMeta?: Array<{
    row: number;
    col: number;
    className?: string;
    type?: string;
    dateFormat?: string;
    correctFormat?: boolean;
    numericFormat?: { pattern?: string; culture?: string };
    source?: string[];
    strict?: boolean;
  }>;
  images?: Array<{
    row: number;
    col: number;
    rowspan?: number;
    colspan?: number;
    dataUrl: string;
  }>;
  colWidthsPx?: number[];
  rowHeightsPx?: number[];
  tabColor?: string;
};

interface HandsontableWorkbookProps {
  data: { sheets: SheetData[] };
  onChange: (next: { sheets: SheetData[] }) => void;
  /** When true, only cells tagged with `meta-fillable` can be edited (runtime / preview). When false, the full template is editable (builder / import). */
  readOnly?: boolean;
  /** Pixel height of the Handsontable viewport when `readOnly` (default 380). */
  readOnlyHotHeight?: number;
}

export type HandsontableWorkbookRef = {
  /**
   * Builder: reads the live grid from Handsontable (cells, merges, meta, col/row sizes) into `workbookRef`, then returns a deep clone for persistence.
   * Read-only: returns the current in-memory workbook without re-collecting from HOT (avoids truncating the full template).
   */
  getWorkbookSnapshot: () => { sheets: SheetData[] } | null;
};

/** Capped grid for read-only / form preview so imports with huge dimensions stay responsive. */
export const MAX_PREVIEW_ROWS = 220;
export const MAX_PREVIEW_COLS = 80;
const FORMULAS_CONFIG = { engine: HyperFormula };

/** Stable fallback so `safeGrid` / memoized slices do not get a fresh ref every render when the sheet is empty. */
const EMPTY_GRID_PLACEHOLDER: string[][] = [[""]];

// ─── helpers ──────────────────────────────────────────────────────────────────

const toSafeGrid = (rawGrid: unknown): string[][] => {
  if (!Array.isArray(rawGrid) || rawGrid.length === 0) return [[""]];
  const rows = rawGrid.map((row) =>
    Array.isArray(row) ? row.map((c) => (c == null ? "" : String(c))) : [""],
  );
  return rows.length > 0 ? rows : [[""]];
};

/** Handsontable mutates `data` in place — never pass the same nested refs stored in `workbookRef`. */
const cloneEditableGrid = (rawGrid: unknown): string[][] => {
  const g = toSafeGrid(rawGrid);
  return g.map((row) => (Array.isArray(row) ? [...row] : [""]));
};

/**
 * Outside Excel imports, merged regions can extend past the stored grid (or past `!ref`).
 * Handsontable misbehaves when `mergeCells` hangs off the data array — clip to the grid.
 */
const clipMergeCellsToGrid = (
  merges: NonNullable<SheetData["mergeCells"]>,
  gridRows: number,
  gridCols: number,
): NonNullable<SheetData["mergeCells"]> => {
  if (!merges.length || gridRows < 1 || gridCols < 1) return [];
  const maxR = gridRows - 1;
  const maxC = gridCols - 1;
  const out: NonNullable<SheetData["mergeCells"]> = [];
  for (const m of merges) {
    if (!m) continue;
    const r0 = +m.row;
    const c0 = +m.col;
    const rs = +m.rowspan;
    const cs = +m.colspan;
    if (![r0, c0, rs, cs].every((n) => Number.isFinite(n)) || rs < 1 || cs < 1)
      continue;
    const r1 = r0 + rs - 1;
    const c1 = c0 + cs - 1;
    const cr0 = Math.max(0, Math.min(maxR, r0));
    const cc0 = Math.max(0, Math.min(maxC, c0));
    const cr1 = Math.max(0, Math.min(maxR, r1));
    const cc1 = Math.max(0, Math.min(maxC, c1));
    if (cr1 < cr0 || cc1 < cc0) continue;
    const rowspan = cr1 - cr0 + 1;
    const colspan = cc1 - cc0 + 1;
    if (rowspan > 0 && colspan > 0)
      out.push({ row: cr0, col: cc0, rowspan, colspan });
  }
  return out;
};

const normalizeSheets = (input?: { sheets?: SheetData[] }): SheetData[] => {
  if (!Array.isArray(input?.sheets) || input.sheets.length === 0)
    return [{ name: "Sheet1", grid: [[""]] }];
  return input.sheets.map((sheet, i) => {
    const grid = toSafeGrid(sheet?.grid);
    const gridRows = grid.length;
    const gridCols = Math.max(
      1,
      grid.reduce(
        (w, row) => Math.max(w, Array.isArray(row) ? row.length : 0),
        0,
      ),
    );

    const mergeCells = clipMergeCellsToGrid(
      Array.isArray(sheet?.mergeCells)
        ? sheet.mergeCells
            .filter(
              (m: any) =>
                m &&
                Number.isFinite(+m.row) &&
                Number.isFinite(+m.col) &&
                Number.isFinite(+m.rowspan) &&
                Number.isFinite(+m.colspan) &&
                +m.rowspan > 0 &&
                +m.colspan > 0,
            )
            .map((m: any) => ({
              row: +m.row,
              col: +m.col,
              rowspan: +m.rowspan,
              colspan: +m.colspan,
            }))
        : [],
      gridRows,
      gridCols,
    );

    return {
      name: sheet?.name || `Sheet${i + 1}`,
      grid,
      mergeCells,
      cellMeta: dedupeCellMetaByCoordinate(
        Array.isArray(sheet?.cellMeta)
          ? sheet.cellMeta
              .filter(
                (m: any) =>
                  m &&
                  Number.isFinite(+m.row) &&
                  Number.isFinite(+m.col) &&
                  +m.row >= 0 &&
                  +m.col >= 0 &&
                  +m.row < gridRows &&
                  +m.col < gridCols,
              )
              .map((m: any) => ({
                row: +m.row,
                col: +m.col,
                className:
                  typeof m.className === "string" ? m.className : undefined,
                type: typeof m.type === "string" ? m.type : undefined,
                dateFormat:
                  typeof m.dateFormat === "string" ? m.dateFormat : undefined,
                correctFormat:
                  typeof m.correctFormat === "boolean"
                    ? m.correctFormat
                    : undefined,
                numericFormat:
                  m.numericFormat && typeof m.numericFormat === "object"
                    ? {
                        pattern:
                          typeof m.numericFormat.pattern === "string"
                            ? m.numericFormat.pattern
                            : undefined,
                        culture:
                          typeof m.numericFormat.culture === "string"
                            ? m.numericFormat.culture
                            : undefined,
                      }
                    : undefined,
                source: Array.isArray(m.source)
                  ? m.source.map(String)
                  : undefined,
                strict: typeof m.strict === "boolean" ? m.strict : undefined,
              }))
          : [],
      ),
      images: dedupeImagesByAnchor(
        Array.isArray((sheet as any)?.images)
          ? (sheet as any).images.filter(
              (img: any) =>
                img &&
                Number.isFinite(+img.row) &&
                Number.isFinite(+img.col) &&
                typeof img.dataUrl === "string" &&
                img.dataUrl.length > 0,
            )
          : [],
      ),
      colWidthsPx: Array.isArray(sheet?.colWidthsPx)
        ? [...sheet.colWidthsPx]
        : undefined,
      rowHeightsPx: Array.isArray(sheet?.rowHeightsPx)
        ? [...sheet.rowHeightsPx]
        : undefined,
      tabColor: sheet?.tabColor,
    };
  });
};

/** Include row/col pixel sizes so incoming `data` syncs when only dimensions change. */
const dimListSignature = (arr?: number[]) => {
  if (!arr?.length) return "";
  if (arr.length > 400) {
    let sum = 0;
    for (let i = 0; i < arr.length; i++) sum += Number(arr[i]) || 0;
    return `${arr.length}:sum${sum}:a${arr[0]}:z${arr[arr.length - 1]}`;
  }
  return arr.join(",");
};

const workbookSignature = (sheets: SheetData[]) =>
  sheets
    .map((s) =>
      [
        s.name,
        `${s.grid?.length || 0}x${s.grid?.[0]?.length || 0}`,
        `m${s.mergeCells?.length || 0}`,
        `c${s.cellMeta?.length || 0}`,
        s.tabColor || "",
        `cw${dimListSignature(s.colWidthsPx)}`,
        `rh${dimListSignature(s.rowHeightsPx)}`,
      ].join("|"),
    )
    .join("::");

/**
 * Shape-only signature for the HotTable React `key`. Must NOT include `cellMeta` length:
 * marking fillable cells changes meta count and was remounting HOT on every click (scroll
 * jump to top, lazy meta loss, glitchy renders).
 */
const hotTableMountSignature = (sheets: SheetData[]) =>
  sheets
    .map((s) =>
      [
        s.name,
        `${s.grid?.length || 0}x${s.grid?.[0]?.length || 0}`,
        `m${s.mergeCells?.length || 0}`,
        s.tabColor || "",
        `cw${dimListSignature(s.colWidthsPx)}`,
        `rh${dimListSignature(s.rowHeightsPx)}`,
      ].join("|"),
    )
    .join("::");

/** Deep snapshot so each sheet index has no shared nested refs with others or the parent. */
const deepCloneSheet = (s: SheetData): SheetData => ({
  name: s.name,
  tabColor: s.tabColor,
  grid: (s.grid?.length ? s.grid : [[""]]).map((row) =>
    Array.isArray(row) ? row.map((c) => (c == null ? "" : String(c))) : [""],
  ),
  mergeCells: (s.mergeCells || []).map((m) => ({
    row: m.row,
    col: m.col,
    rowspan: m.rowspan,
    colspan: m.colspan,
  })),
  cellMeta: dedupeCellMetaByCoordinate(
    (s.cellMeta || []).map((m) => ({
      row: m.row,
      col: m.col,
      className: m.className,
      type: m.type,
      dateFormat: m.dateFormat,
      correctFormat: m.correctFormat,
      numericFormat:
        m.numericFormat && typeof m.numericFormat === "object"
          ? {
              pattern: m.numericFormat.pattern,
              culture: m.numericFormat.culture,
            }
          : undefined,
      source: Array.isArray(m.source) ? m.source.map(String) : undefined,
      strict: m.strict,
    })),
  ),
  images: dedupeImagesByAnchor(
    (s.images || []).map((img) => ({
      row: img.row,
      col: img.col,
      rowspan: img.rowspan,
      colspan: img.colspan,
      dataUrl: img.dataUrl,
    })),
  ),
  colWidthsPx: s.colWidthsPx?.length ? [...s.colWidthsPx] : undefined,
  rowHeightsPx: s.rowHeightsPx?.length ? [...s.rowHeightsPx] : undefined,
});

type CellMetaEntry = NonNullable<SheetData["cellMeta"]>[number];

/** Stable 0-based cell index used everywhere we key cells (one logical cell = one bucket). */
const cellCoordKey = (row: number, col: number) => `${row}:${col}`;

const classNameHasFillable = (className?: string) =>
  String(className || "")
    .split(/\s+/)
    .filter(Boolean)
    .includes("meta-fillable");

const mergeClassNameStrings = (a?: string, b?: string) => {
  const tokens = new Set<string>();
  for (const t of String(a || "")
    .split(/\s+/)
    .filter(Boolean))
    tokens.add(t);
  for (const t of String(b || "")
    .split(/\s+/)
    .filter(Boolean))
    tokens.add(t);
  const out = [...tokens].join(" ").trim();
  return out || undefined;
};

/**
 * Handsontable / imports can produce multiple meta rows for the same (row,col).
 * Collapse to exactly one record per coordinate so cells cannot "fight" each other.
 */
const dedupeCellMetaByCoordinate = (
  list: NonNullable<SheetData["cellMeta"]>,
): NonNullable<SheetData["cellMeta"]> => {
  const map = new Map<string, CellMetaEntry>();
  for (const raw of list) {
    if (!raw || !Number.isFinite(+raw.row) || !Number.isFinite(+raw.col))
      continue;
    const row = +raw.row;
    const col = +raw.col;
    const key = cellCoordKey(row, col);
    const next: CellMetaEntry = {
      row,
      col,
      className: raw.className ? String(raw.className) : undefined,
      type: typeof raw.type === "string" ? raw.type : undefined,
      dateFormat:
        typeof raw.dateFormat === "string" ? raw.dateFormat : undefined,
      correctFormat:
        typeof raw.correctFormat === "boolean" ? raw.correctFormat : undefined,
      numericFormat:
        raw.numericFormat && typeof raw.numericFormat === "object"
          ? {
              pattern:
                typeof raw.numericFormat.pattern === "string"
                  ? raw.numericFormat.pattern
                  : undefined,
              culture:
                typeof raw.numericFormat.culture === "string"
                  ? raw.numericFormat.culture
                  : undefined,
            }
          : undefined,
      source: Array.isArray(raw.source) ? raw.source.map(String) : undefined,
      strict: typeof raw.strict === "boolean" ? raw.strict : undefined,
    };
    const prev = map.get(key);
    if (!prev) {
      map.set(key, next);
      continue;
    }
    map.set(key, {
      row,
      col,
      className: mergeClassNameStrings(prev.className, next.className),
      type: next.type ?? prev.type,
      dateFormat: next.dateFormat ?? prev.dateFormat,
      correctFormat:
        typeof next.correctFormat === "boolean"
          ? next.correctFormat
          : prev.correctFormat,
      numericFormat: next.numericFormat ?? prev.numericFormat,
      source: next.source ?? prev.source,
      strict: typeof next.strict === "boolean" ? next.strict : prev.strict,
    });
  }
  return [...map.values()].sort((a, b) =>
    a.row !== b.row ? a.row - b.row : a.col - b.col,
  );
};

/**
 * When the parent echoes `data` without `meta-fillable` (e.g. save/API round-trip), keep
 * fillable marks that still exist on the previous in-memory workbook.
 */
const mergeFillableMetaFromPrevSheet = (
  prev: SheetData | undefined,
  incoming: SheetData,
): SheetData => {
  const out = deepCloneSheet(incoming);
  if (!prev?.cellMeta?.length) return out;

  const prevFillKeys = new Set<string>();
  for (const m of prev.cellMeta) {
    if (!classNameHasFillable(m.className)) continue;
    if (!Number.isFinite(+m.row) || !Number.isFinite(+m.col)) continue;
    prevFillKeys.add(cellCoordKey(+m.row, +m.col));
  }
  if (prevFillKeys.size === 0) return out;

  const metaByKey = new Map<string, CellMetaEntry>();
  for (const m of out.cellMeta || []) {
    if (!Number.isFinite(+m.row) || !Number.isFinite(+m.col)) continue;
    metaByKey.set(cellCoordKey(+m.row, +m.col), {
      ...m,
      row: +m.row,
      col: +m.col,
    });
  }

  for (const key of prevFillKeys) {
    const cur = metaByKey.get(key);
    const cn = cur ? String(cur.className || "") : "";
    if (classNameHasFillable(cn)) continue;
    const [rs, cs] = key.split(":").map(Number);
    const tokens = new Set(cn.split(/\s+/).filter(Boolean));
    tokens.add("meta-fillable");
    metaByKey.set(key, {
      ...(cur || { row: rs, col: cs }),
      row: rs,
      col: cs,
      className: [...tokens].join(" ").trim() || undefined,
    });
  }

  out.cellMeta = dedupeCellMetaByCoordinate([...metaByKey.values()]);
  return out;
};

/** One image anchor per top-left cell (avoids duplicate overlays on same cell). */
const dedupeImagesByAnchor = (
  images: NonNullable<SheetData["images"]>,
): NonNullable<SheetData["images"]> => {
  const map = new Map<string, (typeof images)[number]>();
  for (const img of images) {
    if (
      !img ||
      !Number.isFinite(+img.row) ||
      !Number.isFinite(+img.col) ||
      typeof img.dataUrl !== "string"
    )
      continue;
    const row = +img.row;
    const col = +img.col;
    const key = cellCoordKey(row, col);
    if (!map.has(key))
      map.set(key, {
        row,
        col,
        rowspan: img.rowspan,
        colspan: img.colspan,
        dataUrl: img.dataUrl,
      });
  }
  return [...map.values()];
};

const toColumnLabel = (index: number) => {
  let n = index + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out || "A";
};

const toRangeLabel = (
  range: {
    startRow: number;
    endRow: number;
    startCol: number;
    endCol: number;
  } | null,
) => {
  if (!range) return "A1";
  const start = `${toColumnLabel(range.startCol)}${range.startRow + 1}`;
  const end = `${toColumnLabel(range.endCol)}${range.endRow + 1}`;
  return start === end ? start : `${start}:${end}`;
};

// Prevent toolbar buttons from stealing focus from the grid
const noFocusSteal = (e: React.MouseEvent) => e.preventDefault();

// ─── component ────────────────────────────────────────────────────────────────

const HandsontableWorkbook = React.forwardRef<
  HandsontableWorkbookRef,
  HandsontableWorkbookProps
>(function HandsontableWorkbook(
  { data, onChange, readOnly = false, readOnlyHotHeight },
  ref,
) {
  const workbookRef = React.useRef<{ sheets: SheetData[] }>({
    sheets: normalizeSheets(data).map(deepCloneSheet),
  });
  const lastIncomingSignatureRef = React.useRef(
    workbookSignature(normalizeSheets(data).map(deepCloneSheet)),
  );
  /** Preview (`readOnly`): edits stay local until focus leaves the workbook; then we sync once. */
  const readOnlyPreviewDirtyRef = React.useRef(false);

  const [activeSheetIndex, setActiveSheetIndex] = React.useState(0);
  const activeSheetIndexRef = React.useRef(0);
  /** Must run before paint so ref never lags `activeSheetIndex` (prevents saving the visible grid to the wrong sheet on fast tab switches). */
  React.useLayoutEffect(() => {
    activeSheetIndexRef.current = activeSheetIndex;
  }, [activeSheetIndex]);

  const [sheetTabs, setSheetTabs] = React.useState(
    workbookRef.current.sheets.map((s) => ({
      name: s.name,
      tabColor: s.tabColor,
    })),
  );
  const [initialGrid, setInitialGrid] = React.useState<string[][]>(() => {
    const first = workbookRef.current.sheets[0];
    const base =
      Array.isArray(first?.grid) && first.grid.length > 0 ? first.grid : [[""]];
    if (!readOnly) return cloneEditableGrid(base);
    const rows = Math.min(MAX_PREVIEW_ROWS, base.length);
    const cols = Math.min(MAX_PREVIEW_COLS, base[0]?.length || 0);
    return base
      .slice(0, rows)
      .map((row) => (Array.isArray(row) ? row.slice(0, cols) : []));
  });

  const [renaming, setRenaming] = React.useState(false);
  const [renameValue, setRenameValue] = React.useState("");
  const [formulaInput, setFormulaInput] = React.useState("");
  const [findValue, setFindValue] = React.useState("");
  const [replaceValue, setReplaceValue] = React.useState("");
  const [dropdownSource, setDropdownSource] =
    React.useState("Option A,Option B");
  const [fontFamily, setFontFamily] = React.useState("Arial");
  const [fontSize, setFontSize] = React.useState("12");
  const [textColor, setTextColor] = React.useState("#111827");
  const [fillColor, setFillColor] = React.useState("#ffffff");
  const [isBoldActive, setIsBoldActive] = React.useState(false);
  const [isItalicActive, setIsItalicActive] = React.useState(false);
  const [isUnderlineActive, setIsUnderlineActive] = React.useState(false);
  const [isStrikeActive, setIsStrikeActive] = React.useState(false);
  const [selectedAlign, setSelectedAlign] = React.useState<
    "left" | "center" | "right" | "justify" | null
  >(null);
  const [selectedVAlign, setSelectedVAlign] = React.useState<
    "top" | "middle" | "bottom" | null
  >(null);
  const [selectionLabel, setSelectionLabel] = React.useState("A1");
  const [canUndo, setCanUndo] = React.useState(false);
  const [canRedo, setCanRedo] = React.useState(false);
  const [formatAllCells, setFormatAllCells] = React.useState(false);
  const [fixedRowsTop, setFixedRowsTop] = React.useState(0);
  const [fixedColumnsStart, setFixedColumnsStart] = React.useState(0);

  // The most reliable selection store — never cleared, always the last valid range
  const lastSelectionRef = React.useRef<{
    startRow: number;
    endRow: number;
    startCol: number;
    endCol: number;
  }>({ startRow: 0, endRow: 0, startCol: 0, endCol: 0 });

  const sheetSelectionRef = React.useRef<
    Record<
      number,
      { startRow: number; endRow: number; startCol: number; endCol: number }
    >
  >({});

  const hotRef = React.useRef<any>(null);

  const textColorApplyTimerRef = React.useRef<ReturnType<
    typeof setTimeout
  > | null>(null);
  const fillColorApplyTimerRef = React.useRef<ReturnType<
    typeof setTimeout
  > | null>(null);

  const normalizedIncomingSheets = React.useMemo(
    () => normalizeSheets(data),
    [data],
  );

  const safeSheets = workbookRef.current.sheets;
  const activeSheet =
    safeSheets[Math.min(activeSheetIndex, safeSheets.length - 1)] ||
    safeSheets[0];
  const safeGrid =
    Array.isArray(activeSheet?.grid) && activeSheet.grid.length > 0
      ? activeSheet.grid
      : EMPTY_GRID_PLACEHOLDER;
  const previewRows = readOnly
    ? Math.min(MAX_PREVIEW_ROWS, safeGrid.length)
    : safeGrid.length;
  const previewCols = readOnly
    ? Math.min(MAX_PREVIEW_COLS, safeGrid[0]?.length || 0)
    : safeGrid[0]?.length || 0;
  /** Slices + merge filter must be memoized: new array refs every render forced HotTable to updateSettings in a tight loop (preview / dialog freeze). */
  const renderedGrid = React.useMemo(() => {
    if (!readOnly) return safeGrid;
    return safeGrid
      .slice(0, previewRows)
      .map((row) => (Array.isArray(row) ? row.slice(0, previewCols) : []));
  }, [readOnly, safeGrid, previewRows, previewCols]);

  const isPreviewTruncated =
    readOnly &&
    (safeGrid.length > previewRows || (safeGrid[0]?.length || 0) > previewCols);

  const renderedMergeCells = React.useMemo(
    () =>
      (activeSheet.mergeCells || []).filter(
        (m) =>
          m &&
          Number.isFinite(+m.row) &&
          Number.isFinite(+m.col) &&
          Number.isFinite(+m.rowspan) &&
          Number.isFinite(+m.colspan) &&
          m.row < previewRows &&
          m.col < previewCols &&
          m.row + m.rowspan <= previewRows &&
          m.col + m.colspan <= previewCols,
      ),
    [activeSheet.mergeCells, previewRows, previewCols],
  );

  const renderedColWidths = React.useMemo(() => {
    if (!readOnly) return activeSheet.colWidthsPx;
    return (activeSheet.colWidthsPx || []).slice(0, previewCols);
  }, [readOnly, activeSheet.colWidthsPx, previewCols]);

  const renderedRowHeights = React.useMemo(() => {
    if (!readOnly) return activeSheet.rowHeightsPx;
    return (activeSheet.rowHeightsPx || []).slice(0, previewRows);
  }, [readOnly, activeSheet.rowHeightsPx, previewRows]);

  /** `stretchH="all"` ignores fixed `colWidthsPx`; only use it when the template has no saved widths. */
  const stretchColumnsInPreview =
    readOnly &&
    (!Array.isArray(activeSheet.colWidthsPx) ||
      activeSheet.colWidthsPx.length === 0);

  const currentCellCount = renderedGrid.reduce(
    (t, row) => t + (Array.isArray(row) ? row.length : 0),
    0,
  );
  // Keep formulas active in preview/runtime mode too, otherwise dependent cells
  // never recalculate when users edit fillable inputs.
  const shouldUseFormulaEngine = currentCellCount <= 20000;

  const imageMap = React.useMemo(() => {
    const map = new Map<
      string,
      { dataUrl: string; rowspan: number; colspan: number }
    >();
    for (const img of (activeSheet as any)?.images || []) {
      if (!img?.dataUrl) continue;
      map.set(cellCoordKey(+img.row, +img.col), {
        dataUrl: img.dataUrl,
        rowspan: Math.max(1, +img.rowspan || 1),
        colspan: Math.max(1, +img.colspan || 1),
      });
    }
    return map;
  }, [activeSheet]);

  const persistedCellMetaMap = React.useMemo(() => {
    const map = new Map<string, NonNullable<SheetData["cellMeta"]>[number]>();
    for (const meta of activeSheet?.cellMeta || [])
      map.set(cellCoordKey(+meta.row, +meta.col), meta);
    return map;
  }, [activeSheet]);

  const fillableCellSet = React.useMemo(() => {
    const set = new Set<string>();
    if (!readOnly) return set;
    for (const [key, meta] of persistedCellMetaMap) {
      if (classNameHasFillable(meta.className)) set.add(key);
    }
    for (const m of renderedMergeCells) {
      if (!m) continue;
      let mergeHasFillable = false;
      for (
        let r = m.row;
        r <= m.row + m.rowspan - 1 && !mergeHasFillable;
        r++
      ) {
        for (
          let c = m.col;
          c <= m.col + m.colspan - 1 && !mergeHasFillable;
          c++
        ) {
          const meta = persistedCellMetaMap.get(cellCoordKey(r, c));
          if (meta && classNameHasFillable(meta.className))
            mergeHasFillable = true;
        }
      }
      if (mergeHasFillable) {
        for (let r = m.row; r <= m.row + m.rowspan - 1; r++) {
          for (let c = m.col; c <= m.col + m.colspan - 1; c++) {
            set.add(cellCoordKey(r, c));
          }
        }
      }
    }
    return set;
  }, [readOnly, persistedCellMetaMap, renderedMergeCells]);

  // ─── selection helpers ──────────────────────────────────────────────────────

  /**
   * Snapshot selection from HOT at action time (getSelectedRangeLast first).
   * Use this in toolbar handlers instead of reading selection after focus moved.
   */
  const getToolbarActionRange = React.useCallback(
    (hot: any) => {
      if (!hot) return null;
      const idx = activeSheetIndexRef.current;
      const sheet = workbookRef.current.sheets[idx];
      const rowCount = Math.max(1, sheet?.grid?.length || 1);
      const colCount = Math.max(1, sheet?.grid?.[0]?.length || 1);

      const clamp = (range: {
        startRow: number;
        endRow: number;
        startCol: number;
        endCol: number;
      }) => ({
        startRow: Math.max(
          0,
          Math.min(rowCount - 1, Math.min(range.startRow, range.endRow)),
        ),
        endRow: Math.max(
          0,
          Math.min(rowCount - 1, Math.max(range.startRow, range.endRow)),
        ),
        startCol: Math.max(
          0,
          Math.min(colCount - 1, Math.min(range.startCol, range.endCol)),
        ),
        endCol: Math.max(
          0,
          Math.min(colCount - 1, Math.max(range.startCol, range.endCol)),
        ),
      });

      if (formatAllCells) {
        return clamp({
          startRow: 0,
          endRow: rowCount - 1,
          startCol: 0,
          endCol: colCount - 1,
        });
      }

      const sel =
        typeof hot.getSelectedRangeLast === "function"
          ? hot.getSelectedRangeLast()
          : null;
      if (sel?.from != null && sel?.to != null) {
        const r = clamp({
          startRow: sel.from.row,
          endRow: sel.to.row,
          startCol: sel.from.col,
          endCol: sel.to.col,
        });
        lastSelectionRef.current = r;
        sheetSelectionRef.current[idx] = r;
        return r;
      }

      const last = hot.getSelectedLast?.();
      if (
        last &&
        last.length >= 4 &&
        last.every((v: any) => Number.isInteger(v))
      ) {
        const r = clamp({
          startRow: last[0],
          endRow: last[2],
          startCol: last[1],
          endCol: last[3],
        });
        lastSelectionRef.current = r;
        sheetSelectionRef.current[idx] = r;
        return r;
      }

      const cached = sheetSelectionRef.current[idx] ?? lastSelectionRef.current;
      return clamp(cached);
    },
    [formatAllCells],
  );

  const restoreHotRange = (
    hot: any,
    range: {
      startRow: number;
      endRow: number;
      startCol: number;
      endCol: number;
    } | null,
  ) => {
    if (!hot || !range) return;
    hot.render();
    hot.selectCell(
      range.startRow,
      range.startCol,
      range.endRow,
      range.endCol,
      false,
      false,
    );
  };

  const refreshUndoRedoState = React.useCallback(() => {
    const hot = hotRef.current?.hotInstance;
    const ur = hot?.getPlugin?.("undoRedo");
    setCanUndo(Boolean(ur?.isUndoAvailable?.()));
    setCanRedo(Boolean(ur?.isRedoAvailable?.()));
  }, []);

  const syncToolbarFromCell = React.useCallback(
    (hot: any, row: number, col: number) => {
      const v = hot.getDataAtCell(row, col);
      setFormulaInput(v == null ? "" : String(v));

      const cls = String(hot.getCellMeta(row, col)?.className || "");
      const tokens = cls.split(" ").filter(Boolean);

      const find = (prefix: string) =>
        tokens.find((t: string) => t.startsWith(prefix));

      setIsBoldActive(tokens.includes("meta-bold"));
      setIsItalicActive(tokens.includes("meta-italic"));
      setIsUnderlineActive(tokens.includes("meta-underline"));
      setIsStrikeActive(tokens.includes("meta-strike"));

      const fontToken = find("meta-font-");
      const sizeToken = find("meta-size-");
      const colorToken = find("meta-color-");
      const fillToken = find("meta-fill-");

      if (fontToken)
        setFontFamily(fontToken.replace("meta-font-", "").replace(/_/g, " "));
      if (sizeToken) setFontSize(sizeToken.replace("meta-size-", ""));
      setTextColor(
        colorToken ? `#${colorToken.replace("meta-color-", "")}` : "#111827",
      );
      setFillColor(
        fillToken ? `#${fillToken.replace("meta-fill-", "")}` : "#ffffff",
      );

      if (tokens.includes("meta-align-left")) setSelectedAlign("left");
      else if (tokens.includes("meta-align-center")) setSelectedAlign("center");
      else if (tokens.includes("meta-align-right")) setSelectedAlign("right");
      else if (tokens.includes("meta-align-justify"))
        setSelectedAlign("justify");
      else setSelectedAlign(null);

      if (tokens.includes("meta-valign-top")) setSelectedVAlign("top");
      else if (tokens.includes("meta-valign-middle"))
        setSelectedVAlign("middle");
      else if (tokens.includes("meta-valign-bottom"))
        setSelectedVAlign("bottom");
      else setSelectedVAlign(null);
    },
    [],
  );

  // ─── sheet sync ────────────────────────────────────────────────────────────

  const collectCurrentSheetFromHot = React.useCallback(
    (includeMeta: boolean, sheetIndex?: number) => {
      const hot = hotRef.current?.hotInstance;
      if (!hot) return;

      const idx = sheetIndex ?? activeSheetIndexRef.current;

      // Persist source values (including formula expressions like "=A1+B1"),
      // not rendered/calculated values, so reopened templates can recalculate.
      const sourceGrid =
        (typeof hot.getSourceDataArray === "function"
          ? hot.getSourceDataArray()
          : typeof hot.getSourceData === "function"
            ? hot.getSourceData()
            : hot.getData?.()) || [];
      const nextGrid = (sourceGrid as any[]).map((row: any[]) =>
        (Array.isArray(row) ? row : []).map((cell) =>
          cell == null ? "" : String(cell),
        ),
      );

      const mergeCells =
        hot
          ?.getPlugin?.("mergeCells")
          ?.mergedCellsCollection?.mergedCells?.map((cell: any) => ({
            row: cell.row,
            col: cell.col,
            rowspan: cell.rowspan,
            colspan: cell.colspan,
          })) || [];

      let cellMeta = workbookRef.current.sheets[idx]?.cellMeta || [];
      if (includeMeta) {
        // HOT's getCellsMeta() only returns *lazy-initialized* meta objects. Replacing the
        // full persisted `cellMeta` with that list drops formatting for cells the table has
        // never touched (e.g. other fillable highlights after marking a new region).
        const metaByKey = new Map<string, CellMetaEntry>();
        for (const m of cellMeta) {
          if (!m || !Number.isFinite(+m.row) || !Number.isFinite(+m.col))
            continue;
          metaByKey.set(cellCoordKey(+m.row, +m.col), {
            ...m,
            row: +m.row,
            col: +m.col,
          });
        }

        const cellsMeta =
          typeof hot.getCellsMeta === "function" ? hot.getCellsMeta() : [];
        for (const meta of cellsMeta || []) {
          if (
            typeof meta?.row !== "number" ||
            typeof meta?.col !== "number" ||
            meta.row < 0 ||
            meta.col < 0
          )
            continue;

          const useful =
            Boolean(meta?.className) ||
            Boolean(meta?.type) ||
            Boolean(meta?.dateFormat) ||
            typeof meta?.correctFormat === "boolean" ||
            Boolean(meta?.numericFormat) ||
            Array.isArray(meta?.source) ||
            typeof meta?.strict === "boolean";
          if (!useful) continue;

          const key = cellCoordKey(meta.row, meta.col);
          const existing = metaByKey.get(key);

          // Preserve persisted class tokens (e.g. `meta-fillable`) when HOT's
          // lazily initialized meta for this cell has no className.
          const existingTokens = String(existing?.className || "")
            .split(/\s+/)
            .filter(Boolean);
          const hotTokens = String(meta?.className || "")
            .split(/\s+/)
            .filter(Boolean);
          const mergedTokens = new Set([...hotTokens]);
          for (const t of existingTokens) {
            if (!mergedTokens.has(t)) mergedTokens.add(t);
          }
          const mergedClassName = [...mergedTokens].join(" ").trim() || undefined;

          metaByKey.set(key, {
            row: meta.row,
            col: meta.col,
            className: mergedClassName,
            type: meta.type ? String(meta.type) : existing?.type,
            dateFormat: meta.dateFormat
              ? String(meta.dateFormat)
              : existing?.dateFormat,
            correctFormat:
              typeof meta.correctFormat === "boolean"
                ? meta.correctFormat
                : existing?.correctFormat,
            numericFormat:
              meta.numericFormat && typeof meta.numericFormat === "object"
                ? {
                    pattern:
                      typeof meta.numericFormat.pattern === "string"
                        ? meta.numericFormat.pattern
                        : existing?.numericFormat?.pattern,
                    culture:
                      typeof meta.numericFormat.culture === "string"
                        ? meta.numericFormat.culture
                        : existing?.numericFormat?.culture,
                  }
                : existing?.numericFormat,
            source: Array.isArray(meta.source)
              ? meta.source.map(String)
              : existing?.source,
            strict: typeof meta.strict === "boolean" ? meta.strict : existing?.strict,
          });
        }
        cellMeta = dedupeCellMetaByCoordinate([...metaByKey.values()]);
      }

      const current = workbookRef.current.sheets[idx] || {
        name: `Sheet${idx + 1}`,
        grid: [[""]],
      };

      const rowCount =
        typeof hot.countRows === "function" ? hot.countRows() : nextGrid.length;
      const colCount =
        typeof hot.countCols === "function"
          ? hot.countCols()
          : Math.max(
              1,
              ...nextGrid.map((row: any[]) =>
                Array.isArray(row) ? row.length : 0,
              ),
            );

      let colWidthsPx: number[];
      let rowHeightsPx: number[];
      if (!readOnly) {
        colWidthsPx = [];
        for (let c = 0; c < colCount; c++) {
          const w =
            typeof hot.getColWidth === "function"
              ? hot.getColWidth(c)
              : undefined;
          const rounded =
            typeof w === "number" && Number.isFinite(w) ? Math.round(w) : NaN;
          colWidthsPx.push(
            Number.isFinite(rounded)
              ? rounded
              : Math.round(Number(current.colWidthsPx?.[c]) || 50),
          );
        }

        rowHeightsPx = [];
        for (let r = 0; r < rowCount; r++) {
          const h =
            typeof hot.getRowHeight === "function"
              ? hot.getRowHeight(r)
              : undefined;
          const rounded =
            typeof h === "number" && Number.isFinite(h) ? Math.round(h) : NaN;
          rowHeightsPx.push(
            Number.isFinite(rounded)
              ? rounded
              : Math.round(Number(current.rowHeightsPx?.[r]) || 24),
          );
        }
      } else {
        colWidthsPx = current.colWidthsPx || [];
        rowHeightsPx = current.rowHeightsPx || [];
      }

      workbookRef.current.sheets[idx] = deepCloneSheet({
        ...current,
        grid: nextGrid,
        mergeCells,
        cellMeta,
        colWidthsPx,
        rowHeightsPx,
      });
    },
    [readOnly],
  );

  React.useImperativeHandle(
    ref,
    () => ({
      getWorkbookSnapshot: () => {
        if (!readOnly) collectCurrentSheetFromHot(true);
        return {
          sheets: workbookRef.current.sheets.map(deepCloneSheet),
        };
      },
    }),
    [readOnly, collectCurrentSheetFromHot],
  );

  const emitWorkbookToParent = React.useCallback(() => {
    const snapshot = { sheets: workbookRef.current.sheets.map(deepCloneSheet) };
    lastIncomingSignatureRef.current = workbookSignature(snapshot.sheets);
    onChange(snapshot);
  }, [onChange]);

  React.useEffect(() => {
    if (readOnly) return;
    const onBeforeUnload = () => {
      collectCurrentSheetFromHot(true);
      emitWorkbookToParent();
    };
    window.addEventListener("beforeunload", onBeforeUnload);
    return () => window.removeEventListener("beforeunload", onBeforeUnload);
  }, [readOnly, collectCurrentSheetFromHot, emitWorkbookToParent]);

  const toVisibleGrid = React.useCallback(
    (sheet?: SheetData) => {
      const base = sheet?.grid?.length ? sheet.grid : [[""]];
      if (!readOnly) return cloneEditableGrid(base);
      const rows = Math.min(MAX_PREVIEW_ROWS, base.length);
      const cols = Math.min(MAX_PREVIEW_COLS, base[0]?.length || 0);
      return base
        .slice(0, rows)
        .map((row) => (Array.isArray(row) ? row.slice(0, cols) : []));
    },
    [readOnly],
  );

  const loadSheetIntoHot = React.useCallback(
    (targetIndex: number) => {
      const hot = hotRef.current?.hotInstance;
      if (!hot) return;
      const sheet = workbookRef.current.sheets[targetIndex];
      if (!sheet) return;
      const visibleGrid = toVisibleGrid(sheet);
      setInitialGrid(visibleGrid);
      hot.loadData(visibleGrid);
      if (!readOnly) {
        for (const meta of dedupeCellMetaByCoordinate(sheet.cellMeta || [])) {
          if (meta.className)
            hot.setCellMeta(meta.row, meta.col, "className", meta.className);
          if (meta.type) hot.setCellMeta(meta.row, meta.col, "type", meta.type);
          if (meta.dateFormat)
            hot.setCellMeta(meta.row, meta.col, "dateFormat", meta.dateFormat);
          if (typeof meta.correctFormat === "boolean")
            hot.setCellMeta(
              meta.row,
              meta.col,
              "correctFormat",
              meta.correctFormat,
            );
          if (meta.numericFormat)
            hot.setCellMeta(
              meta.row,
              meta.col,
              "numericFormat",
              meta.numericFormat,
            );
          if (Array.isArray(meta.source))
            hot.setCellMeta(meta.row, meta.col, "source", meta.source);
          if (typeof meta.strict === "boolean")
            hot.setCellMeta(meta.row, meta.col, "strict", meta.strict);
        }
      }
      hot.render();

      const rows = Math.max(1, sheet?.grid?.length || 1);
      const cols = Math.max(1, sheet?.grid?.[0]?.length || 1);
      const saved = sheetSelectionRef.current[targetIndex];
      const nextRange = saved
        ? {
            startRow: Math.max(0, Math.min(rows - 1, saved.startRow)),
            endRow: Math.max(0, Math.min(rows - 1, saved.endRow)),
            startCol: Math.max(0, Math.min(cols - 1, saved.startCol)),
            endCol: Math.max(0, Math.min(cols - 1, saved.endCol)),
          }
        : { startRow: 0, endRow: 0, startCol: 0, endCol: 0 };

      hot.selectCell(
        nextRange.startRow,
        nextRange.startCol,
        nextRange.endRow,
        nextRange.endCol,
        false,
        false,
      );
      lastSelectionRef.current = nextRange;
      sheetSelectionRef.current[targetIndex] = nextRange;
      setSelectionLabel(toRangeLabel(nextRange));
    },
    [readOnly, toVisibleGrid],
  );

  const handleSheetSwitch = (targetIndex: number) => {
    if (targetIndex === activeSheetIndex) return;
    if (!readOnly) {
      collectCurrentSheetFromHot(true, activeSheetIndex);
      emitWorkbookToParent();
    }
    const hot = hotRef.current?.hotInstance;
    if (hot) getToolbarActionRange(hot);
    setInitialGrid(toVisibleGrid(workbookRef.current.sheets[targetIndex]));
    const saved = sheetSelectionRef.current[targetIndex];
    if (saved) lastSelectionRef.current = saved;
    setSelectionLabel(toRangeLabel(lastSelectionRef.current));
    setActiveSheetIndex(targetIndex);
  };

  // ─── class-based formatting ─────────────────────────────────────────────────

  const applyClassToSelection = React.useCallback(
    (classToken: string, toggle = false, replacePrefix?: string) => {
      const hot = hotRef.current?.hotInstance;
      if (!hot || readOnly) return;
      const range = getToolbarActionRange(hot);
      if (!range) return;
      const prefix = replacePrefix || classToken;

      const apply = () => {
        for (let r = range.startRow; r <= range.endRow; r++) {
          for (let c = range.startCol; c <= range.endCol; c++) {
            const current = String(hot.getCellMeta(r, c)?.className || "")
              .split(" ")
              .filter(Boolean);
            const has = current.includes(classToken);
            const next = toggle
              ? has
                ? current.filter((x: string) => x !== classToken)
                : [...current, classToken]
              : [
                  ...current.filter((x: string) => !x.startsWith(prefix)),
                  classToken,
                ];
            hot.setCellMeta(r, c, "className", next.join(" ").trim());
          }
        }
      };

      if (typeof hot.batch === "function") hot.batch(apply);
      else apply();
      collectCurrentSheetFromHot(true);
      restoreHotRange(hot, range);
    },
    [getToolbarActionRange, readOnly, collectCurrentSheetFromHot],
  );

  const setFontStyle = (style: "bold" | "italic" | "underline" | "strike") => {
    applyClassToSelection(`meta-${style}`, true);
    if (style === "bold") setIsBoldActive((p) => !p);
    if (style === "italic") setIsItalicActive((p) => !p);
    if (style === "underline") setIsUnderlineActive((p) => !p);
    if (style === "strike") setIsStrikeActive((p) => !p);
  };

  const setAlignment = (align: "left" | "center" | "right" | "justify") => {
    applyClassToSelection(`meta-align-${align}`, false, "meta-align-");
    setSelectedAlign(align);
  };

  const setVerticalAlignment = (align: "top" | "middle" | "bottom") => {
    applyClassToSelection(`meta-valign-${align}`, false, "meta-valign-");
    setSelectedVAlign(align);
  };

  const setWrapText = () => applyClassToSelection("meta-wrap", true);

  const applyFontFamily = () =>
    applyClassToSelection(
      `meta-font-${fontFamily.replace(/\s+/g, "_")}`,
      false,
      "meta-font-",
    );
  const applyFontSize = () =>
    applyClassToSelection(`meta-size-${fontSize}`, false, "meta-size-");

  const flushPendingColorTimers = React.useCallback(() => {
    if (textColorApplyTimerRef.current) {
      clearTimeout(textColorApplyTimerRef.current);
      textColorApplyTimerRef.current = null;
    }
    if (fillColorApplyTimerRef.current) {
      clearTimeout(fillColorApplyTimerRef.current);
      fillColorApplyTimerRef.current = null;
    }
  }, []);

  React.useEffect(
    () => () => {
      flushPendingColorTimers();
    },
    [flushPendingColorTimers],
  );

  const scheduleApplyTextColorValue = React.useCallback(
    (hex: string) => {
      if (textColorApplyTimerRef.current)
        clearTimeout(textColorApplyTimerRef.current);
      textColorApplyTimerRef.current = setTimeout(() => {
        textColorApplyTimerRef.current = null;
        applyClassToSelection(`meta-color-${hex}`, false, "meta-color-");
      }, 150);
    },
    [applyClassToSelection],
  );

  const scheduleApplyFillColorValue = React.useCallback(
    (hex: string) => {
      if (fillColorApplyTimerRef.current)
        clearTimeout(fillColorApplyTimerRef.current);
      fillColorApplyTimerRef.current = setTimeout(() => {
        fillColorApplyTimerRef.current = null;
        applyClassToSelection(`meta-fill-${hex}`, false, "meta-fill-");
      }, 150);
    },
    [applyClassToSelection],
  );

  const applyTextColor = () => {
    flushPendingColorTimers();
    const hex = textColor.replace("#", "");
    applyClassToSelection(`meta-color-${hex}`, false, "meta-color-");
  };

  const applyFillColor = () => {
    flushPendingColorTimers();
    const hex = fillColor.replace("#", "");
    applyClassToSelection(`meta-fill-${hex}`, false, "meta-fill-");
  };

  // ─── numeric / date formatting ──────────────────────────────────────────────

  const formatSelectedAs = (
    kind: "number" | "currency" | "percent" | "date",
  ) => {
    const hot = hotRef.current?.hotInstance;
    if (!hot || readOnly) return;
    const range = getToolbarActionRange(hot);
    if (!range) return;

    const apply = () => {
      for (let r = range.startRow; r <= range.endRow; r++) {
        for (let c = range.startCol; c <= range.endCol; c++) {
          const raw = hot.getDataAtCell(r, c);

          if (kind === "date") {
            hot.setCellMeta(r, c, "type", "date");
            hot.setCellMeta(r, c, "dateFormat", "YYYY-MM-DD");
            hot.setCellMeta(r, c, "correctFormat", true);
            if (raw != null && raw !== "") {
              const d = new Date(raw);
              if (!isNaN(d.getTime()))
                hot.setDataAtCell(r, c, d.toISOString().slice(0, 10));
            }
            continue;
          }

          hot.setCellMeta(r, c, "type", "numeric");
          const patterns: Record<string, string> = {
            number: "0,0.00",
            currency: "$0,0.00",
            percent: "0.00%",
          };
          hot.setCellMeta(r, c, "numericFormat", {
            pattern: patterns[kind],
            culture: "en-US",
          });

          if (raw != null && raw !== "") {
            const rawText = String(raw).trim();
            let numeric = Number(rawText.replace(/[$,%\s,]/g, ""));
            if (kind === "percent" && rawText.includes("%")) numeric /= 100;
            if (!isNaN(numeric)) hot.setDataAtCell(r, c, numeric);
          }
        }
      }
    };

    if (typeof hot.batch === "function") hot.batch(apply);
    else apply();
    collectCurrentSheetFromHot(true);
    restoreHotRange(hot, range);
  };

  // ─── validation helpers ─────────────────────────────────────────────────────

  const applyDropdownValidation = () => {
    const hot = hotRef.current?.hotInstance;
    if (!hot || readOnly) return;
    const range = getToolbarActionRange(hot);
    if (!range) return;
    const source = dropdownSource
      .split(",")
      .map((v) => v.trim())
      .filter(Boolean);
    for (let r = range.startRow; r <= range.endRow; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        hot.setCellMeta(r, c, "type", "dropdown");
        hot.setCellMeta(r, c, "source", source);
        hot.setCellMeta(r, c, "strict", true);
      }
    }
    collectCurrentSheetFromHot(true);
    restoreHotRange(hot, range);
  };

  const applyDateCellType = () => {
    const hot = hotRef.current?.hotInstance;
    if (!hot || readOnly) return;
    const range = getToolbarActionRange(hot);
    if (!range) return;
    for (let r = range.startRow; r <= range.endRow; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        hot.setCellMeta(r, c, "type", "date");
        hot.setCellMeta(r, c, "dateFormat", "YYYY-MM-DD");
        hot.setCellMeta(r, c, "correctFormat", true);
      }
    }
    collectCurrentSheetFromHot(true);
    restoreHotRange(hot, range);
  };

  const toggleFillableSelection = () => {
    const hot = hotRef.current?.hotInstance;
    if (!hot || readOnly) return;

    const root = hot.rootElement as HTMLElement | undefined;
    const container =
      root?.closest('[role="dialog"]') ??
      root?.closest(
        ".overflow-y-auto, .overflow-auto, [data-radix-scroll-area-viewport]",
      ) ??
      document.documentElement;

    const savedScrollTop = container?.scrollTop ?? window.scrollY;
    const savedScrollLeft = container?.scrollLeft ?? window.scrollX;

    const idx = activeSheetIndexRef.current;
    collectCurrentSheetFromHot(true, idx);

    const range = getToolbarActionRange(hot);
    if (!range) return;

    const sheet = workbookRef.current.sheets[idx];
    if (!sheet) return;

    const metaByKey = new Map<
      string,
      NonNullable<SheetData["cellMeta"]>[number]
    >();
    for (const meta of sheet.cellMeta || []) {
      metaByKey.set(`${meta.row}:${meta.col}`, meta);
    }

    for (let r = range.startRow; r <= range.endRow; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        const key = `${r}:${c}`;
        const current = metaByKey.get(key) || { row: r, col: c };
        const tokens = String(current.className || "")
          .split(" ")
          .filter(Boolean);
        if (!tokens.includes("meta-fillable")) tokens.push("meta-fillable");
        const newClassName = tokens.join(" ").trim();

        hot.setCellMeta(r, c, "className", newClassName);

        metaByKey.set(key, {
          ...current,
          row: r,
          col: c,
          className: newClassName || undefined,
        });
      }
    }

    sheet.cellMeta = dedupeCellMetaByCoordinate(Array.from(metaByKey.values()));
    workbookRef.current.sheets[idx] = deepCloneSheet(sheet);

    lastIncomingSignatureRef.current = workbookSignature(
      workbookRef.current.sheets,
    );

    hot.render();

    hot.selectCell(
      range.startRow,
      range.startCol,
      range.endRow,
      range.endCol,
      false,
      false,
    );

    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        if (container && container !== document.documentElement) {
          container.scrollTop = savedScrollTop;
          container.scrollLeft = savedScrollLeft;
        } else {
          window.scrollTo(savedScrollLeft, savedScrollTop);
        }
      });
    });
  };

  // ─── sort / find-replace ────────────────────────────────────────────────────

  const sortSelectedColumn = (order: "asc" | "desc") => {
    const hot = hotRef.current?.hotInstance;
    if (!hot) return;
    const range = getToolbarActionRange(hot);
    if (!range) return;
    const sorting = hot.getPlugin?.("columnSorting");
    if (typeof sorting?.sort === "function")
      sorting.sort({ column: range.startCol, sortOrder: order });
    restoreHotRange(hot, range);
  };

  const doFindReplace = () => {
    const hot = hotRef.current?.hotInstance;
    if (!hot || !findValue || readOnly) return;
    const range = getToolbarActionRange(hot);
    const data = hot.getData();
    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        const val = String(data[r][c] ?? "");
        if (val.includes(findValue))
          hot.setDataAtCell(r, c, val.split(findValue).join(replaceValue));
      }
    }
    collectCurrentSheetFromHot(false);
    if (range) restoreHotRange(hot, range);
  };

  // ─── undo / redo ────────────────────────────────────────────────────────────

  const undoAction = () => {
    const hot = hotRef.current?.hotInstance;
    const ur = hot?.getPlugin?.("undoRedo");
    if (!ur?.undo || readOnly) return;
    const r =
      sheetSelectionRef.current[activeSheetIndexRef.current] ??
      lastSelectionRef.current;
    ur.undo();
    refreshUndoRedoState();
    if (hot && r) restoreHotRange(hot, r);
  };

  const redoAction = () => {
    const hot = hotRef.current?.hotInstance;
    const ur = hot?.getPlugin?.("undoRedo");
    if (!ur?.redo || readOnly) return;
    const r =
      sheetSelectionRef.current[activeSheetIndexRef.current] ??
      lastSelectionRef.current;
    ur.redo();
    refreshUndoRedoState();
    if (hot && r) restoreHotRange(hot, r);
  };

  // ─── merge ──────────────────────────────────────────────────────────────────

  const mergeSelection = () => {
    const hot = hotRef.current?.hotInstance;
    if (!hot || readOnly) return;
    const range = getToolbarActionRange(hot);
    if (!range) return;
    const plugin = hot.getPlugin?.("mergeCells");
    if (!plugin) return;
    plugin.merge(range.startRow, range.startCol, range.endRow, range.endCol);
    collectCurrentSheetFromHot(true);
    restoreHotRange(hot, range);
  };

  const unmergeSelection = () => {
    const hot = hotRef.current?.hotInstance;
    if (!hot || readOnly) return;
    const range = getToolbarActionRange(hot);
    if (!range) return;
    const plugin = hot.getPlugin?.("mergeCells");
    if (!plugin) return;
    plugin.unmerge(range.startRow, range.startCol, range.endRow, range.endCol);
    collectCurrentSheetFromHot(true);
    restoreHotRange(hot, range);
  };

  // ─── row / col operations ───────────────────────────────────────────────────

  const alterBySelection = (
    kind:
      | "insert_row_above"
      | "insert_row_below"
      | "insert_col_start"
      | "insert_col_end"
      | "remove_row"
      | "remove_col",
  ) => {
    const hot = hotRef.current?.hotInstance;
    if (!hot || readOnly) return;
    const range = getToolbarActionRange(hot);
    if (!range) return;
    if (kind === "insert_row_above")
      hot.alter("insert_row_above", range.startRow, 1);
    if (kind === "insert_row_below")
      hot.alter("insert_row_below", range.endRow, 1);
    if (kind === "insert_col_start")
      hot.alter("insert_col_start", range.startCol, 1);
    if (kind === "insert_col_end") hot.alter("insert_col_end", range.endCol, 1);
    if (kind === "remove_row")
      hot.alter(
        "remove_row",
        range.startRow,
        range.endRow - range.startRow + 1,
      );
    if (kind === "remove_col")
      hot.alter(
        "remove_col",
        range.startCol,
        range.endCol - range.startCol + 1,
      );
    collectCurrentSheetFromHot(true);
    refreshUndoRedoState();
    restoreHotRange(hot, range);
  };

  const clearSelectionValues = () => {
    const hot = hotRef.current?.hotInstance;
    if (!hot || readOnly) return;
    const range = getToolbarActionRange(hot);
    if (!range) return;
    const updates: [number, number, string][] = [];
    for (let r = range.startRow; r <= range.endRow; r++)
      for (let c = range.startCol; c <= range.endCol; c++)
        updates.push([r, c, ""]);
    hot.setDataAtCell(updates);
    collectCurrentSheetFromHot(false);
    refreshUndoRedoState();
    restoreHotRange(hot, range);
  };

  const clearSelectionFormatting = () => {
    const hot = hotRef.current?.hotInstance;
    if (!hot || readOnly) return;
    const range = getToolbarActionRange(hot);
    if (!range) return;
    const apply = () => {
      for (let r = range.startRow; r <= range.endRow; r++) {
        for (let c = range.startCol; c <= range.endCol; c++) {
          const cls = String(hot.getCellMeta(r, c)?.className || "")
            .split(" ")
            .filter(
              (t: string) => t === "meta-fillable" || !t.startsWith("meta-"),
            );
          hot.setCellMeta(r, c, "className", cls.join(" ").trim());
          hot.setCellMeta(r, c, "type", undefined as any);
          hot.setCellMeta(r, c, "numericFormat", undefined as any);
          hot.setCellMeta(r, c, "dateFormat", undefined as any);
          hot.setCellMeta(r, c, "correctFormat", undefined as any);
          hot.setCellMeta(r, c, "source", undefined as any);
          hot.setCellMeta(r, c, "strict", undefined as any);
        }
      }
    };
    if (typeof hot.batch === "function") hot.batch(apply);
    else apply();
    collectCurrentSheetFromHot(true);
    restoreHotRange(hot, range);
  };

  // ─── formula bar ────────────────────────────────────────────────────────────

  const applyFormulaBar = () => {
    const hot = hotRef.current?.hotInstance;
    if (!hot || readOnly) return;
    const range = getToolbarActionRange(hot);
    if (!range) return;
    hot.setDataAtCell(range.startRow, range.startCol, formulaInput);
    collectCurrentSheetFromHot(false);
    restoreHotRange(hot, range);
  };

  // ─── export ─────────────────────────────────────────────────────────────────

  const emitWorkbookSnapshot = () => {
    if (!readOnly) collectCurrentSheetFromHot(true);
    emitWorkbookToParent();
  };

  /** Persist column/row sizes into `workbookRef` (parent gets them on Save / tab switch / unload). */
  const flushLayoutToParent = React.useCallback(() => {
    if (readOnly) return;
    collectCurrentSheetFromHot(true);
  }, [readOnly, collectCurrentSheetFromHot]);

  const exportXlsx = async () => {
    if (!readOnly) collectCurrentSheetFromHot(true);
    const workbook = new ExcelJS.Workbook();
    workbookRef.current.sheets.forEach((sheet) => {
      const ws = workbook.addWorksheet(sheet.name || "Sheet");
      sheet.grid.forEach((row) => ws.addRow(row));
    });
    const buf = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buf], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "workbook.xlsx";
    a.click();
    URL.revokeObjectURL(url);
  };

  const exportCsv = () => {
    if (!readOnly) collectCurrentSheetFromHot(true);
    const sheets = workbookRef.current.sheets;
    const idx = Math.min(
      Math.max(0, activeSheetIndexRef.current),
      Math.max(0, sheets.length - 1),
    );
    const sh = sheets[idx];
    const grid =
      Array.isArray(sh?.grid) && sh.grid.length > 0 ? sh.grid : [[""]];
    const csv = grid
      .map((row) =>
        row.map((v) => `"${String(v ?? "").replace(/"/g, '""')}"`).join(","),
      )
      .join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${sh?.name || "sheet"}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  // ─── sheet management ────────────────────────────────────────────────────────

  const duplicateActiveSheet = () => {
    if (!readOnly) collectCurrentSheetFromHot(true);
    const idx = activeSheetIndexRef.current;
    const source = workbookRef.current.sheets[idx];
    if (!source) return;
    const cloned = JSON.parse(JSON.stringify(source)) as SheetData;
    cloned.name = `${source.name} Copy`;
    const nextSheets = [...workbookRef.current.sheets, cloned];
    workbookRef.current.sheets = nextSheets;
    setSheetTabs(
      nextSheets.map((s) => ({ name: s.name, tabColor: s.tabColor })),
    );
    setInitialGrid(toVisibleGrid(nextSheets[nextSheets.length - 1]));
    setActiveSheetIndex(nextSheets.length - 1);
  };

  const moveSheet = (direction: "left" | "right") => {
    if (!readOnly) collectCurrentSheetFromHot(true);
    const sheets = workbookRef.current.sheets;
    const target = activeSheetIndex + (direction === "left" ? -1 : 1);
    if (target < 0 || target >= sheets.length) return;
    const next = [...sheets];
    const [moved] = next.splice(activeSheetIndex, 1);
    next.splice(target, 0, moved);
    workbookRef.current.sheets = next;
    setSheetTabs(next.map((s) => ({ name: s.name, tabColor: s.tabColor })));
    setInitialGrid(toVisibleGrid(next[target]));
    setActiveSheetIndex(target);
  };

  const applySheetColor = (color: string) => {
    const nextSheets = safeSheets.map((s, i) =>
      i === activeSheetIndex ? { ...s, tabColor: color } : s,
    );
    workbookRef.current.sheets = nextSheets;
    setSheetTabs(
      nextSheets.map((s) => ({ name: s.name, tabColor: s.tabColor })),
    );
  };

  // ─── effects ─────────────────────────────────────────────────────────────────

  React.useEffect(() => {
    const nextSheets = normalizedIncomingSheets;
    const sig = workbookSignature(nextSheets);
    if (sig === lastIncomingSignatureRef.current) return;

    const hot = hotRef.current?.hotInstance;
    if (hot && typeof hot.isEditorOpened === "function" && hot.isEditorOpened())
      return;

    lastIncomingSignatureRef.current = sig;

    const prevSheetCount = workbookRef.current.sheets.length;
    const nextSheetCount = nextSheets.length;
    const sheetCountChanged = prevSheetCount !== nextSheetCount;

    const prevSheets = workbookRef.current.sheets;
    workbookRef.current = {
      sheets: nextSheets.map((inc, i) =>
        mergeFillableMetaFromPrevSheet(prevSheets[i], inc),
      ),
    };
    setSheetTabs(
      nextSheets.map((s) => ({ name: s.name, tabColor: s.tabColor })),
    );

    // Do not jump to sheet 0 when the parent only reflects an in-place edit
    // (cell values, cellMeta, merges, etc.): that caused "Mark Fillable" to
    // reset the tab and briefly bind sheet 0's grid to the wrong context.
    if (sheetCountChanged) {
      setActiveSheetIndex(0);
      const first = nextSheets[0]?.grid?.length ? nextSheets[0].grid : [[""]];
      if (!readOnly) {
        setInitialGrid(cloneEditableGrid(first));
      } else {
        const rows = Math.min(MAX_PREVIEW_ROWS, first.length);
        const cols = Math.min(MAX_PREVIEW_COLS, first[0]?.length || 0);
        setInitialGrid(
          first
            .slice(0, rows)
            .map((row) => (Array.isArray(row) ? row.slice(0, cols) : [])),
        );
      }
    } else {
      setActiveSheetIndex((prev) =>
        Math.min(prev, Math.max(0, nextSheets.length - 1)),
      );
      // Let loadSheetIntoHot (incomingWorkbookKey) refresh HOT for the
      // preserved tab; avoid forcing sheet 0's grid into state here.
    }
  }, [normalizedIncomingSheets, readOnly]);

  const incomingWorkbookKey = React.useMemo(
    () => workbookSignature(normalizedIncomingSheets),
    [normalizedIncomingSheets],
  );

  const hotTableMountKey = React.useMemo(
    () => hotTableMountSignature(normalizedIncomingSheets),
    [normalizedIncomingSheets],
  );

  React.useEffect(() => {
    loadSheetIntoHot(activeSheetIndex);
  }, [activeSheetIndex, loadSheetIntoHot, incomingWorkbookKey]);

  // ─── cell renderer ───────────────────────────────────────────────────────────

  const cellsCallback = React.useCallback(
    (row: number, col: number) => {
      const persistedMeta = persistedCellMetaMap.get(cellCoordKey(row, col));
      const cp: any = {};
      const persistedClassName = String(persistedMeta?.className || "");
      const classTokens = persistedClassName.split(" ").filter(Boolean);
      const isFillable = readOnly
        ? fillableCellSet.has(cellCoordKey(row, col))
        : classTokens.includes("meta-fillable");
      cp.readOnly = readOnly ? !isFillable : false;
      if (persistedMeta?.className) cp.className = persistedMeta.className;
      if (persistedMeta?.type) cp.type = persistedMeta.type;
      if (persistedMeta?.dateFormat) cp.dateFormat = persistedMeta.dateFormat;
      if (typeof persistedMeta?.correctFormat === "boolean")
        cp.correctFormat = persistedMeta.correctFormat;
      if (persistedMeta?.numericFormat)
        cp.numericFormat = persistedMeta.numericFormat;
      if (Array.isArray(persistedMeta?.source))
        cp.source = persistedMeta.source;
      if (typeof persistedMeta?.strict === "boolean")
        cp.strict = persistedMeta.strict;

      const image = imageMap.get(cellCoordKey(row, col));
      if (image || !readOnly) {
        cp.renderer = (
          instance: any,
          td: HTMLTableCellElement,
          rowIndex: number,
          colIndex: number,
          prop: any,
          value: any,
          cellProperties: any,
        ) => {
          const base = (window as any).Handsontable?.renderers?.TextRenderer;
          if (base)
            base(instance, td, rowIndex, colIndex, prop, value, cellProperties);
          else td.textContent = value == null ? "" : String(value);

          const cls = String(cellProperties?.className || "");
          const tokens = cls.split(" ").filter(Boolean);

          // Reset recycled cell styles
          const s = td.style;
          s.fontWeight =
            s.fontStyle =
            s.textDecoration =
            s.fontFamily =
            s.fontSize =
              "";
          s.color =
            s.backgroundColor =
            s.textAlign =
            s.verticalAlign =
            s.whiteSpace =
              "";

          if (tokens.includes("meta-bold")) s.fontWeight = "700";
          if (tokens.includes("meta-italic")) s.fontStyle = "italic";

          const decorations = [
            tokens.includes("meta-underline") ? "underline" : "",
            tokens.includes("meta-strike") ? "line-through" : "",
          ].filter(Boolean);
          if (decorations.length) s.textDecoration = decorations.join(" ");

          const fontToken = tokens.find((t: string) =>
            t.startsWith("meta-font-"),
          );
          const sizeToken = tokens.find((t: string) =>
            t.startsWith("meta-size-"),
          );
          const colorToken = tokens.find((t: string) =>
            t.startsWith("meta-color-"),
          );
          const fillToken = tokens.find((t: string) =>
            t.startsWith("meta-fill-"),
          );
          const alignToken = tokens.find((t: string) =>
            t.startsWith("meta-align-"),
          );
          const vAlignToken = tokens.find((t: string) =>
            t.startsWith("meta-valign-"),
          );

          if (fontToken)
            s.fontFamily = fontToken
              .replace("meta-font-", "")
              .replace(/_/g, " ");
          if (sizeToken)
            s.fontSize = `${sizeToken.replace("meta-size-", "")}px`;
          if (colorToken) s.color = `#${colorToken.replace("meta-color-", "")}`;
          if (fillToken)
            s.backgroundColor = `#${fillToken.replace("meta-fill-", "")}`;
          if (alignToken) s.textAlign = alignToken.replace("meta-align-", "");
          if (vAlignToken)
            s.verticalAlign = vAlignToken.replace("meta-valign-", "");
          if (tokens.includes("meta-wrap")) s.whiteSpace = "normal";

          if (image?.dataUrl) {
            const colWidths = Array.isArray(renderedColWidths)
              ? renderedColWidths
              : [];
            const rowHeights = Array.isArray(renderedRowHeights)
              ? renderedRowHeights
              : [];
            let imgW = 0;
            for (let cx = colIndex; cx < colIndex + image.colspan; cx++)
              imgW += Number(colWidths[cx] || 80);
            let imgH = 0;
            for (let rx = rowIndex; rx < rowIndex + image.rowspan; rx++)
              imgH += Number(rowHeights[rx] || 24);
            td.style.padding = "0";
            td.style.position = "relative";
            td.style.overflow = "visible";
            td.textContent = "";
            const img = document.createElement("img");
            img.src = image.dataUrl;
            img.style.cssText = `position:absolute;left:0;top:0;width:${Math.max(16, imgW)}px;height:${Math.max(16, imgH)}px;object-fit:fill;display:block;pointer-events:none;z-index:3`;
            td.appendChild(img);
          }
          return td;
        };
      }
      return cp;
    },
    [
      persistedCellMetaMap,
      fillableCellSet,
      imageMap,
      readOnly,
      renderedColWidths,
      renderedRowHeights,
    ],
  );

  /**
   * Runs after other plugins' `afterGetCellMeta` logic so template editing
   * (`readOnly={false}`) always stays fully editable (e.g. formula engine meta).
   */
  const afterGetCellMeta = React.useCallback(
    (row: number, col: number, cellProps: Record<string, unknown>) => {
      if (!readOnly) {
        (cellProps as { readOnly?: boolean }).readOnly = false;
        return;
      }
      // Formula cells must stay writable for HyperFormula recalculation updates.
      const hot = hotRef.current?.hotInstance;
      const cellValue = hot?.getSourceDataAtCell?.(row, col);
      if (typeof cellValue === "string" && cellValue.startsWith("=")) {
        (cellProps as { readOnly?: boolean }).readOnly = false;
        return;
      }
      // Do not trust `cellProps.className` alone: in preview we skip `setCellMeta` in
      // `loadSheetIntoHot`, so HOT often has no `meta-fillable` on the meta layer even when
      // persisted `cellMeta` does — that left `readOnly` stuck true and blocked all typing.
      const isFillable = fillableCellSet.has(cellCoordKey(row, col));
      (cellProps as { readOnly?: boolean }).readOnly = !isFillable;
    },
    [readOnly, fillableCellSet],
  );

  // ─── Toolbar button wrapper — prevents focus loss ─────────────────────────

  const TB = ({
    onClick,
    children,
    title,
    variant = "outline",
    disabled = false,
    active = false,
    className = "",
  }: {
    onClick: () => void;
    children: React.ReactNode;
    title?: string;
    variant?: "outline" | "default";
    disabled?: boolean;
    active?: boolean;
    className?: string;
  }) => (
    <Button
      type="button"
      size="sm"
      variant={active ? "default" : variant}
      disabled={disabled}
      title={title}
      className={className}
      onMouseDown={noFocusSteal} // ← THE KEY FIX: keeps grid focused
      onClick={onClick}
    >
      {children}
    </Button>
  );

  // ─── render ──────────────────────────────────────────────────────────────────

  return (
    <div
      className="space-y-2"
      onBlur={(e) => {
        if (!readOnly) return;
        const next = e.relatedTarget as Node | null;
        if (next && e.currentTarget.contains(next)) return;
        if (!readOnlyPreviewDirtyRef.current) return;
        readOnlyPreviewDirtyRef.current = false;
        emitWorkbookToParent();
      }}
    >
      <style>{`
        .meta-bold { font-weight: 700 !important; }
        .meta-italic { font-style: italic !important; }
        .meta-underline { text-decoration: underline !important; }
        .meta-strike { text-decoration: line-through !important; }
        .meta-wrap { white-space: normal !important; }
        .meta-align-left { text-align: left !important; }
        .meta-align-center { text-align: center !important; }
        .meta-align-right { text-align: right !important; }
        .meta-align-justify { text-align: justify !important; }
        .meta-valign-top { vertical-align: top !important; }
        .meta-valign-middle { vertical-align: middle !important; }
        .meta-valign-bottom { vertical-align: bottom !important; }
        .meta-fillable { background-color: #fffbe6 !important; }
      `}</style>

      {/* ── Toolbar ── */}
      {!readOnly && (
        <div className="relative z-10 flex flex-wrap items-center gap-1 p-2 border rounded-md bg-slate-50">
          <span
            className="px-2 text-xs font-medium border rounded bg-white min-w-[3rem] text-center"
            title="Active selection"
            onMouseDown={noFocusSteal}
          >
            {selectionLabel}
          </span>

          <TB onClick={undoAction} disabled={!canUndo}>
            Undo
          </TB>
          <TB onClick={redoAction} disabled={!canRedo}>
            Redo
          </TB>

          <span className="mx-1 h-6 border-l" />

          <TB onClick={() => setFontStyle("bold")} active={isBoldActive}>
            <b>B</b>
          </TB>
          <TB onClick={() => setFontStyle("italic")} active={isItalicActive}>
            <i>I</i>
          </TB>
          <TB
            onClick={() => setFontStyle("underline")}
            active={isUnderlineActive}
          >
            <u>U</u>
          </TB>
          <TB onClick={() => setFontStyle("strike")} active={isStrikeActive}>
            <s>S</s>
          </TB>

          <select
            value={fontFamily}
            onChange={(e) => {
              const v = e.target.value;
              setFontFamily(v);
              applyClassToSelection(
                `meta-font-${v.replace(/\s+/g, "_")}`,
                false,
                "meta-font-",
              );
            }}
            onBlur={() => {
              const hot = hotRef.current?.hotInstance;
              const idx = activeSheetIndexRef.current;
              const r =
                sheetSelectionRef.current[idx] ?? lastSelectionRef.current;
              if (hot)
                hot.selectCell(
                  r.startRow,
                  r.startCol,
                  r.endRow,
                  r.endCol,
                  false,
                  false,
                );
            }}
            onMouseDown={noFocusSteal}
            className="h-8 px-2 text-sm border rounded"
          >
            <option>Arial</option>
            <option>Calibri</option>
            <option>Times New Roman</option>
            <option>Verdana</option>
            <option>Courier New</option>
            <option>Georgia</option>
          </select>

          <input
            value={fontSize}
            onChange={(e) => {
              const v = e.target.value;
              setFontSize(v);
              applyClassToSelection(`meta-size-${v}`, false, "meta-size-");
            }}
            onBlur={() => {
              const hot = hotRef.current?.hotInstance;
              const idx = activeSheetIndexRef.current;
              const r =
                sheetSelectionRef.current[idx] ?? lastSelectionRef.current;
              if (hot)
                hot.selectCell(
                  r.startRow,
                  r.startCol,
                  r.endRow,
                  r.endCol,
                  false,
                  false,
                );
            }}
            onMouseDown={noFocusSteal}
            className="w-14 h-8 px-2 text-sm border rounded"
          />
          <TB onClick={applyFontFamily}>Font</TB>
          <TB onClick={applyFontSize}>Size</TB>

          <span className="mx-1 h-6 border-l" />

          <div className="flex items-center gap-1">
            <input
              type="color"
              value={textColor}
              onChange={(e) => {
                const val = e.target.value;
                setTextColor(val);
                scheduleApplyTextColorValue(val.replace("#", ""));
              }}
              onMouseDown={noFocusSteal}
              className="w-8 h-8 p-0 border rounded cursor-pointer"
              title="Text color"
            />
            <TB onClick={applyTextColor} title="Apply text color">
              A
            </TB>
          </div>
          <div className="flex items-center gap-1">
            <input
              type="color"
              value={fillColor}
              onChange={(e) => {
                const val = e.target.value;
                setFillColor(val);
                scheduleApplyFillColorValue(val.replace("#", ""));
              }}
              onMouseDown={noFocusSteal}
              className="w-8 h-8 p-0 border rounded cursor-pointer"
              title="Fill color"
            />
            <TB onClick={applyFillColor} title="Apply fill color">
              Fill
            </TB>
          </div>

          <span className="mx-1 h-6 border-l" />

          <TB
            onClick={() => setAlignment("left")}
            active={selectedAlign === "left"}
            title="Align left"
          >
            Left
          </TB>
          <TB
            onClick={() => setAlignment("center")}
            active={selectedAlign === "center"}
            title="Align center"
          >
            Center
          </TB>
          <TB
            onClick={() => setAlignment("right")}
            active={selectedAlign === "right"}
            title="Align right"
          >
            Right
          </TB>
          <TB
            onClick={() => setAlignment("justify")}
            active={selectedAlign === "justify"}
            title="Justify"
          >
            Justify
          </TB>
          <TB
            onClick={() => setVerticalAlignment("top")}
            active={selectedVAlign === "top"}
          >
            Top
          </TB>
          <TB
            onClick={() => setVerticalAlignment("middle")}
            active={selectedVAlign === "middle"}
          >
            Middle
          </TB>
          <TB
            onClick={() => setVerticalAlignment("bottom")}
            active={selectedVAlign === "bottom"}
          >
            Bottom
          </TB>
          <TB onClick={setWrapText}>Wrap</TB>

          <span className="mx-1 h-6 border-l" />

          <TB
            onClick={() => formatSelectedAs("number")}
            title="Format as number"
          >
            123
          </TB>
          <TB
            onClick={() => formatSelectedAs("currency")}
            title="Format as currency"
          >
            $
          </TB>
          <TB
            onClick={() => formatSelectedAs("percent")}
            title="Format as percent"
          >
            %
          </TB>
          <TB onClick={() => formatSelectedAs("date")} title="Format as date">
            Date
          </TB>

          <span className="mx-1 h-6 border-l" />

          <TB onClick={() => sortSelectedColumn("asc")}>A→Z</TB>
          <TB onClick={() => sortSelectedColumn("desc")}>Z→A</TB>

          <span className="mx-1 h-6 border-l" />

          <TB onClick={mergeSelection}>Merge</TB>
          <TB onClick={unmergeSelection}>Unmerge</TB>

          <span className="mx-1 h-6 border-l" />

          <TB onClick={() => alterBySelection("insert_row_above")}>+Row↑</TB>
          <TB onClick={() => alterBySelection("insert_row_below")}>+Row↓</TB>
          <TB onClick={() => alterBySelection("insert_col_start")}>+Col←</TB>
          <TB onClick={() => alterBySelection("insert_col_end")}>+Col→</TB>
          <TB onClick={() => alterBySelection("remove_row")}>Del Row</TB>
          <TB onClick={() => alterBySelection("remove_col")}>Del Col</TB>

          <span className="mx-1 h-6 border-l" />

          <TB onClick={clearSelectionValues}>Clear Values</TB>
          <TB onClick={clearSelectionFormatting}>Clear Format</TB>

          <span className="mx-1 h-6 border-l" />

          <input
            value={findValue}
            onChange={(e) => setFindValue(e.target.value)}
            onMouseDown={noFocusSteal}
            placeholder="Find"
            className="h-8 px-2 text-sm border rounded w-24"
          />
          <input
            value={replaceValue}
            onChange={(e) => setReplaceValue(e.target.value)}
            onMouseDown={noFocusSteal}
            placeholder="Replace"
            className="h-8 px-2 text-sm border rounded w-24"
          />
          <TB onClick={doFindReplace}>Replace</TB>

          <span className="mx-1 h-6 border-l" />

          <TB
            onClick={() => {
              const hot = hotRef.current?.hotInstance;
              const idx = activeSheetIndexRef.current;
              const r =
                sheetSelectionRef.current[idx] ?? lastSelectionRef.current;
              setFormatAllCells((p) => !p);
              queueMicrotask(() => {
                if (hot && r) restoreHotRange(hot, r);
              });
            }}
            active={formatAllCells}
            title="Apply formatting to ALL cells in sheet"
          >
            All Cells
          </TB>
          <TB onClick={emitWorkbookSnapshot} variant="default">
            Save Workbook
          </TB>
          <TB onClick={exportXlsx}>Export .xlsx</TB>
          <TB onClick={exportCsv}>CSV</TB>
        </div>
      )}

      {!readOnly && (
        <div className="relative z-10 flex items-center gap-2 p-2 border rounded-md bg-white">
          <span className="text-xs text-gray-500 font-medium w-12 shrink-0">
            {selectionLabel}
          </span>
          <span className="text-xs text-gray-400 select-none">fx</span>
          <input
            value={formulaInput}
            onChange={(e) => setFormulaInput(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") {
                e.preventDefault();
                applyFormulaBar();
              }
              if (e.key === "Escape") {
                e.preventDefault();
                const hot = hotRef.current?.hotInstance;
                const idx = activeSheetIndexRef.current;
                const r =
                  sheetSelectionRef.current[idx] ?? lastSelectionRef.current;
                if (hot) {
                  const v = hot.getDataAtCell(r.startRow, r.startCol);
                  setFormulaInput(v == null ? "" : String(v));
                  restoreHotRange(hot, r);
                }
              }
            }}
            className="flex-1 h-8 px-2 text-sm border rounded font-mono"
            placeholder="Type formula/value for active cell (e.g. =SUM(A1:A5))"
          />
          <Button
            type="button"
            size="sm"
            variant="outline"
            onMouseDown={noFocusSteal}
            onClick={applyFormulaBar}
          >
            ✓ Apply
          </Button>
        </div>
      )}

      {/* ── Sheet tabs + sheet actions (separate rows) ── */}
      <div className="relative z-10 space-y-2">
        <div className="flex flex-wrap items-center gap-2 border-b border-slate-200 pb-2">
          {sheetTabs.map((sheet, index) => (
            <Button
              key={`${sheet.name}-${index}`}
              type="button"
              variant={index === activeSheetIndex ? "default" : "outline"}
              size="sm"
              style={
                sheet.tabColor
                  ? { backgroundColor: sheet.tabColor, color: "#111827" }
                  : undefined
              }
              onMouseDown={noFocusSteal}
              onClick={() => handleSheetSwitch(index)}
            >
              {sheet.name}
            </Button>
          ))}
        </div>

        {!readOnly && (
          <div className="flex flex-wrap items-center gap-2">
            {renaming ? (
              <div className="flex items-center gap-1">
                <input
                  value={renameValue}
                  onChange={(e) => setRenameValue(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") {
                      const n = renameValue.trim();
                      if (!n) return;
                      const next = safeSheets.map((s, i) =>
                        i === activeSheetIndex ? { ...s, name: n } : s,
                      );
                      workbookRef.current.sheets = next;
                      setSheetTabs(
                        next.map((s) => ({
                          name: s.name,
                          tabColor: s.tabColor,
                        })),
                      );
                      setRenaming(false);
                    }
                    if (e.key === "Escape") setRenaming(false);
                  }}
                  autoFocus
                  className="h-8 px-2 text-sm border rounded"
                />
                <Button
                  type="button"
                  variant="outline"
                  size="sm"
                  onMouseDown={noFocusSteal}
                  onClick={() => {
                    const n = renameValue.trim();
                    if (!n) return;
                    const next = safeSheets.map((s, i) =>
                      i === activeSheetIndex ? { ...s, name: n } : s,
                    );
                    workbookRef.current.sheets = next;
                    setSheetTabs(
                      next.map((s) => ({ name: s.name, tabColor: s.tabColor })),
                    );
                    setRenaming(false);
                  }}
                >
                  Save
                </Button>
                <Button
                  type="button"
                  variant="outline"
                  size="sm"
                  onMouseDown={noFocusSteal}
                  onClick={() => setRenaming(false)}
                >
                  Cancel
                </Button>
              </div>
            ) : (
              <Button
                type="button"
                variant="outline"
                size="sm"
                className="text-yellow-900 border-yellow-500 bg-yellow-300 hover:bg-yellow-400"
                onMouseDown={noFocusSteal}
                onClick={() => {
                  setRenameValue(activeSheet?.name || "");
                  setRenaming(true);
                }}
              >
                Rename Sheet
              </Button>
            )}

            <Button
              type="button"
              variant="outline"
              size="sm"
              className="text-green-900 border-green-500 bg-green-300 hover:bg-green-400"
              onMouseDown={noFocusSteal}
              onClick={() => {
                const nextSheets = [
                  ...safeSheets,
                  { name: `Sheet${safeSheets.length + 1}`, grid: [[""]] },
                ];
                workbookRef.current.sheets = nextSheets;
                setSheetTabs(
                  nextSheets.map((s) => ({
                    name: s.name,
                    tabColor: s.tabColor,
                  })),
                );
                setInitialGrid(
                  toVisibleGrid(nextSheets[nextSheets.length - 1]),
                );
                setActiveSheetIndex(nextSheets.length - 1);
              }}
            >
              + Add Sheet
            </Button>

            <TB onClick={duplicateActiveSheet}>Duplicate</TB>
            <TB onClick={() => moveSheet("left")}>Move Left</TB>
            <TB onClick={() => moveSheet("right")}>Move Right</TB>

            <input
              type="color"
              className="w-8 h-8 p-0 border rounded cursor-pointer"
              title="Tab color"
              onMouseDown={noFocusSteal}
              onChange={(e) => applySheetColor(e.target.value)}
            />

            <span className="mx-1 h-6 border-l" />

            <input
              value={dropdownSource}
              onChange={(e) => setDropdownSource(e.target.value)}
              onMouseDown={noFocusSteal}
              className="h-8 px-2 text-sm border rounded"
              placeholder="Dropdown: A,B,C"
            />
            <TB onClick={applyDropdownValidation}>Set Dropdown</TB>
            <TB onClick={applyDateCellType}>Set Date Cell</TB>
            <TB
              onClick={toggleFillableSelection}
              title="Mark selected cells as fillable in Preview"
            >
              Mark Fillable
            </TB>

            <span className="text-xs text-gray-500 ml-1">Freeze</span>
            <input
              value={fixedRowsTop}
              type="number"
              min={0}
              onChange={(e) =>
                setFixedRowsTop(Math.max(0, +e.target.value || 0))
              }
              onMouseDown={noFocusSteal}
              className="w-12 h-8 px-2 text-sm border rounded"
              title="Freeze rows"
            />
            <input
              value={fixedColumnsStart}
              type="number"
              min={0}
              onChange={(e) =>
                setFixedColumnsStart(Math.max(0, +e.target.value || 0))
              }
              onMouseDown={noFocusSteal}
              className="w-12 h-8 px-2 text-sm border rounded"
              title="Freeze columns"
            />
          </div>
        )}
      </div>

      {/* ── Grid ── */}
      <div className="relative z-0 overflow-hidden border rounded-md">
        <HotTable
          /* New instance per sheet / workbook shape: Handsontable reuses `metaManager` across
           * `loadData()`, so dropdowns, types, merge flags, etc. from one sheet could otherwise
           * leak onto another at the same coordinates. */
          key={`ht-wb-${activeSheetIndex}-${hotTableMountKey}`}
          ref={hotRef}
          data={initialGrid}
          themeName="ht-theme-main"
          rowHeaders
          colHeaders
          licenseKey="non-commercial-and-evaluation"
          readOnly={false}
          trimWhitespace={false}
          width="100%"
          stretchH={stretchColumnsInPreview ? "all" : "none"}
          height={readOnly ? readOnlyHotHeight ?? 380 : 320}
          formulas={shouldUseFormulaEngine ? FORMULAS_CONFIG : undefined}
          mergeCells={
            renderedMergeCells.length > 0 ? renderedMergeCells : !readOnly
          }
          filters={!readOnly}
          dropdownMenu={!readOnly}
          columnSorting={!readOnly}
          {...(!readOnly
            ? {
                hiddenRows: { indicators: true } as const,
                hiddenColumns: { indicators: true } as const,
              }
            : {})}
          multiColumnSorting={!readOnly}
          manualColumnFreeze={!readOnly}
          autoColumnSize={false}
          autoRowSize={false}
          fillHandle={!readOnly}
          fixedRowsTop={fixedRowsTop}
          fixedColumnsStart={fixedColumnsStart}
          contextMenu={
            readOnly
              ? false
              : {
                  items: {
                    row_above: {},
                    row_below: {},
                    col_left: {},
                    col_right: {},
                    hsep1: "---------",
                    remove_row: {},
                    remove_col: {},
                    hidden_rows_hide: {},
                    hidden_rows_show: {},
                    hidden_columns_hide: {},
                    hidden_columns_show: {},
                    hsep2: "---------",
                    mergeCells: {},
                    hsep3: "---------",
                    alignment: {},
                    freeze_column: {},
                    unfreeze_column: {},
                    hsep4: "---------",
                    copy: {},
                    cut: {},
                    hsep5: "---------",
                    undo: {},
                    redo: {},
                  },
                }
          }
          className="ht-theme-main"
          manualRowResize={!readOnly}
          manualColumnResize={!readOnly}
          colWidths={renderedColWidths as any}
          rowHeights={renderedRowHeights as any}
          wordWrap
          autoWrapRow
          autoWrapCol
          cells={cellsCallback}
          afterGetCellMeta={afterGetCellMeta}
          afterColumnResize={() => flushLayoutToParent()}
          afterRowResize={() => flushLayoutToParent()}
          afterChange={(changes, source) => {
            // Preview (`readOnly`): never call undo/redo setState here — it runs on every
            // keystroke and re-renders the HotTable wrapper, which closes the editor and drops input.
            if (!readOnly && Array.isArray(changes) && changes.length > 0) {
              refreshUndoRedoState();
            }
            if (
              readOnly &&
              changes &&
              source !== "loadData" &&
              source !== "updateData"
            ) {
              const idx = activeSheetIndexRef.current;
              const sheet = workbookRef.current.sheets[idx];
              if (!sheet) return;
              const baseGrid = sheet.grid;
              let newGrid: string[][] = baseGrid;
              for (const [row, col, , newValue] of changes as [
                number,
                number,
                unknown,
                unknown,
              ][]) {
                if (
                  typeof row !== "number" ||
                  typeof col !== "number" ||
                  !baseGrid[row]
                )
                  continue;
                if (newGrid === baseGrid) {
                  newGrid = baseGrid.map((r, ri) =>
                    ri === row ? [...r] : r,
                  ) as string[][];
                } else if (newGrid[row] === baseGrid[row]) {
                  newGrid[row] = [...baseGrid[row]];
                }
                newGrid[row][col] = newValue == null ? "" : String(newValue);
              }
              if (newGrid !== baseGrid) {
                workbookRef.current.sheets[idx] = {
                  ...sheet,
                  grid: newGrid,
                };
                readOnlyPreviewDirtyRef.current = true;
              }
            }
          }}
          afterSelection={(r, c, r2, c2) => {
            if (!Number.isInteger(r) || !Number.isInteger(c) || r < 0 || c < 0)
              return;
            const endRow = Number.isInteger(r2) ? r2 : r;
            const endCol = Number.isInteger(c2) ? c2 : c;
            const range = {
              startRow: Math.min(r, endRow),
              endRow: Math.max(r, endRow),
              startCol: Math.min(c, endCol),
              endCol: Math.max(c, endCol),
            };
            lastSelectionRef.current = range;
            sheetSelectionRef.current[activeSheetIndexRef.current] = range;
            // Preview: never setState here — Handsontable can fire `afterSelection` while the
            // cell editor is open; re-rendering the React wrapper tears down the editor and
            // focus jumps away (typing appears "blocked").
            if (!readOnly) setSelectionLabel(toRangeLabel(range));
          }}
          afterSelectionEnd={(r, c, r2, c2) => {
            const hot = hotRef.current?.hotInstance;
            if (
              !hot ||
              !Number.isInteger(r) ||
              !Number.isInteger(c) ||
              r < 0 ||
              c < 0
            )
              return;
            const endRow = Number.isInteger(r2) ? r2 : r;
            const endCol = Number.isInteger(c2) ? c2 : c;
            const range = {
              startRow: Math.min(r, endRow),
              endRow: Math.max(r, endRow),
              startCol: Math.min(c, endCol),
              endCol: Math.max(c, endCol),
            };
            lastSelectionRef.current = range;
            sheetSelectionRef.current[activeSheetIndexRef.current] = range;

            if (readOnly) {
              // Find the scrollable modal container
              const root = hot.rootElement as HTMLElement | undefined;
              const container =
                root?.closest('[role="dialog"]') ??
                root?.closest(
                  "[data-radix-scroll-area-viewport], .overflow-y-auto, .overflow-auto",
                ) ??
                document.documentElement;

              const savedTop = (container as HTMLElement)?.scrollTop ?? 0;
              const savedLeft = (container as HTMLElement)?.scrollLeft ?? 0;

              // These setState calls cause re-render → scroll jump
              setSelectionLabel(toRangeLabel(range));
              const v = hot.getDataAtCell(range.startRow, range.startCol);
              setFormulaInput(v == null ? "" : String(v));

              // Restore scroll after React commits the re-render
              requestAnimationFrame(() => {
                requestAnimationFrame(() => {
                  if (container && container !== document.documentElement) {
                    (container as HTMLElement).scrollTop = savedTop;
                    (container as HTMLElement).scrollLeft = savedLeft;
                  } else {
                    window.scrollTo(savedLeft, savedTop);
                  }
                });
              });
              return;
            }

            setSelectionLabel(toRangeLabel(range));
            syncToolbarFromCell(hot, r, c);
          }}
          afterMergeCells={() => {
            if (!readOnly) {
              collectCurrentSheetFromHot(true);
              refreshUndoRedoState();
            }
          }}
          afterUnmergeCells={() => {
            if (!readOnly) {
              collectCurrentSheetFromHot(true);
              refreshUndoRedoState();
            }
          }}
          afterCreateRow={() => refreshUndoRedoState()}
          afterCreateCol={() => refreshUndoRedoState()}
          afterRemoveRow={() => refreshUndoRedoState()}
          afterRemoveCol={() => refreshUndoRedoState()}
        />
      </div>

      {isPreviewTruncated && (
        <div className="px-2 py-1 text-xs text-amber-700 border border-amber-200 rounded bg-amber-50">
          Preview mode — showing first {previewRows} rows × {previewCols}{" "}
          columns for stability.
        </div>
      )}
    </div>
  );
});

export default HandsontableWorkbook;
