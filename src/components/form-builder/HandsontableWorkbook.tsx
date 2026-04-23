import React from "react";
import { HotTable } from "@handsontable/react";
import "handsontable/styles/handsontable.css";
import "handsontable/styles/ht-theme-main.css";
import { registerAllModules } from "handsontable/registry";
import { checkboxRenderer, textRenderer } from "handsontable/renderers";
import { Button } from "../ui/button";
import { HyperFormula } from "hyperformula";
import {
  HandsontableWorkbookProps,
  HandsontableWorkbookRef,
  MAX_PREVIEW_COLS,
  MAX_PREVIEW_ROWS,
  SheetData,
} from "./workbook/workbookTypes";
import {
  EMPTY_GRID_PLACEHOLDER,
  SINGLE_CHECKBOX_CLASS,
  buildYesNoOppositeMap,
  cellCoordKey,
  classNameHasFillable,
  cloneEditableGrid,
  dedupeCellMetaByCoordinate,
  deepCloneSheet,
  extractYesNoPairToken,
  hotTableMountSignature,
  mergeFillableMetaFromPrevSheet,
  normalizeSheets,
  toCheckboxChecked,
  toRangeLabel,
  workbookSignature,
} from "./workbook/workbookUtils";
import { useWorkbookSelection } from "./workbook/useWorkbookSelection";
import { useWorkbookToolbarActions } from "./workbook/useWorkbookToolbarActions";
import { useWorkbookStateSync } from "./workbook/useWorkbookStateSync";
import { useWorkbookHotCallbacks } from "./workbook/useWorkbookHotCallbacks";

export type {
  HandsontableWorkbookProps,
  HandsontableWorkbookRef,
  SheetData,
} from "./workbook/workbookTypes";
export { MAX_PREVIEW_COLS, MAX_PREVIEW_ROWS } from "./workbook/workbookTypes";

registerAllModules();
const FORMULA_PREFIX = "=";

/** Handsontable text editor — duck-typed (avoid importing private editor class). */
type HotTextEditorLike = {
  TEXTAREA?: HTMLTextAreaElement;
  TEXTAREA_PARENT?: HTMLDivElement;
  row: number;
  col: number;
  TD?: HTMLTableCellElement | null;
  autoResize?: { unObserve?: () => void };
};

/** Match template styling: only `meta-wrap` marks a multi-line cell in our renderer. */
function cellHasMetaWrap(meta: { className?: string }) {
  return String(meta?.className || "")
    .split(/\s+/)
    .filter(Boolean)
    .includes("meta-wrap");
}

function getMergedRegionFromHot(
  hot: any,
  row: number,
  col: number,
): { row: number; col: number; rowspan: number; colspan: number } | null {
  const merged = hot
    ?.getPlugin?.("mergeCells")
    ?.mergedCellsCollection?.get?.(row, col);
  if (!merged || merged === false) return null;
  return {
    row: merged.row,
    col: merged.col,
    rowspan: merged.rowspan,
    colspan: merged.colspan,
  };
}

function sumColWidthsForMerge(
  hot: any,
  merge: { col: number; colspan: number },
) {
  let sum = 0;
  for (let c = merge.col; c < merge.col + merge.colspan; c++) {
    const w =
      typeof hot?.getColWidth === "function" ? Number(hot.getColWidth(c)) : NaN;
    sum += Number.isFinite(w) && w > 0 ? w : 50;
  }
  return sum;
}

function sumRowHeightsForMerge(
  hot: any,
  merge: { row: number; rowspan: number },
) {
  let sum = 0;
  for (let r = merge.row; r < merge.row + merge.rowspan; r++) {
    const h =
      typeof hot?.getRowHeight === "function"
        ? Number(hot.getRowHeight(r))
        : NaN;
    sum += Number.isFinite(h) && h > 0 ? h : 23;
  }
  return sum;
}

/**
 * Sizes the default TEXTAREA editor to the rendered TD (merged colspan/rowspan
 * included). Disables HOT's autoResize observer so it cannot shrink the editor
 * to a single-column text measure.
 */
function syncHandsontableTextEditorToCell(hot: any) {
  if (!hot) return;
  const opened =
    typeof hot.isEditorOpened === "function" && hot.isEditorOpened();
  if (!opened) return;
  const editor = hot.getActiveEditor?.() as HotTextEditorLike | undefined;
  const ta = editor?.TEXTAREA;
  const holder = editor?.TEXTAREA_PARENT;
  if (!editor || !ta || !holder) return;

  editor.autoResize?.unObserve?.();

  const row = editor.row;
  const col = editor.col;
  const td =
    (editor.TD as HTMLTableCellElement | null | undefined) ??
    (hot.getCell(row, col, true) as HTMLTableCellElement | null);
  if (!td) return;

  const merge = getMergedRegionFromHot(hot, row, col);
  const rect = td.getBoundingClientRect();
  let cellW = Math.max(1, Math.round(rect.width));
  let cellH = Math.max(1, Math.round(rect.height));
  if (merge) {
    const sumW = Math.round(sumColWidthsForMerge(hot, merge));
    const sumH = Math.round(sumRowHeightsForMerge(hot, merge));
    if (sumW > 0) cellW = Math.max(cellW, sumW);
    if (sumH > 0) cellH = Math.max(cellH, sumH);
  }

  const meta = hot.getCellMeta(row, col) as { className?: string };
  const wraps = cellHasMetaWrap(meta);

  const hs = holder.style;
  const ts = ta.style;

  hs.boxSizing = "border-box";
  hs.width = `${cellW}px`;
  hs.minWidth = `${cellW}px`;
  hs.maxWidth = `${cellW}px`;
  hs.overflow = "hidden";

  ts.boxSizing = "border-box";
  ts.width = "100%";
  ts.minWidth = "100%";
  ts.maxWidth = "100%";
  ts.margin = "0";
  ts.resize = "none";

  const tdStyle = hot.rootWindow.getComputedStyle(td);
  ts.fontSize = tdStyle.fontSize;
  ts.fontFamily = tdStyle.fontFamily;
  ts.lineHeight = tdStyle.lineHeight;
  ts.padding = `${tdStyle.paddingTop} ${tdStyle.paddingRight} ${tdStyle.paddingBottom} ${tdStyle.paddingLeft}`;

  if (wraps) {
    ts.whiteSpace = "pre-wrap";
    ts.wordBreak = "break-word";
    ts.overflowX = "hidden";
    ts.overflowY = "auto";
    const grow = () => {
      ts.height = "auto";
      const innerMin = Math.max(18, cellH - 2);
      const nextH = Math.max(innerMin, ta.scrollHeight);
      ts.minHeight = `${innerMin}px`;
      ts.height = `${nextH}px`;
      hs.minHeight = `${cellH}px`;
      hs.height = `${Math.max(cellH, nextH + 2)}px`;
      hs.maxHeight = "none";
    };
    grow();
    (ta as any).__htGrowWrap = grow;
  } else {
    delete (ta as any).__htGrowWrap;
    ts.whiteSpace = "pre";
    ts.overflowX = "auto";
    ts.overflowY = "hidden";
    ts.height = `${cellH}px`;
    ts.minHeight = `${cellH}px`;
    ts.maxHeight = `${cellH}px`;
    hs.height = `${cellH}px`;
    hs.minHeight = `${cellH}px`;
    hs.maxHeight = `${cellH}px`;
  }
}

const toFormulaDisplayValue = (value: unknown) => {
  if (value == null) return "";
  if (value === "") return "";
  if (typeof value === "object") {
    const err = (value as any)?.value;
    if (typeof err === "string" && err.startsWith("#")) return err;
  }
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "number")
    return Number.isFinite(value) ? String(value) : "#NUM!";
  return String(value);
};

const normalizeVisibleGrid = (rawGrid: unknown): string[][] => {
  const base =
    Array.isArray(rawGrid) && rawGrid.length > 0 ? (rawGrid as unknown[]) : [[""]];
  const maxCols = Math.max(
    1,
    ...base.map((row) => (Array.isArray(row) ? row.length : 0)),
  );
  return base.map((row) => {
    const normalizedRow = Array.isArray(row) ? [...row] : [];
    while (normalizedRow.length < maxCols) normalizedRow.push("");
    return normalizedRow.map((cell) => (cell == null ? "" : String(cell)));
  });
};

// Prevent toolbar buttons from stealing focus from the grid
const noFocusSteal = (e: React.MouseEvent) => e.preventDefault();

type ToolbarButtonProps = {
  onClick: () => void;
  children: React.ReactNode;
  title?: string;
  variant?: "outline" | "default";
  disabled?: boolean;
  active?: boolean;
  className?: string;
};

const TB = ({
  onClick,
  children,
  title,
  variant = "outline",
  disabled = false,
  active = false,
  className = "",
}: ToolbarButtonProps) => (
  <Button
    type="button"
    size="sm"
    variant={active ? "default" : variant}
    disabled={disabled}
    title={title}
    className={className}
    onMouseDown={noFocusSteal}
    onClick={onClick}
  >
    {children}
  </Button>
);

// ─── component ────────────────────────────────────────────────────────────────

const HandsontableWorkbook = React.forwardRef<
  HandsontableWorkbookRef,
  HandsontableWorkbookProps
>(function HandsontableWorkbook(
  {
    data,
    onChange,
    readOnly = false,
    strictViewOnly = false,
    allowReadOnlyWorkbookActions = false,
    readOnlyHotHeight,
    lightweightPerformance = false,
  },
  ref,
) {
  type CellMetaEntry = NonNullable<SheetData["cellMeta"]>[number];
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
    return normalizeVisibleGrid(base);
  });

  const [renaming, setRenaming] = React.useState(false);
  const [renameValue, setRenameValue] = React.useState("");
  const [formulaInput, setFormulaInput] = React.useState("");
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

  const {
    lastSelectionRef,
    sheetSelectionRef,
    getToolbarActionRange,
    restoreHotRange,
  } = useWorkbookSelection({
    workbookRef,
    activeSheetIndexRef,
  });

  const hotRef = React.useRef<any>(null);
  const hotViewportRef = React.useRef<HTMLDivElement | null>(null);
  /** Removes `input` / `blur` listeners used to grow the text editor while typing. */
  const editorTextLayoutCleanupRef = React.useRef<(() => void) | null>(null);
  const [hotViewportWidth, setHotViewportWidth] = React.useState(0);

  const textColorApplyTimerRef = React.useRef<ReturnType<
    typeof setTimeout
  > | null>(null);
  const fillColorApplyTimerRef = React.useRef<ReturnType<
    typeof setTimeout
  > | null>(null);
  const undoRedoRefreshTimerRef = React.useRef<ReturnType<
    typeof setTimeout
  > | null>(null);
  const formulaCellSetRef = React.useRef<Set<string>>(new Set());
  const hfRef = React.useRef<HyperFormula | null>(null);
  const yesNoOppositeCellMapRef = React.useRef<
    Map<string, { row: number; col: number }>
  >(new Map());
  const cellMetaRef = React.useRef<
    Record<string, Record<number, Record<number, CellMetaEntry>>>
  >({});
  const suppressNextHotReloadRef = React.useRef(false);
  const pendingIncomingReloadRef = React.useRef(false);
  const pendingIncomingReloadSheetIndexRef = React.useRef<number | null>(null);
  const pendingIncomingReloadWorkbookKeyRef = React.useRef<string | null>(null);
  const lastLoadedSheetIndexRef = React.useRef<number>(-1);
  const lastLoadedWorkbookKeyRef = React.useRef<string>("__NONE__");
  const isEditingRef = React.useRef(false);
  const pendingReadOnlyEmitRef = React.useRef(false);
  const readOnlyEmitDebounceTimerRef = React.useRef<ReturnType<
    typeof setTimeout
  > | null>(null);
  const previewEditingSettleTimerRef = React.useRef<ReturnType<
    typeof setTimeout
  > | null>(null);
  const cellsCacheRef = React.useRef<Map<string, any>>(new Map());
  const mergeCacheFrameRef = React.useRef<{
    frameId: number;
    mergedSet: Set<string>;
  }>({ frameId: -1, mergedSet: new Set() });
  const originalSheetColCountRef = React.useRef<Map<number, number>>(new Map());
  const originalSheetRowCountRef = React.useRef<Map<number, number>>(new Map());
  const columnStructureDirtyRef = React.useRef<Map<number, boolean>>(new Map());
  const rowStructureDirtyRef = React.useRef<Map<number, boolean>>(new Map());
  const preserveScrollOnNextLoadRef = React.useRef(true);
  const pendingScrollRestoreRafRef = React.useRef<number | null>(null);
  const pendingScrollRestoreNestedRafRef = React.useRef<number | null>(null);
  const pendingEditorLayoutRafRef = React.useRef<number | null>(null);
  const pendingEditorLayoutNestedRafRef = React.useRef<number | null>(null);
  const [menuContainer, setMenuContainer] = React.useState<HTMLElement | null>(
    null,
  );
  const disableEditorCompletely = readOnly && strictViewOnly;
  const canUseReadOnlyWorkbookActions =
    readOnly && !strictViewOnly && allowReadOnlyWorkbookActions;
  const showFormulaBar = !readOnly || canUseReadOnlyWorkbookActions;
  const showSheetActions = !readOnly || canUseReadOnlyWorkbookActions;

  const normalizedIncomingSheets = React.useMemo(
    () => normalizeSheets(data),
    [data],
  );
  const incomingWorkbookKey = React.useMemo(
    () => workbookSignature(normalizedIncomingSheets),
    [normalizedIncomingSheets],
  );

  const safeSheets = workbookRef.current.sheets;
  const activeSheet =
    safeSheets[Math.min(activeSheetIndex, safeSheets.length - 1)] ||
    safeSheets[0];
  const activeSheetName = activeSheet?.name || `Sheet${activeSheetIndex + 1}`;

  const getCellMeta = React.useCallback(
    (sheetName: string, row: number, col: number): Partial<CellMetaEntry> => {
      return cellMetaRef.current[sheetName]?.[row]?.[col] ?? {};
    },
    [],
  );

  const setCellMeta = React.useCallback(
    (
      sheetName: string,
      row: number,
      col: number,
      props: Partial<CellMetaEntry>,
    ) => {
      if (!cellMetaRef.current[sheetName]) cellMetaRef.current[sheetName] = {};
      if (!cellMetaRef.current[sheetName][row])
        cellMetaRef.current[sheetName][row] = {};
      const prev = cellMetaRef.current[sheetName][row][col];
      cellMetaRef.current[sheetName][row][col] = {
        ...(prev || { row, col }),
        ...props,
        row,
        col,
      } as CellMetaEntry;
    },
    [],
  );

  const setSheetCellMetaFromList = React.useCallback(
    (sheetName: string, list: NonNullable<SheetData["cellMeta"]>) => {
      const byRow: Record<number, Record<number, CellMetaEntry>> = {};
      for (const meta of list || []) {
        if (!meta || !Number.isFinite(+meta.row) || !Number.isFinite(+meta.col))
          continue;
        const row = +meta.row;
        const col = +meta.col;
        if (!byRow[row]) byRow[row] = {};
        byRow[row][col] = { ...meta, row, col };
      }
      cellMetaRef.current[sheetName] = byRow;
    },
    [],
  );

  const getSheetCellMetaList = React.useCallback(
    (sheetName: string): NonNullable<SheetData["cellMeta"]> => {
      const byRow = cellMetaRef.current[sheetName] || {};
      const out: NonNullable<SheetData["cellMeta"]> = [];
      for (const rowKey of Object.keys(byRow)) {
        const rowMeta = byRow[+rowKey] || {};
        for (const colKey of Object.keys(rowMeta)) {
          const meta = rowMeta[+colKey];
          if (meta) out.push({ ...meta, row: +meta.row, col: +meta.col });
        }
      }
      return dedupeCellMetaByCoordinate(out);
    },
    [],
  );
  const safeGrid =
    Array.isArray(activeSheet?.grid) && activeSheet.grid.length > 0
      ? activeSheet.grid
      : EMPTY_GRID_PLACEHOLDER;
  const previewRows = safeGrid.length;
  const previewCols = safeGrid[0]?.length || 0;
  const renderedGrid = safeGrid;

  const isPreviewTruncated = false;

  const lastMergeSigRef = React.useRef<string>("");
  const stableMergesRef = React.useRef<any[]>([]);

  const renderedMergeCells = React.useMemo(() => {
    const filtered = (activeSheet.mergeCells || []).filter(
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
    );
    const sig = JSON.stringify(filtered);
    if (sig === lastMergeSigRef.current) {
      return stableMergesRef.current;
    }
    lastMergeSigRef.current = sig;
    stableMergesRef.current = filtered;
    return filtered;
  }, [activeSheet.mergeCells, previewRows, previewCols]);

  const renderedColWidths = React.useMemo(() => {
    if (!readOnly) return activeSheet.colWidthsPx;
    return (activeSheet.colWidthsPx || []).slice(0, previewCols);
  }, [readOnly, activeSheet.colWidthsPx, previewCols]);

  const renderedRowHeights = React.useMemo(() => {
    if (!readOnly) return activeSheet.rowHeightsPx;
    return (activeSheet.rowHeightsPx || []).slice(0, previewRows);
  }, [readOnly, activeSheet.rowHeightsPx, previewRows]);

  const stretchColumnsInPreview =
    readOnly &&
    (!Array.isArray(activeSheet.colWidthsPx) ||
      activeSheet.colWidthsPx.length === 0);

  React.useEffect(() => {
    const el = hotViewportRef.current;
    if (!el || typeof ResizeObserver === "undefined") return;
    const ro = new ResizeObserver((entries) => {
      const width = entries?.[0]?.contentRect?.width ?? 0;
      setHotViewportWidth(width);
    });
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  React.useEffect(() => {
    const hot = hotRef.current?.hotInstance;
    if (!hot) return;
    if (Array.isArray(renderedColWidths) && renderedColWidths.length > 0) {
      hot.updateSettings({ colWidths: renderedColWidths }, false);
    }
  }, [renderedColWidths]);

  React.useEffect(() => {
    const hot = hotRef.current?.hotInstance;
    if (!hot) return;
    if (Array.isArray(renderedRowHeights) && renderedRowHeights.length > 0) {
      hot.updateSettings({ rowHeights: renderedRowHeights }, false);
    }
  }, [renderedRowHeights]);

  React.useLayoutEffect(() => {
    const dialog = hotViewportRef.current?.closest(
      '[role="dialog"]',
    ) as HTMLElement | null;
    if (dialog) setMenuContainer(dialog);
  }, []);

  const hotTableZoom = React.useMemo(
    () => 1,
    [readOnly, hotViewportWidth, renderedGrid, renderedColWidths],
  );

  const hotTableScaleStyle = React.useMemo<React.CSSProperties>(() => {
    if (hotTableZoom >= 0.999) return {};
    return {
      transform: `scaleX(${hotTableZoom})`,
      transformOrigin: "top left",
      width: `${100 / hotTableZoom}%`,
    };
  }, [hotTableZoom]);

  const currentCellCount = React.useMemo(
    () =>
      renderedGrid.reduce(
        (t, row) => t + (Array.isArray(row) ? row.length : 0),
        0,
      ),
    [renderedGrid],
  );
  const shouldUseFormulaEngine = currentCellCount <= 20000;

  const [isHotLoading, setIsHotLoading] = React.useState(false);

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
    const scopedMeta = getSheetCellMetaList(activeSheetName);
    const sourceMeta =
      scopedMeta.length > 0
        ? scopedMeta
        : dedupeCellMetaByCoordinate(activeSheet?.cellMeta || []);
    for (const meta of sourceMeta)
      map.set(cellCoordKey(+meta.row, +meta.col), meta);
    return map;
  }, [
    activeSheetName,
    activeSheet?.cellMeta,
    getSheetCellMetaList,
    incomingWorkbookKey,
  ]);

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

  const refreshUndoRedoState = React.useCallback(() => {
    const hot = hotRef.current?.hotInstance;
    const ur = hot?.getPlugin?.("undoRedo");
    setCanUndo(Boolean(ur?.isUndoAvailable?.()));
    setCanRedo(Boolean(ur?.isRedoAvailable?.()));
  }, []);

  const scheduleUndoRedoRefresh = React.useCallback(() => {
    if (undoRedoRefreshTimerRef.current) {
      clearTimeout(undoRedoRefreshTimerRef.current);
    }
    undoRedoRefreshTimerRef.current = setTimeout(() => {
      undoRedoRefreshTimerRef.current = null;
      refreshUndoRedoState();
    }, 300);
  }, [refreshUndoRedoState]);

  const getFormulaMeta = React.useCallback(
    (sheet: SheetData | undefined, row: number, col: number) =>
      (sheet?.cellMeta || []).find((m) => m.row === row && m.col === col),
    [],
  );

  const initializeHyperFormula = React.useCallback(() => {
    hfRef.current?.destroy();
    const sheets = workbookRef.current.sheets;
    const byName: Record<string, (string | number | boolean)[][]> = {};
    for (let sIdx = 0; sIdx < sheets.length; sIdx++) {
      const sheet = sheets[sIdx];
      const metaByKey = new Map<string, CellMetaEntry>();
      for (const m of sheet.cellMeta || []) {
        const normalizedMeta: CellMetaEntry = { ...m };
        if (
          typeof normalizedMeta.formula === "string" &&
          normalizedMeta.formula.startsWith(FORMULA_PREFIX) &&
          normalizedMeta.formulaCachedValue === ""
        ) {
          delete (normalizedMeta as any).formulaCachedValue;
        }
        metaByKey.set(cellCoordKey(m.row, m.col), normalizedMeta);
      }
      const rows = Math.max(sheet.grid?.length || 1, 1);
      const cols = Math.max(
        1,
        ...(sheet.grid || []).map((r) => (Array.isArray(r) ? r.length : 0)),
      );
      byName[sheet.name || `Sheet${sIdx + 1}`] = Array.from(
        { length: rows },
        (_, row) =>
          Array.from({ length: cols }, (_, col) => {
            const formula = metaByKey.get(cellCoordKey(row, col))?.formula;
            if (
              typeof formula === "string" &&
              formula.startsWith(FORMULA_PREFIX)
            ) {
              return formula;
            }
            const raw = sheet.grid?.[row]?.[col] ?? "";
            if (
              typeof raw === "string" &&
              raw.startsWith(FORMULA_PREFIX) &&
              raw.length > 1
            ) {
              const key = cellCoordKey(row, col);
              const current =
                metaByKey.get(key) || ({ row, col } as CellMetaEntry);
              current.formula = raw;
              metaByKey.set(key, current);
              return raw;
            }
            const num = Number(raw);
            if (String(raw).trim() !== "" && Number.isFinite(num)) return num;
            return String(raw);
          }),
      );
      sheet.cellMeta = dedupeCellMetaByCoordinate([...metaByKey.values()]);
    }
    hfRef.current = HyperFormula.buildFromSheets(byName, {
      licenseKey: "gpl-v3",
    });
  }, []);

  const refreshFormulaDisplays = React.useCallback(() => {
    const hf = hfRef.current;
    if (!hf) return new Map<number, Array<[number, number, string]>>();
    const sheets = workbookRef.current.sheets;
    const updatesBySheet = new Map<number, Array<[number, number, string]>>();
    for (let sIdx = 0; sIdx < sheets.length; sIdx++) {
      const sheet = sheets[sIdx];
      const sheetId = hf.getSheetId(sheet.name || `Sheet${sIdx + 1}`);
      if (sheetId == null) continue;
      const updates: Array<[number, number, string]> = [];
      for (const meta of sheet.cellMeta || []) {
        if (typeof meta.formula !== "string" || !meta.formula.startsWith("="))
          continue;
        const value = hf.getCellValue({
          sheet: sheetId,
          row: meta.row,
          col: meta.col,
        });
        const display = toFormulaDisplayValue(value);
        meta.formulaCachedValue = display;
        updates.push([meta.row, meta.col, display]);
      }
      if (updates.length > 0) updatesBySheet.set(sIdx, updates);
    }
    return updatesBySheet;
  }, []);

  const syncToolbarFromCell = React.useCallback(
    (hot: any, row: number, col: number) => {
      const v = hot.getDataAtCell(row, col);
      const sheet = workbookRef.current.sheets[activeSheetIndexRef.current];
      const formula = getFormulaMeta(sheet, row, col)?.formula;
      setFormulaInput(formula ?? (v == null ? "" : String(v)));

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
    [getFormulaMeta],
  );

  // ─── sheet sync ────────────────────────────────────────────────────────────

  const collectCurrentSheetFromHot = React.useCallback(
    (includeMeta: boolean, sheetIndex?: number) => {
      const hot = hotRef.current?.hotInstance;
      if (!hot) return;

      const idx = sheetIndex ?? activeSheetIndexRef.current;

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
      const originalRowCount = originalSheetRowCountRef.current.get(idx) || 1;
      const rowStructureChanged = rowStructureDirtyRef.current.get(idx) === true;
      if (
        !rowStructureChanged &&
        Number.isFinite(originalRowCount) &&
        originalRowCount > 0
      ) {
        while (nextGrid.length < originalRowCount) nextGrid.push([]);
      }
      if (!readOnly) {
        const originalColCount = originalSheetColCountRef.current.get(idx) || 1;
        const structureChanged =
          columnStructureDirtyRef.current.get(idx) === true;
        if (
          !structureChanged &&
          Number.isFinite(originalColCount) &&
          originalColCount > 0
        ) {
          for (let r = 0; r < nextGrid.length; r++) {
            const row = nextGrid[r];
            if (!Array.isArray(row)) continue;
            if (row.length >= originalColCount) continue;
            nextGrid[r] = [
              ...row,
              ...Array.from(
                { length: originalColCount - row.length },
                () => "",
              ),
            ];
          }
        }
      }

      const mergeCells =
        hot
          ?.getPlugin?.("mergeCells")
          ?.mergedCellsCollection?.mergedCells?.map((cell: any) => ({
            row: cell.row,
            col: cell.col,
            rowspan: cell.rowspan,
            colspan: cell.colspan,
          })) || [];

      const targetSheet =
        workbookRef.current.sheets[idx] ||
        ({ name: `Sheet${idx + 1}` } as SheetData);
      const targetSheetName = targetSheet.name || `Sheet${idx + 1}`;
      let cellMeta: NonNullable<SheetData["cellMeta"]>;
      if (includeMeta) {
        cellMeta = getSheetCellMetaList(targetSheetName);
        if ((!cellMeta || cellMeta.length === 0) && Array.isArray(targetSheet.cellMeta)) {
          cellMeta = dedupeCellMetaByCoordinate(targetSheet.cellMeta);
        }
      } else {
        cellMeta = targetSheet.cellMeta || [];
      }
      const hf = hfRef.current;
      const sheetId = hf?.getSheetId(targetSheetName);
      if (hf && sheetId != null) {
        for (const meta of cellMeta || []) {
          const formula = (meta as any)?.formula;
          if (
            typeof formula !== "string" ||
            !formula.startsWith(FORMULA_PREFIX)
          ) {
            continue;
          }
          while (nextGrid.length <= meta.row) nextGrid.push([]);
          if (!Array.isArray(nextGrid[meta.row])) nextGrid[meta.row] = [];
          // Prefer HOT's raw source value when it is already a formula string —
          // this means the user typed a new formula that hasn't been committed
          // to cellMeta yet (e.g. "Save Changes" clicked before afterChange's
          // handleCellChanges ran). Falling back to cellMeta.formula only when
          // HOT stored an evaluated result (a non-formula string).
          const hotRaw = nextGrid[meta.row][meta.col] ?? "";
          if (
            typeof hotRaw !== "string" ||
            !hotRaw.startsWith(FORMULA_PREFIX)
          ) {
            nextGrid[meta.row][meta.col] = formula;
          }
        }
      }

      const current = targetSheet || {
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

      workbookRef.current.sheets[idx] = {
        ...current,
        grid: nextGrid,
        mergeCells,
        cellMeta,
        colWidthsPx,
        rowHeightsPx,
      };
      setSheetCellMetaFromList(targetSheetName, cellMeta);
    },
    [readOnly, getSheetCellMetaList, setSheetCellMetaFromList],
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

  const { emitWorkbookToParent } = useWorkbookStateSync({
    workbookRef,
    lastIncomingSignatureRef,
    suppressNextHotReloadRef,
    onChange,
  });

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
    (sheet?: SheetData) => normalizeVisibleGrid(sheet?.grid),
    [],
  );

  const normalizeLegacyCheckboxValues = React.useCallback(
    (sheet?: SheetData) => {
      if (!sheet?.grid?.length || !sheet?.cellMeta?.length) return;
      const checkboxCoords = new Set<string>();
      for (const meta of sheet.cellMeta) {
        const cls = String(meta?.className || "");
        const isCheckboxMeta =
          String(meta?.type || "") === "checkbox" ||
          Boolean(extractYesNoPairToken(cls)) ||
          cls.split(/\s+/).includes(SINGLE_CHECKBOX_CLASS);
        if (isCheckboxMeta)
          checkboxCoords.add(cellCoordKey(meta.row, meta.col));
      }
      if (checkboxCoords.size === 0) return;

      let gridChanged = false;
      const nextGrid = sheet.grid.map((row, r) => {
        if (!Array.isArray(row)) return row;
        let rowChanged = false;
        const nextRow = [...row];
        for (let c = 0; c < row.length; c++) {
          if (!checkboxCoords.has(cellCoordKey(r, c))) continue;
          const raw = row[c];
          const normalized = toCheckboxChecked(raw) ? "true" : "";
          if (String(raw ?? "") !== normalized) {
            nextRow[c] = normalized;
            rowChanged = true;
            gridChanged = true;
          }
        }
        return rowChanged ? nextRow : row;
      });

      if (gridChanged) {
        sheet.grid = nextGrid;
      }
      if (checkboxCoords.size > 0) {
        let fallbackChanged = false;
        const fallbackGrid = sheet.grid.map((row, r) => {
          if (!Array.isArray(row)) return row;
          let rowChanged = false;
          const nextRow = [...row];
          for (let c = 0; c < row.length; c++) {
            if (!checkboxCoords.has(cellCoordKey(r, c))) continue;
            const raw = row[c];
            const normalizedRaw = String(raw ?? "")
              .trim()
              .toLowerCase();
            if (normalizedRaw === "false") {
              nextRow[c] = "";
              rowChanged = true;
              fallbackChanged = true;
            }
          }
          return rowChanged ? nextRow : row;
        });
        if (fallbackChanged) {
          sheet.grid = fallbackGrid;
        }
      }
    },
    [],
  );

  const loadSheetIntoHot = React.useCallback(
    (targetIndex: number) => {
      const hot = hotRef.current?.hotInstance;
      if (!hot) return;
      const sheet = workbookRef.current.sheets[targetIndex];
      if (!sheet) return;
      setIsHotLoading(true);
      lastLoadedSheetIndexRef.current = targetIndex;
      lastLoadedWorkbookKeyRef.current = incomingWorkbookKey;
      pendingIncomingReloadRef.current = false;
      pendingIncomingReloadSheetIndexRef.current = null;
      pendingIncomingReloadWorkbookKeyRef.current = null;
      normalizeLegacyCheckboxValues(sheet);
      cellsCacheRef.current.clear();
      mergeCacheFrameRef.current = { frameId: -1, mergedSet: new Set() };
      // Clear undo/redo history so per-sheet actions don't bleed across tabs.
      // Previously a full HOT remount (via key change) reset history for free;
      // without the remount we must do it explicitly.
      const urPlugin = (hot as any)?.getPlugin?.("undoRedo") as any;
      if (urPlugin) {
        if (typeof urPlugin.clear === "function") {
          urPlugin.clear();
        } else {
          if (Array.isArray(urPlugin.doneActions)) urPlugin.doneActions = [];
          if (Array.isArray(urPlugin.undoneActions)) urPlugin.undoneActions = [];
        }
      }

      const masterHolder = hot.rootElement?.querySelector?.(
        ".ht_master .wtHolder, .wtHolder",
      ) as HTMLElement | null;
      const savedScrollTop = masterHolder?.scrollTop ?? 0;
      const savedScrollLeft = masterHolder?.scrollLeft ?? 0;

      const visibleGrid = toVisibleGrid(sheet);
      const hf = hfRef.current;
      if (hf) {
        const sheetId = hf.getSheetId(sheet.name || `Sheet${targetIndex + 1}`);
        if (sheetId != null) {
          for (const meta of sheet.cellMeta || []) {
            if (
              typeof meta?.formula !== "string" ||
              !meta.formula.startsWith(FORMULA_PREFIX)
            ) {
              continue;
            }
            const value = hf.getCellValue({
              sheet: sheetId,
              row: meta.row,
              col: meta.col,
            });
            const display = toFormulaDisplayValue(value);
            while (visibleGrid.length <= meta.row) visibleGrid.push([]);
            if (!Array.isArray(visibleGrid[meta.row])) visibleGrid[meta.row] = [];
            visibleGrid[meta.row][meta.col] = display ?? "";
          }
        }
      }
      const sourceColCount = Math.max(
        1,
        ...((sheet?.grid || []).map((row) =>
          Array.isArray(row) ? row.length : 0,
        ) || [1]),
      );
      const sourceRowCount = Math.max(1, sheet?.grid?.length || 1);
      originalSheetColCountRef.current.set(targetIndex, sourceColCount);
      originalSheetRowCountRef.current.set(targetIndex, sourceRowCount);
      columnStructureDirtyRef.current.set(targetIndex, false);
      rowStructureDirtyRef.current.set(targetIndex, false);
      const formulaSet = new Set<string>();
      for (const m of sheet.cellMeta || []) {
        if (
          typeof (m as any).formula === "string" &&
          (m as any).formula.startsWith("=")
        ) {
          formulaSet.add(cellCoordKey(m.row, m.col));
        }
      }
      formulaCellSetRef.current = formulaSet;
      yesNoOppositeCellMapRef.current = buildYesNoOppositeMap(sheet.cellMeta);
      setInitialGrid(visibleGrid);
      hot.loadData(visibleGrid);
      const colWidths = sheet.colWidthsPx;
      if (Array.isArray(colWidths) && colWidths.length > 0) {
        hot.updateSettings({ colWidths }, false);
      }
      const rowHeights = sheet.rowHeightsPx;
      if (Array.isArray(rowHeights) && rowHeights.length > 0) {
        hot.updateSettings({ rowHeights }, false);
      }
      const scopedMeta = getSheetCellMetaList(
        sheet.name || `Sheet${targetIndex + 1}`,
      );
      for (const meta of scopedMeta) {
        if (meta.className)
          hot.setCellMeta(meta.row, meta.col, "className", meta.className);
        if (meta.type) hot.setCellMeta(meta.row, meta.col, "type", meta.type);
        if (meta.checkedTemplate !== undefined)
          hot.setCellMeta(
            meta.row,
            meta.col,
            "checkedTemplate",
            meta.checkedTemplate,
          );
        if (meta.uncheckedTemplate !== undefined)
          hot.setCellMeta(
            meta.row,
            meta.col,
            "uncheckedTemplate",
            meta.uncheckedTemplate,
          );
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
      if (typeof hot.refreshDimensions === "function") {
        hot.refreshDimensions();
      }

      const shouldRestoreScroll =
        preserveScrollOnNextLoadRef.current &&
        targetIndex === activeSheetIndexRef.current &&
        (savedScrollTop > 0 || savedScrollLeft > 0);
      if (shouldRestoreScroll) {
        if (pendingScrollRestoreRafRef.current != null) {
          cancelAnimationFrame(pendingScrollRestoreRafRef.current);
          pendingScrollRestoreRafRef.current = null;
        }
        if (pendingScrollRestoreNestedRafRef.current != null) {
          cancelAnimationFrame(pendingScrollRestoreNestedRafRef.current);
          pendingScrollRestoreNestedRafRef.current = null;
        }
        pendingScrollRestoreRafRef.current = requestAnimationFrame(() => {
          pendingScrollRestoreRafRef.current = null;
          pendingScrollRestoreNestedRafRef.current = requestAnimationFrame(
            () => {
              pendingScrollRestoreNestedRafRef.current = null;
              const h = hotRef.current?.hotInstance;
              const holder = h?.rootElement?.querySelector?.(
                ".ht_master .wtHolder, .wtHolder",
              ) as HTMLElement | null;
              if (holder) {
                holder.scrollTop = savedScrollTop;
                holder.scrollLeft = savedScrollLeft;
              }
            },
          );
        });
      }
      preserveScrollOnNextLoadRef.current = true;
      setIsHotLoading(false);
    },
    [
      incomingWorkbookKey,
      readOnly,
      toVisibleGrid,
      normalizeLegacyCheckboxValues,
      getSheetCellMetaList,
    ],
  );

  const handleSheetSwitch = (targetIndex: number) => {
    if (targetIndex === activeSheetIndex) return;
    preserveScrollOnNextLoadRef.current = false;
    if (!readOnly) {
      // Use ref (not state) to avoid stale closure — state may lag behind
      // the actual loaded sheet index on rapid tab switches.
      collectCurrentSheetFromHot(true, activeSheetIndexRef.current);
      emitWorkbookToParent();
    }
    const hot = hotRef.current?.hotInstance;
    if (hot) getToolbarActionRange(hot);

    // Build the grid with formula results evaluated so the newly mounted
    // HotTable instance (key changes on every tab switch) gets correct
    // display values immediately, before loadSheetIntoHot runs.
    const _switchSheet = workbookRef.current.sheets[targetIndex];
    const _evaluatedGrid = toVisibleGrid(_switchSheet);
    const _hf = hfRef.current;
    if (_hf && _switchSheet) {
      const _sheetId = _hf.getSheetId(
        _switchSheet.name || `Sheet${targetIndex + 1}`
      );
      if (_sheetId != null) {
        for (const _meta of _switchSheet.cellMeta || []) {
          if (
            typeof _meta?.formula !== "string" ||
            !_meta.formula.startsWith("=")
          )
            continue;
          while (_evaluatedGrid.length <= _meta.row) _evaluatedGrid.push([]);
          if (!Array.isArray(_evaluatedGrid[_meta.row]))
            _evaluatedGrid[_meta.row] = [];
          const _value = _hf.getCellValue({
            sheet: _sheetId,
            row: _meta.row,
            col: _meta.col,
          });
          _evaluatedGrid[_meta.row][_meta.col] = toFormulaDisplayValue(_value);
        }
      }
    }
    setInitialGrid(_evaluatedGrid);

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
            setCellMeta(activeSheetName, r, c, {
              className: next.join(" ").trim() || undefined,
            });
          }
        }
      };

      if (typeof hot.batch === "function") hot.batch(apply);
      else apply();
      collectCurrentSheetFromHot(true);
      restoreHotRange(hot, range);
    },
    [
      activeSheetName,
      getToolbarActionRange,
      readOnly,
      collectCurrentSheetFromHot,
      setCellMeta,
    ],
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

  const flushReadOnlyEmitDebounce = React.useCallback(() => {
    if (readOnlyEmitDebounceTimerRef.current) {
      clearTimeout(readOnlyEmitDebounceTimerRef.current);
      readOnlyEmitDebounceTimerRef.current = null;
    }
  }, []);

  const scheduleReadOnlyEmit = React.useCallback(() => {
    flushReadOnlyEmitDebounce();
    readOnlyEmitDebounceTimerRef.current = setTimeout(() => {
      const hot = hotRef.current?.hotInstance;
      const editorOpen =
        typeof hot?.isEditorOpened === "function" && hot.isEditorOpened();
      if (readOnly && (editorOpen || isEditingRef.current)) {
        scheduleReadOnlyEmit();
        return;
      }
      readOnlyEmitDebounceTimerRef.current = null;
      emitWorkbookToParent();
    }, 120);
  }, [
    emitWorkbookToParent,
    flushReadOnlyEmitDebounce,
    hotRef,
    isEditingRef,
    readOnly,
  ]);

  React.useEffect(
    () => () => {
      flushPendingColorTimers();
      flushReadOnlyEmitDebounce();
      if (previewEditingSettleTimerRef.current) {
        clearTimeout(previewEditingSettleTimerRef.current);
      }
      if (undoRedoRefreshTimerRef.current) {
        clearTimeout(undoRedoRefreshTimerRef.current);
      }
    },
    [flushPendingColorTimers, flushReadOnlyEmitDebounce],
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
            setCellMeta(activeSheetName, r, c, {
              type: "date",
              dateFormat: "YYYY-MM-DD",
              correctFormat: true,
            });
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
          setCellMeta(activeSheetName, r, c, {
            type: "numeric",
            numericFormat: { pattern: patterns[kind], culture: "en-US" },
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

  const { setYesNoToggle, setSingleCheckbox } = useWorkbookToolbarActions({
    hotRef,
    readOnly,
    activeSheetIndexRef,
    workbookRef,
    yesNoOppositeCellMapRef,
    getToolbarActionRange,
    collectCurrentSheetFromHot,
    scheduleUndoRedoRefresh,
    restoreHotRange,
    setScopedCellMeta: (sheetName, row, col, props) =>
      setCellMeta(sheetName, row, col, props as Partial<CellMetaEntry>),
  });

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
    const sheetName = sheet.name || `Sheet${idx + 1}`;

    const metaByKey = new Map<
      string,
      NonNullable<SheetData["cellMeta"]>[number]
    >();
    for (const meta of sheet.cellMeta || []) {
      metaByKey.set(`${meta.row}:${meta.col}`, meta);
    }

    let allSelectedAreFillable = true;
    for (let r = range.startRow; r <= range.endRow; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        const key = `${r}:${c}`;
        const current = metaByKey.get(key);
        const tokens = String(current?.className || "")
          .split(" ")
          .filter(Boolean);
        if (!tokens.includes("meta-fillable")) {
          allSelectedAreFillable = false;
          break;
        }
      }
      if (!allSelectedAreFillable) break;
    }
    const shouldSetFillable = !allSelectedAreFillable;

    for (let r = range.startRow; r <= range.endRow; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        const key = `${r}:${c}`;
        const current = metaByKey.get(key);
        const tokens = String(current?.className || "")
          .split(" ")
          .filter(Boolean)
          .filter((token) => token !== "meta-fillable");
        if (shouldSetFillable) tokens.push("meta-fillable");
        const newClassName = tokens.join(" ").trim();

        hot.setCellMeta(r, c, "className", newClassName);
        setCellMeta(sheetName, r, c, { className: newClassName || undefined });

        metaByKey.set(key, {
          ...(current || { row: r, col: c }),
          row: r,
          col: c,
          className: newClassName || undefined,
        });
      }
    }

    sheet.cellMeta = dedupeCellMetaByCoordinate(Array.from(metaByKey.values()));
    setSheetCellMetaFromList(sheetName, sheet.cellMeta);
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
        }
      });
    });
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

  // ─── clear cells ────────────────────────────────────────────────────────────

  const clearSelectedCells = () => {
    const hot = hotRef.current?.hotInstance;
    if (!hot || readOnly) return;
    const range = getToolbarActionRange(hot);
    if (!range) return;

    const clearKeys = new Set<string>();
    for (let r = range.startRow; r <= range.endRow; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        clearKeys.add(cellCoordKey(r, c));
      }
    }

    const changes: Array<[number, number, string]> = [];
    for (let r = range.startRow; r <= range.endRow; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        changes.push([r, c, ""]);
      }
    }
    if (changes.length) hot.setDataAtCell(changes);

    const stripHotMeta = () => {
      const clearMetaKey = (row: number, col: number, key: string) => {
        if (typeof hot.removeCellMeta === "function") {
          hot.removeCellMeta(row, col, key);
        } else {
          hot.setCellMeta(row, col, key, undefined);
        }
      };
      for (let r = range.startRow; r <= range.endRow; r++) {
        for (let c = range.startCol; c <= range.endCol; c++) {
          clearMetaKey(r, c, "className");
          clearMetaKey(r, c, "type");
          clearMetaKey(r, c, "checkedTemplate");
          clearMetaKey(r, c, "uncheckedTemplate");
          clearMetaKey(r, c, "numericFormat");
          clearMetaKey(r, c, "dateFormat");
          clearMetaKey(r, c, "correctFormat");
          clearMetaKey(r, c, "source");
          clearMetaKey(r, c, "strict");
          clearMetaKey(r, c, "renderer");
          clearMetaKey(r, c, "editor");
          clearMetaKey(r, c, "readOnly");
        }
      }
    };
    if (typeof hot.batch === "function") hot.batch(stripHotMeta);
    else stripHotMeta();

    const idx = activeSheetIndexRef.current;
    const sheet = workbookRef.current.sheets[idx];
    const sheetName = sheet?.name || `Sheet${idx + 1}`;
    if (sheet?.cellMeta) {
      sheet.cellMeta = sheet.cellMeta.filter(
        (m) => !clearKeys.has(cellCoordKey(+m.row, +m.col)),
      );
      setSheetCellMetaFromList(sheetName, sheet.cellMeta);
    }
    for (const key of clearKeys) {
      formulaCellSetRef.current.delete(key);
    }

    const hf = hfRef.current;
    const sheetId =
      hf && sheet ? hf.getSheetId(sheet.name || `Sheet${idx + 1}`) : null;
    if (hf && sheetId != null) {
      for (let r = range.startRow; r <= range.endRow; r++) {
        for (let c = range.startCol; c <= range.endCol; c++) {
          hf.setCellContents({ sheet: sheetId, row: r, col: c }, [[""]]);
        }
      }
      refreshFormulaDisplays();
    }

    collectCurrentSheetFromHot(true);
    scheduleUndoRedoRefresh();
    hot.render();
    restoreHotRange(hot, range);
    syncToolbarFromCell(hot, range.startRow, range.startCol);
    setFormulaInput("");
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

  // ─── formula bar ────────────────────────────────────────────────────────────

  const applyFormulaBar = () => {
    const hot = hotRef.current?.hotInstance;
    if (!hot || (readOnly && !canUseReadOnlyWorkbookActions)) return;
    const range = getToolbarActionRange(hot);
    if (!range) return;
    hot.setDataAtCell(range.startRow, range.startCol, formulaInput);
    collectCurrentSheetFromHot(false);
    restoreHotRange(hot, range);
  };

  // ─── export ─────────────────────────────────────────────────────────────────

  const flushLayoutToParent = React.useCallback(() => {
    if (readOnly) return;
    collectCurrentSheetFromHot(true);
  }, [readOnly, collectCurrentSheetFromHot]);

  const exportXlsx = async () => {
    if (!readOnly) collectCurrentSheetFromHot(true);
    const XLSX = await import("xlsx");
    const workbook = XLSX.utils.book_new();
    workbookRef.current.sheets.forEach((sheet) => {
      const ws = XLSX.utils.aoa_to_sheet(sheet.grid || [[""]]);
      for (const meta of sheet.cellMeta || []) {
        if (typeof (meta as any).formula !== "string") continue;
        const addr = XLSX.utils.encode_cell({ r: meta.row, c: meta.col });
        const wsCell = (ws as any)[addr] || {};
        wsCell.f = String((meta as any).formula).replace(/^=/, "");
        wsCell.v = String(sheet.grid?.[meta.row]?.[meta.col] ?? "");
        wsCell.t = "s";
        (ws as any)[addr] = wsCell;
      }
      XLSX.utils.book_append_sheet(workbook, ws, sheet.name || "Sheet");
    });
    const buf = XLSX.write(workbook, { type: "array", bookType: "xlsx" });
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

  const deleteActiveSheet = () => {
    const sheets = workbookRef.current.sheets;
    if (sheets.length <= 1) return;
    const idx = activeSheetIndexRef.current;
    const nextSheets = sheets.filter((_, i) => i !== idx);
    workbookRef.current.sheets = nextSheets;
    setSheetTabs(nextSheets.map((s) => ({ name: s.name, tabColor: s.tabColor })));
    const nextIdx = Math.min(idx, nextSheets.length - 1);
    const nextSheet = nextSheets[nextIdx];
    setInitialGrid(toVisibleGrid(nextSheet));
    setActiveSheetIndex(nextIdx);
    if (!readOnly) {
      emitWorkbookToParent();
    }
  };

  const duplicateActiveSheet = () => {
    if (!readOnly) collectCurrentSheetFromHot(true);
    const idx = activeSheetIndexRef.current;
    const source = workbookRef.current.sheets[idx];
    if (!source) return;
    const cloned = JSON.parse(JSON.stringify(source)) as SheetData;
    cloned.name = `${source.name} Copy`;
    const nextSheets = [...workbookRef.current.sheets, cloned];
    workbookRef.current.sheets = nextSheets;
    setSheetCellMetaFromList(
      cloned.name || `Sheet${nextSheets.length}`,
      cloned.cellMeta || [],
    );
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
    if (readOnly && isEditingRef.current) return;
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
    for (const s of workbookRef.current.sheets) {
      setSheetCellMetaFromList(s.name || "Sheet", s.cellMeta || []);
    }
    setSheetTabs(
      nextSheets.map((s) => ({ name: s.name, tabColor: s.tabColor })),
    );

    if (sheetCountChanged) {
      setActiveSheetIndex(0);
      const first = nextSheets[0]?.grid?.length ? nextSheets[0].grid : [[""]];
      setInitialGrid(cloneEditableGrid(first));
    } else {
      setActiveSheetIndex((prev) =>
        Math.min(prev, Math.max(0, nextSheets.length - 1)),
      );
    }
  }, [normalizedIncomingSheets, readOnly]);

  const hotTableMountKey = React.useMemo(
    () => hotTableMountSignature(normalizedIncomingSheets),
    [normalizedIncomingSheets],
  );

  React.useEffect(() => {
    // A sheet switch MUST always load the new sheet's data — never suppress it.
    // suppression only applies to parent-roundtrip reloads (same sheet, parent
    // reflected our emitted data back), which would redundantly reload the data
    // we just collected. Without this guard the removed-key optimisation would
    // block the only path that loads the new sheet into the persistent HOT instance.
    const sheetSwitched = activeSheetIndex !== lastLoadedSheetIndexRef.current;
    if (!sheetSwitched && suppressNextHotReloadRef.current) {
      suppressNextHotReloadRef.current = false;
      return;
    }
    suppressNextHotReloadRef.current = false;
    const needsReload =
      sheetSwitched ||
      incomingWorkbookKey !== lastLoadedWorkbookKeyRef.current;
    if (!needsReload) return;
    const hot = hotRef.current?.hotInstance;
    const isEditorOpen =
      typeof hot?.isEditorOpened === "function" && hot.isEditorOpened();
    if (readOnly && (isEditorOpen || isEditingRef.current)) {
      pendingIncomingReloadRef.current = true;
      pendingIncomingReloadSheetIndexRef.current = activeSheetIndex;
      pendingIncomingReloadWorkbookKeyRef.current = incomingWorkbookKey;
      return;
    }
    loadSheetIntoHot(activeSheetIndex);
  }, [activeSheetIndex, loadSheetIntoHot, incomingWorkbookKey, readOnly]);

  React.useEffect(() => {
    initializeHyperFormula();
    refreshFormulaDisplays();
    return () => {
      hfRef.current?.destroy();
      hfRef.current = null;
    };
  }, [incomingWorkbookKey, initializeHyperFormula, refreshFormulaDisplays]);

  // ─── cell renderer ───────────────────────────────────────────────────────────
  React.useEffect(() => {
    cellsCacheRef.current.clear();
    mergeCacheFrameRef.current = { frameId: -1, mergedSet: new Set() };
  }, [
    activeSheetIndex,
    persistedCellMetaMap,
    fillableCellSet,
    imageMap,
    readOnly,
    renderedColWidths,
    renderedRowHeights,
    renderedMergeCells,
  ]);

  const cellsCallback = React.useCallback(
    (row: number, col: number) => {
      const cacheKey = cellCoordKey(row, col);
      const cached = cellsCacheRef.current.get(cacheKey);
      if (cached) return cached;
      const hot = hotRef.current?.hotInstance;
      const currentFrame = (hot as any)?._renderCount ?? 0;
      if (mergeCacheFrameRef.current.frameId !== currentFrame) {
        const mergePlugin = hot?.getPlugin?.("mergeCells");
        const allMerges: any[] =
          mergePlugin?.mergedCellsCollection?.mergedCells || [];
        const covered = new Set<string>();
        for (const m of allMerges) {
          for (let r = m.row; r < m.row + m.rowspan; r++) {
            for (let c = m.col; c < m.col + m.colspan; c++) {
              if (r !== m.row || c !== m.col) {
                covered.add(`${r}:${c}`);
              }
            }
          }
        }
        mergeCacheFrameRef.current = {
          frameId: currentFrame,
          mergedSet: covered,
        };
      }
      if (mergeCacheFrameRef.current.mergedSet.has(cacheKey)) {
        cellsCacheRef.current.set(cacheKey, {});
        return {};
      }
      const persistedMeta =
        persistedCellMetaMap.get(cacheKey) ||
        (getCellMeta(activeSheetName, row, col) as CellMetaEntry | undefined);
      const cp: any = {};
      const persistedClassName = String(persistedMeta?.className || "");
      const classTokens = persistedClassName.split(" ").filter(Boolean);
      const isYesNoCheckboxCell = Boolean(
        extractYesNoPairToken(persistedClassName),
      );
      const isSingleCheckboxCell = classTokens.includes(SINGLE_CHECKBOX_CLASS);
      const isFillable = readOnly
        ? fillableCellSet.has(cellCoordKey(row, col))
        : classTokens.includes("meta-fillable");
      cp.readOnly = readOnly ? !isFillable : false;
      if (persistedMeta?.className) cp.className = persistedMeta.className;
      if (persistedMeta?.type) cp.type = persistedMeta.type;
      if (persistedMeta?.type === "checkbox") {
        cp.type = "checkbox";
        cp.checkedTemplate =
          persistedMeta.checkedTemplate !== undefined
            ? persistedMeta.checkedTemplate
            : "true";
        cp.uncheckedTemplate =
          persistedMeta.uncheckedTemplate !== undefined
            ? persistedMeta.uncheckedTemplate
            : "";
      }
      if (isSingleCheckboxCell) {
        cp.type = "checkbox";
        if (cp.checkedTemplate === undefined) cp.checkedTemplate = "true";
        if (cp.uncheckedTemplate === undefined) cp.uncheckedTemplate = "";
      }
      if (isYesNoCheckboxCell) {
        cp.type = "checkbox";
        if (cp.checkedTemplate === undefined) cp.checkedTemplate = "YES";
        if (cp.uncheckedTemplate === undefined) cp.uncheckedTemplate = "NO";
      }
      if (persistedMeta?.dateFormat) cp.dateFormat = persistedMeta.dateFormat;
      if (typeof persistedMeta?.correctFormat === "boolean")
        cp.correctFormat = persistedMeta.correctFormat;
      if (persistedMeta?.numericFormat)
        cp.numericFormat = persistedMeta.numericFormat;
      if (Array.isArray(persistedMeta?.source))
        cp.source = persistedMeta.source;
      if (typeof persistedMeta?.strict === "boolean")
        cp.strict = persistedMeta.strict;

      const image = imageMap.get(cacheKey);
      cp.renderer = (
          instance: any,
          td: HTMLTableCellElement,
          rowIndex: number,
          colIndex: number,
          prop: any,
          value: any,
          cellProperties: any,
        ) => {
          if (!td) return td;
          const isCheckboxCell =
            String(cellProperties?.type || "") === "checkbox";
          if (isCheckboxCell) {
            checkboxRenderer(
              instance,
              td,
              rowIndex,
              colIndex,
              prop,
              value,
              cellProperties,
            );
            return td;
          }

          textRenderer(
            instance,
            td,
            rowIndex,
            colIndex,
            prop,
            (() => {
              const currentSheet =
                workbookRef.current.sheets[activeSheetIndexRef.current];
              const formulaMeta = getFormulaMeta(
                currentSheet,
                rowIndex,
                colIndex,
              ) as CellMetaEntry | undefined;
              if (
                formulaMeta &&
                typeof formulaMeta.formula === "string" &&
                formulaMeta.formula.startsWith(FORMULA_PREFIX)
              ) {
                if (
                  typeof formulaMeta.formulaCachedValue === "string" &&
                  (formulaMeta.formulaCachedValue !== "" ||
                    value == null ||
                    String(value) === "")
                ) {
                  return formulaMeta.formulaCachedValue;
                }
                return value;
              }
              return value;
            })(),
            cellProperties,
          );

          const cls = String(cellProperties?.className || "");
          const tokens = cls.split(" ").filter(Boolean);

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
          const isFillableCell = tokens.includes("meta-fillable");
          if (fillToken && !(readOnly && isFillableCell))
            s.backgroundColor = `#${fillToken.replace("meta-fill-", "")}`;
          if (readOnly && isFillableCell) s.backgroundColor = "#dbeafe";
          if (alignToken) s.textAlign = alignToken.replace("meta-align-", "");
          if (vAlignToken)
            s.verticalAlign = vAlignToken.replace("meta-valign-", "");
          if (tokens.includes("meta-wrap")) s.whiteSpace = "normal";

          const selectedRange = instance?.getSelectedRangeLast?.();
          const from = selectedRange?.from;
          const to = selectedRange?.to;
          if (
            from &&
            to &&
            Number.isInteger(from.row) &&
            Number.isInteger(from.col) &&
            Number.isInteger(to.row) &&
            Number.isInteger(to.col)
          ) {
            const minRow = Math.min(from.row, to.row);
            const maxRow = Math.max(from.row, to.row);
            const minCol = Math.min(from.col, to.col);
            const maxCol = Math.max(from.col, to.col);
            const inSelection =
              rowIndex >= minRow &&
              rowIndex <= maxRow &&
              colIndex >= minCol &&
              colIndex <= maxCol;
            if (inSelection) {
              s.backgroundColor = "rgba(26, 115, 232, 0.14)";
            }
          }

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
      if (disableEditorCompletely) {
        cp.editor = false;
      }
      if (disableEditorCompletely) cp.readOnly = true;
      cellsCacheRef.current.set(cacheKey, cp);
      return cp;
    },
    [
      persistedCellMetaMap,
      activeSheetName,
      getCellMeta,
      fillableCellSet,
      imageMap,
      readOnly,
      disableEditorCompletely,
      renderedColWidths,
      renderedRowHeights,
    ],
  );

  const afterGetCellMeta = React.useCallback(
    (row: number, col: number, cellProps: Record<string, unknown>) => {
      if (disableEditorCompletely) {
        (cellProps as { readOnly?: boolean; editor?: false }).readOnly = true;
        (cellProps as { readOnly?: boolean; editor?: false }).editor = false;
        return;
      }
      if (!readOnly) {
        (cellProps as { readOnly?: boolean }).readOnly = false;
        return;
      }
      if (formulaCellSetRef.current.has(cellCoordKey(row, col))) {
        (cellProps as { readOnly?: boolean }).readOnly = false;
        return;
      }
      const isFillable = fillableCellSet.has(cellCoordKey(row, col));
      (cellProps as { readOnly?: boolean }).readOnly = !isFillable;
    },
    [readOnly, fillableCellSet, disableEditorCompletely],
  );

  const clearEditorTextLayoutListeners = React.useCallback(() => {
    editorTextLayoutCleanupRef.current?.();
    editorTextLayoutCleanupRef.current = null;
  }, []);

  React.useLayoutEffect(
    () => () => {
      clearEditorTextLayoutListeners();
      if (pendingEditorLayoutRafRef.current != null) {
        cancelAnimationFrame(pendingEditorLayoutRafRef.current);
        pendingEditorLayoutRafRef.current = null;
      }
      if (pendingEditorLayoutNestedRafRef.current != null) {
        cancelAnimationFrame(pendingEditorLayoutNestedRafRef.current);
        pendingEditorLayoutNestedRafRef.current = null;
      }
    },
    [clearEditorTextLayoutListeners, activeSheetIndex],
  );

  const syncEditorIfOpen = React.useCallback(() => {
    const hot = hotRef.current?.hotInstance;
    if (!hot || readOnly || disableEditorCompletely) return;
    if (typeof hot.isEditorOpened !== "function" || !hot.isEditorOpened())
      return;
    syncHandsontableTextEditorToCell(hot);
    const editor = hot.getActiveEditor?.() as HotTextEditorLike | undefined;
    const grow = editor?.TEXTAREA && (editor.TEXTAREA as any).__htGrowWrap;
    if (typeof grow === "function") grow();
  }, [readOnly, disableEditorCompletely]);

  const afterBeginEditingForCellLayout = React.useCallback(() => {
    if (readOnly || disableEditorCompletely) return;
    clearEditorTextLayoutListeners();
    const hot = hotRef.current?.hotInstance;
    if (!hot) return;
    const setup = () => {
      if (typeof hot.isEditorOpened !== "function" || !hot.isEditorOpened())
        return;
      const editor = hot.getActiveEditor?.() as HotTextEditorLike | undefined;
      const ta = editor?.TEXTAREA;
      if (!ta) return;
      syncHandsontableTextEditorToCell(hot);
      const onInput = () => {
        const g = (ta as any).__htGrowWrap as (() => void) | undefined;
        if (typeof g === "function") g();
      };
      const onBlur = () => clearEditorTextLayoutListeners();
      ta.addEventListener("input", onInput);
      ta.addEventListener("blur", onBlur);
      editorTextLayoutCleanupRef.current = () => {
        ta.removeEventListener("input", onInput);
        ta.removeEventListener("blur", onBlur);
      };
    };
    if (pendingEditorLayoutRafRef.current != null) {
      cancelAnimationFrame(pendingEditorLayoutRafRef.current);
      pendingEditorLayoutRafRef.current = null;
    }
    if (pendingEditorLayoutNestedRafRef.current != null) {
      cancelAnimationFrame(pendingEditorLayoutNestedRafRef.current);
      pendingEditorLayoutNestedRafRef.current = null;
    }
    pendingEditorLayoutRafRef.current = requestAnimationFrame(() => {
      pendingEditorLayoutRafRef.current = null;
      pendingEditorLayoutNestedRafRef.current = requestAnimationFrame(() => {
        pendingEditorLayoutNestedRafRef.current = null;
        setup();
      });
    });
  }, [readOnly, disableEditorCompletely, clearEditorTextLayoutListeners]);

  React.useEffect(() => {
    return () => {
      if (pendingScrollRestoreRafRef.current != null) {
        cancelAnimationFrame(pendingScrollRestoreRafRef.current);
        pendingScrollRestoreRafRef.current = null;
      }
      if (pendingScrollRestoreNestedRafRef.current != null) {
        cancelAnimationFrame(pendingScrollRestoreNestedRafRef.current);
        pendingScrollRestoreNestedRafRef.current = null;
      }
    };
  }, [activeSheetIndex]);

  const afterColumnResize = React.useCallback(() => {
    flushLayoutToParent();
    syncEditorIfOpen();
  }, [flushLayoutToParent, syncEditorIfOpen]);

  const afterRowResize = React.useCallback(() => {
    flushLayoutToParent();
    syncEditorIfOpen();
  }, [flushLayoutToParent, syncEditorIfOpen]);

  const afterMergeCells = React.useCallback(() => {
    if (!readOnly) {
      scheduleUndoRedoRefresh();
      setTimeout(() => collectCurrentSheetFromHot(true), 0);
    }
  }, [readOnly, collectCurrentSheetFromHot, scheduleUndoRedoRefresh]);

  const afterUnmergeCells = React.useCallback(() => {
    if (!readOnly) {
      scheduleUndoRedoRefresh();
      setTimeout(() => collectCurrentSheetFromHot(true), 0);
    }
  }, [readOnly, collectCurrentSheetFromHot, scheduleUndoRedoRefresh]);

  const afterCreateRow = React.useCallback(() => {
    rowStructureDirtyRef.current.set(activeSheetIndexRef.current, true);
    scheduleUndoRedoRefresh();
  }, [activeSheetIndexRef, scheduleUndoRedoRefresh]);

  const afterCreateCol = React.useCallback(() => {
    columnStructureDirtyRef.current.set(activeSheetIndexRef.current, true);
    scheduleUndoRedoRefresh();
  }, [scheduleUndoRedoRefresh]);

  const afterRemoveRow = React.useCallback(() => {
    rowStructureDirtyRef.current.set(activeSheetIndexRef.current, true);
    scheduleUndoRedoRefresh();
  }, [activeSheetIndexRef, scheduleUndoRedoRefresh]);

  const afterRemoveCol = React.useCallback(() => {
    columnStructureDirtyRef.current.set(activeSheetIndexRef.current, true);
    scheduleUndoRedoRefresh();
  }, [scheduleUndoRedoRefresh]);

  const handleCellChanges = React.useCallback(
    (
      changes: [number, number, unknown, unknown][],
      _source?: string,
      sheetIndexOverride?: number,
    ) => {
      const sheetIndex =
        Number.isInteger(sheetIndexOverride) &&
        (sheetIndexOverride as number) >= 0
          ? (sheetIndexOverride as number)
          : activeSheetIndexRef.current;
      const sheet = workbookRef.current.sheets[sheetIndex];
      const hf = hfRef.current;
      if (!sheet || !hf || !Array.isArray(changes)) return;
      const sheetId = hf.getSheetId(sheet.name || `Sheet${sheetIndex + 1}`);
      if (sheetId == null) return;
      const metaByKey = new Map<string, CellMetaEntry>();
      for (const m of sheet.cellMeta || []) {
        metaByKey.set(cellCoordKey(m.row, m.col), { ...m });
      }
      for (const [row, col, , newValue] of changes) {
        if (!Number.isFinite(row) || !Number.isFinite(col)) continue;
        const valueText = newValue == null ? "" : String(newValue);
        const key = cellCoordKey(row, col);
        const current = metaByKey.get(key) || ({ row, col } as CellMetaEntry);
        if (valueText.startsWith(FORMULA_PREFIX)) {
          current.formula = valueText;
          current.formulaCachedValue = String(sheet.grid?.[row]?.[col] ?? "");
          metaByKey.set(key, current);
          hf.setCellContents({ sheet: sheetId, row, col }, [[valueText]]);
        } else {
          delete (current as any).formula;
          delete (current as any).formulaCachedValue;
          metaByKey.set(key, current);
          hf.setCellContents({ sheet: sheetId, row, col }, [[valueText]]);
        }
      }
      sheet.cellMeta = dedupeCellMetaByCoordinate([...metaByKey.values()]);
      const updatesBySheet = refreshFormulaDisplays();
      // Re-sync cellMetaRef after refreshFormulaDisplays has written the
      // correct formulaCachedValues into the sheet.cellMeta objects.
      setSheetCellMetaFromList(sheet.name || `Sheet${sheetIndex + 1}`, sheet.cellMeta);
      // Keep formulaCellSetRef in sync so the cells callback marks newly
      // entered formula cells as editable without requiring a full reload.
      if (sheetIndex === activeSheetIndexRef.current) {
        const nextFormulaSet = new Set<string>();
        for (const m of sheet.cellMeta) {
          if (
            typeof (m as any).formula === "string" &&
            (m as any).formula.startsWith(FORMULA_PREFIX)
          ) {
            nextFormulaSet.add(cellCoordKey(m.row, m.col));
          }
        }
        formulaCellSetRef.current = nextFormulaSet;
      }
      const visibleSheetIndex = activeSheetIndexRef.current;
      const activeUpdates = updatesBySheet.get(visibleSheetIndex) || [];
      const hot = hotRef.current?.hotInstance;
      if (hot && activeUpdates.length > 0) {
        hot.render();
      }
    },
    [activeSheetIndexRef, workbookRef, refreshFormulaDisplays, hotRef, setSheetCellMetaFromList],
  );

  const { afterChange } = useWorkbookHotCallbacks({
    hotRef,
    yesNoOppositeCellMapRef,
    readOnly,
    scheduleUndoRedoRefresh,
    activeSheetIndexRef,
    workbookRef,
    readOnlyPreviewDirtyRef,
    isEditingRef,
    pendingReadOnlyEmitRef,
    onReadOnlyEdit: scheduleReadOnlyEmit,
    onCellChanges: handleCellChanges,
  });

  const flushPendingPreviewSyncs = React.useCallback(() => {
    if (pendingReadOnlyEmitRef.current) {
      pendingReadOnlyEmitRef.current = false;
      scheduleReadOnlyEmit();
    }
    if (pendingIncomingReloadRef.current) {
      const pendingSheetIndex =
        pendingIncomingReloadSheetIndexRef.current ??
        activeSheetIndexRef.current;
      const pendingWorkbookKey =
        pendingIncomingReloadWorkbookKeyRef.current ?? incomingWorkbookKey;
      pendingIncomingReloadRef.current = false;
      pendingIncomingReloadSheetIndexRef.current = null;
      pendingIncomingReloadWorkbookKeyRef.current = null;
      const stillNeedsReload =
        pendingSheetIndex !== lastLoadedSheetIndexRef.current ||
        pendingWorkbookKey !== lastLoadedWorkbookKeyRef.current;
      if (stillNeedsReload) {
        loadSheetIntoHot(pendingSheetIndex);
      }
    }
  }, [incomingWorkbookKey, loadSheetIntoHot, scheduleReadOnlyEmit]);

  const schedulePreviewEditingSettle = React.useCallback(() => {
    if (!readOnly) return;
    if (previewEditingSettleTimerRef.current) {
      clearTimeout(previewEditingSettleTimerRef.current);
    }
    previewEditingSettleTimerRef.current = setTimeout(() => {
      previewEditingSettleTimerRef.current = null;
      const hot = hotRef.current?.hotInstance;
      const editorOpen =
        typeof hot?.isEditorOpened === "function" && hot.isEditorOpened();
      if (editorOpen) return;
      isEditingRef.current = false;
      flushPendingPreviewSyncs();
    }, 160);
  }, [flushPendingPreviewSyncs, readOnly]);

  const beforeBeginEditing = React.useCallback(() => {
    if (!readOnly) return;
    if (previewEditingSettleTimerRef.current) {
      clearTimeout(previewEditingSettleTimerRef.current);
      previewEditingSettleTimerRef.current = null;
    }
    isEditingRef.current = true;
  }, [readOnly]);

  const afterDeselect = React.useCallback(() => {
    if (!readOnly) return;
    schedulePreviewEditingSettle();
  }, [readOnly, schedulePreviewEditingSettle]);

  const afterChangeWithEditTracking = React.useCallback(
    (changes: any, source: string) => {
      afterChange(changes, source);
      if (!readOnly) return;
      if (source === "afterAutofill" || source === "Autofill.fill") {
        isEditingRef.current = false;
      }
      schedulePreviewEditingSettle();
    },
    [afterChange, readOnly, schedulePreviewEditingSettle],
  );

  const afterSelection = React.useCallback(
    (
      r: number,
      c: number,
      r2: number,
      c2: number,
      preventScrolling?: { value: boolean },
    ) => {
      const hot = hotRef.current?.hotInstance;
      if (
        readOnly &&
        preventScrolling &&
        typeof preventScrolling.value === "boolean" &&
        hot &&
        typeof hot.selection?.getSelectionSource === "function" &&
        hot.selection.getSelectionSource() !== "keyboard"
      ) {
        preventScrolling.value = true;
      }
      const rowCount =
        typeof hot?.countRows === "function"
          ? Math.max(1, hot.countRows())
          : Math.max(1, safeGrid.length);
      const colCount =
        typeof hot?.countCols === "function"
          ? Math.max(1, hot.countCols())
          : Math.max(1, safeGrid[0]?.length || 1);

      if (!Number.isInteger(r) || !Number.isInteger(c)) return;
      const endRowRaw = Number.isInteger(r2) ? r2 : r;
      const endColRaw = Number.isInteger(c2) ? c2 : c;
      const startRowRaw = r < 0 ? 0 : r;
      const endRowNormalized = endRowRaw < 0 ? rowCount - 1 : endRowRaw;
      const startColRaw = c < 0 ? 0 : c;
      const endColNormalized = endColRaw < 0 ? colCount - 1 : endColRaw;
      const range = {
        startRow: Math.max(0, Math.min(startRowRaw, endRowNormalized)),
        endRow: Math.min(rowCount - 1, Math.max(startRowRaw, endRowNormalized)),
        startCol: Math.max(0, Math.min(startColRaw, endColNormalized)),
        endCol: Math.min(colCount - 1, Math.max(startColRaw, endColNormalized)),
      };
      lastSelectionRef.current = range;
      sheetSelectionRef.current[activeSheetIndexRef.current] = range;
      if (!readOnly) setSelectionLabel(toRangeLabel(range));
    },
    [readOnly, safeGrid],
  );

  const afterSelectionFocusSet = React.useCallback(
    (_row: number, _col: number, preventScrolling?: { value: boolean }) => {
      const hot = hotRef.current?.hotInstance;
      if (
        readOnly &&
        preventScrolling &&
        typeof preventScrolling.value === "boolean" &&
        hot &&
        typeof hot.selection?.getSelectionSource === "function" &&
        hot.selection.getSelectionSource() !== "keyboard"
      ) {
        preventScrolling.value = true;
      }
    },
    [readOnly],
  );

  const afterSelectionEnd = React.useCallback(
    (r: number, c: number, r2: number, c2: number) => {
      const hot = hotRef.current?.hotInstance;
      if (!hot || !Number.isInteger(r) || !Number.isInteger(c)) return;
      const rowCount =
        typeof hot.countRows === "function"
          ? Math.max(1, hot.countRows())
          : Math.max(1, safeGrid.length);
      const colCount =
        typeof hot.countCols === "function"
          ? Math.max(1, hot.countCols())
          : Math.max(1, safeGrid[0]?.length || 1);
      const endRowRaw = Number.isInteger(r2) ? r2 : r;
      const endColRaw = Number.isInteger(c2) ? c2 : c;
      const startRowRaw = r < 0 ? 0 : r;
      const endRowNormalized = endRowRaw < 0 ? rowCount - 1 : endRowRaw;
      const startColRaw = c < 0 ? 0 : c;
      const endColNormalized = endColRaw < 0 ? colCount - 1 : endColRaw;
      const range = {
        startRow: Math.max(0, Math.min(startRowRaw, endRowNormalized)),
        endRow: Math.min(rowCount - 1, Math.max(startRowRaw, endRowNormalized)),
        startCol: Math.max(0, Math.min(startColRaw, endColNormalized)),
        endCol: Math.min(colCount - 1, Math.max(startColRaw, endColNormalized)),
      };
      lastSelectionRef.current = range;
      sheetSelectionRef.current[activeSheetIndexRef.current] = range;

      if (readOnly) {
        return;
      }
      setSelectionLabel(toRangeLabel(range));
      syncToolbarFromCell(hot, range.startRow, range.startCol);
    },
    [readOnly, safeGrid, syncToolbarFromCell],
  );

  const hotTableContextMenu = React.useMemo<any>(() => {
    if (readOnly) return false;
    return {
      ...(menuContainer ? { uiContainer: menuContainer } : {}),
      items: {
        row_above: {},
        row_below: {},
        col_left: {},
        col_right: {},
        hsep1: "---------",
        remove_row: {},
        remove_col: {},
        clear_column: {},
        hidden_rows_hide: {},
        hidden_rows_show: {},
        hidden_columns_hide: {},
        hidden_columns_show: {},
        hsep2: "---------",
        mergeCells: {},
        hsep3: "---------",
        read_only: {},
        alignment: {},
        row_height: {},
        col_width: {},
        freeze_column: {},
        unfreeze_column: {},
        hsep4: "---------",
        copy: {},
        cut: {},
        hsep5: "---------",
        undo: {},
        redo: {},
      },
    };
  }, [readOnly, menuContainer]);

  const heavyPluginsEnabled = !readOnly && !lightweightPerformance;
  const hotTableSettings = React.useMemo(
    () => ({
      data: initialGrid,
      rowHeaders: true,
      colHeaders: true,
      selectionMode: "multiple" as const,
      licenseKey: "non-commercial-and-evaluation" as const,
      readOnly: disableEditorCompletely ? true : false,
      disableVisualSelection: disableEditorCompletely ? false : undefined,
      editor: disableEditorCompletely ? false : undefined,
      beforeKeyDown: disableEditorCompletely
        ? (e: KeyboardEvent) => {
            e.stopImmediatePropagation();
          }
        : undefined,
      trimWhitespace: false,
      stretchH: (stretchColumnsInPreview ? "all" : "none") as "all" | "none",
      height: readOnly ? (readOnlyHotHeight ?? 380) : 320,
      renderAllRows: readOnly,
      viewportRowRenderingOffset: lightweightPerformance ? 8 : 20,
      viewportColumnRenderingOffset: lightweightPerformance ? 4 : 10,
      // We run a workbook-wide HyperFormula instance manually for cross-sheet
      // references. Enabling HOT's per-grid formulas plugin here causes
      // cross-sheet refs (e.g. 'CREW LIST'!C4) to resolve as #REF!.
      formulas: undefined,
      mergeCells:
        renderedMergeCells.length > 0 ? renderedMergeCells : !readOnly,
      filters: heavyPluginsEnabled,
      dropdownMenu: heavyPluginsEnabled
        ? menuContainer
          ? { uiContainer: menuContainer }
          : true
        : false,
      columnSorting: !readOnly,
      manualColumnMove: !readOnly,
      hiddenRows: !readOnly ? ({ indicators: true } as const) : undefined,
      hiddenColumns: !readOnly ? ({ indicators: true } as const) : undefined,
      multiColumnSorting: !readOnly,
      manualColumnFreeze: !readOnly,
      autoColumnSize: false,
      autoRowSize: false,
      fillHandle: !readOnly,
      fixedRowsTop: 0,
      fixedColumnsStart: 0,
      contextMenu: hotTableContextMenu,
      afterContextMenuShow: menuContainer
        ? (contextMenuPlugin: any) => {
            const menu = contextMenuPlugin?.menu;
            if (!menu || (menu as any).__htDialogPositionPatched) return;
            (menu as any).__htDialogPositionPatched = true;
            const mc = menuContainer;
            const orig = (menu.setPosition as (c: any) => void).bind(menu);
            menu.setPosition = (coords: any) => {
              orig(coords);
              const dr = mc.getBoundingClientRect();
              const t = parseFloat(menu.container.style.top) || 0;
              const l = parseFloat(menu.container.style.left) || 0;
              menu.container.style.top = `${t - window.scrollY - dr.top}px`;
              menu.container.style.left = `${l - window.scrollX - dr.left}px`;
            };
          }
        : undefined,
      afterDropdownMenuShow: menuContainer
        ? (dropdownMenuPlugin: any) => {
            const menu = dropdownMenuPlugin?.menu;
            if (!menu || (menu as any).__htDialogPositionPatched) return;
            (menu as any).__htDialogPositionPatched = true;
            const mc = menuContainer;
            const orig = (menu.setPosition as (c: any) => void).bind(menu);
            menu.setPosition = (coords: any) => {
              orig(coords);
              const dr = mc.getBoundingClientRect();
              const t = parseFloat(menu.container.style.top) || 0;
              const l = parseFloat(menu.container.style.left) || 0;
              menu.container.style.top = `${t - window.scrollY - dr.top}px`;
              menu.container.style.left = `${l - window.scrollX - dr.left}px`;
            };
          }
        : undefined,
      wordWrap: true,
      autoWrapRow: true,
      autoWrapCol: true,
      cells: cellsCallback,
      afterGetCellMeta,
      beforeChange: disableEditorCompletely ? () => false : undefined,
      afterColumnResize,
      afterRowResize,
      afterChange: disableEditorCompletely
        ? undefined
        : afterChangeWithEditTracking,
      afterSelection,
      afterSelectionFocusSet: readOnly ? afterSelectionFocusSet : undefined,
      afterSelectionEnd,
      beforeBeginEditing: disableEditorCompletely
        ? undefined
        : beforeBeginEditing,
      afterBeginEditing:
        readOnly || disableEditorCompletely
          ? undefined
          : afterBeginEditingForCellLayout,
      afterScrollVertically:
        readOnly || disableEditorCompletely ? undefined : syncEditorIfOpen,
      afterScrollHorizontally:
        readOnly || disableEditorCompletely ? undefined : syncEditorIfOpen,
      afterScroll:
        readOnly || disableEditorCompletely ? undefined : syncEditorIfOpen,
      afterDeselect: disableEditorCompletely ? undefined : afterDeselect,
      afterMergeCells,
      afterUnmergeCells,
      afterCreateRow,
      afterCreateCol,
      afterRemoveRow,
      afterRemoveCol,
    }),
    [
      initialGrid,
      stretchColumnsInPreview,
      readOnly,
      readOnlyHotHeight,
      shouldUseFormulaEngine,
      renderedMergeCells,
      heavyPluginsEnabled,
      hotTableContextMenu,
      menuContainer,
      lightweightPerformance,
      disableEditorCompletely,
      cellsCallback,
      afterGetCellMeta,
      afterColumnResize,
      afterRowResize,
      afterChangeWithEditTracking,
      afterSelection,
      afterSelectionFocusSet,
      afterSelectionEnd,
      beforeBeginEditing,
      afterBeginEditingForCellLayout,
      syncEditorIfOpen,
      afterDeselect,
      afterMergeCells,
      afterUnmergeCells,
      afterCreateRow,
      afterCreateCol,
      afterRemoveRow,
      afterRemoveCol,
    ],
  );

  // ─── render ──────────────────────────────────────────────────────────────────

  return (
    <div
      className="space-y-2"
      onBlur={(e) => {
        if (!readOnly) return;
        const next = e.relatedTarget as Node | null;
        if (next && e.currentTarget.contains(next)) return;
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
        .ht-theme-main {
          --ht-cell-selection-background-color: #1a73e8;
          --ht-cell-selection-border-color: #1a73e8;
        }
        .handsontable tr td.current { background-color: rgba(26, 115, 232, 0.05) !important; }
        .hot-wrapper {
          width: 100%;
        }
        .hot-wrapper .handsontable td.current,
        .hot-wrapper .handsontable td.area,
        .hot-wrapper .handsontable td[class*="area-"] {
          background-color: rgba(26, 115, 232, 0.14) !important;
        }
        .hot-wrapper .handsontable td.area::before,
        .hot-wrapper .handsontable td[class*="area-"]::before {
          background-color: rgba(26, 115, 232, 0.18) !important;
          opacity: 1 !important;
        }
        .hot-wrapper .handsontable .ht__highlight,
        .hot-wrapper .handsontable .ht__active_highlight {
          background-color: rgba(26, 115, 232, 0.14) !important;
        }
        .hot-wrapper .handsontable .wtBorder,
        .hot-wrapper .handsontable .wtBorder div {
          background-color: #1a73e8 !important;
          opacity: 1 !important;
        }
      `}</style>

      {/* ── Toolbar ── */}
      {!readOnly && (
        <div className="relative z-10 flex flex-wrap items-center gap-1 p-2 border rounded-md bg-slate-50">
          <span
            className="px-2 text-xs font-mono font-semibold border rounded bg-white min-w-[3.5rem] text-center select-none"
            title="Active cell / selection"
            onMouseDown={noFocusSteal}
          >
            {selectionLabel}
          </span>

          <TB onClick={undoAction} disabled={!canUndo} title="Undo (Ctrl+Z)">
            ↩
          </TB>
          <TB onClick={redoAction} disabled={!canRedo} title="Redo (Ctrl+Y)">
            ↪
          </TB>
          <TB
            onClick={clearSelectedCells}
            title="Clear selected cells — removes content and all formatting"
            className="border-red-400 text-red-600 hover:bg-red-50"
          >
            Clear Cell
          </TB>

          <span className="mx-1 h-6 border-l" />

          <TB
            onClick={() => setFontStyle("bold")}
            active={isBoldActive}
            title="Bold (Ctrl+B)"
          >
            <b>B</b>
          </TB>
          <TB
            onClick={() => setFontStyle("italic")}
            active={isItalicActive}
            title="Italic (Ctrl+I)"
          >
            <i>I</i>
          </TB>
          <TB
            onClick={() => setFontStyle("underline")}
            active={isUnderlineActive}
            title="Underline (Ctrl+U)"
          >
            <u>U</u>
          </TB>
          <TB
            onClick={() => setFontStyle("strike")}
            active={isStrikeActive}
            title="Strikethrough"
          >
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
            title="Font family (auto-applies)"
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
            title="Font size px (auto-applies)"
            className="w-14 h-8 px-2 text-sm border rounded"
          />

          <span className="mx-1 h-6 border-l" />

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
            title="Text color (auto-applies on pick)"
          />
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
            title="Fill / background color (auto-applies on pick)"
          />

          <span className="mx-1 h-6 border-l" />

          <TB
            onClick={() => setAlignment("left")}
            active={selectedAlign === "left"}
            title="Align left"
          >
            ◀≡
          </TB>
          <TB
            onClick={() => setAlignment("center")}
            active={selectedAlign === "center"}
            title="Align center"
          >
            ≡
          </TB>
          <TB
            onClick={() => setAlignment("right")}
            active={selectedAlign === "right"}
            title="Align right"
          >
            ≡▶
          </TB>
          <TB
            onClick={() => setAlignment("justify")}
            active={selectedAlign === "justify"}
            title="Justify"
          >
            ⇔
          </TB>
          <TB
            onClick={() => setVerticalAlignment("bottom")}
            active={selectedVAlign === "bottom"}
            title="Align bottom"
          >
            ⊥
          </TB>

          <span className="mx-1 h-6 border-l" />

          <TB onClick={mergeSelection} title="Merge selected cells">
            ⊞ Merge
          </TB>
          <TB onClick={unmergeSelection} title="Unmerge selected cells">
            ⊟ Split
          </TB>

          <span className="mx-1 h-6 border-l" />

          <TB
            onClick={() => alterBySelection("insert_row_above")}
            title="Insert row above selection"
          >
            ↑ Row
          </TB>
          <TB
            onClick={() => alterBySelection("insert_row_below")}
            title="Insert row below selection"
          >
            ↓ Row
          </TB>
          <TB
            onClick={() => alterBySelection("insert_col_start")}
            title="Insert column to the left"
          >
            ← Col
          </TB>
          <TB
            onClick={() => alterBySelection("insert_col_end")}
            title="Insert column to the right"
          >
            → Col
          </TB>
          <TB
            onClick={() => alterBySelection("remove_row")}
            title="Delete selected rows"
          >
            ✕ Row
          </TB>
          <TB
            onClick={() => alterBySelection("remove_col")}
            title="Delete selected columns"
          >
            ✕ Col
          </TB>

          <span className="mx-1 h-6 border-l" />

          <TB
            onClick={() => formatSelectedAs("percent")}
            title="Format selection as percentage (0.00%)"
          >
            %
          </TB>

          <span className="mx-1 h-6 border-l" />

          <TB onClick={exportXlsx} title="Export workbook as Excel (.xlsx)">
            ↓ xlsx
          </TB>
          <TB onClick={exportCsv} title="Export active sheet as CSV">
            ↓ csv
          </TB>
        </div>
      )}

      {showFormulaBar && (
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

        {showSheetActions && (
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

            {!readOnly && (
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
            )}

            <TB onClick={duplicateActiveSheet} title="Duplicate active sheet">
              ⧉ Duplicate
            </TB>
            <TB onClick={() => moveSheet("left")} title="Move sheet left">
              ← Move
            </TB>
            <TB onClick={() => moveSheet("right")} title="Move sheet right">
              Move →
            </TB>
            <TB
              onClick={deleteActiveSheet}
              title="Delete active sheet"
              disabled={sheetTabs.length <= 1}
            >
              🗑 Delete Sheet
            </TB>

            {!readOnly && (
              <>
                <input
                  type="color"
                  className="w-8 h-8 p-0 border rounded cursor-pointer"
                  title="Sheet tab color"
                  onMouseDown={noFocusSteal}
                  onChange={(e) => applySheetColor(e.target.value)}
                />

                <span className="mx-1 h-6 border-l" />

                <TB
                  onClick={setYesNoToggle}
                  title="Select 2 side-by-side cells to create mutually exclusive YES/NO toggle checkboxes"
                >
                  ☑ YES/NO
                </TB>
                <TB
                  onClick={setSingleCheckbox}
                  title="Insert a standalone checkbox in each selected cell"
                >
                  ☑ Checkbox
                </TB>
                <TB
                  onClick={toggleFillableSelection}
                  title="Toggle selected cells between fillable and not fillable in Preview/runtime mode"
                >
                  ✏ Fillable
                </TB>
              </>
            )}
          </div>
        )}
      </div>

      {/* ── Grid ── */}
      <div
        ref={hotViewportRef}
        className="hot-wrapper relative z-0 overflow-hidden border rounded-md ht-theme-main"
        style={{
          width: "100%",
          height: readOnly ? (readOnlyHotHeight ?? 380) : 320,
        }}
      >
        {isHotLoading && (
          <div className="absolute inset-0 z-20 flex items-center justify-center bg-white/70">
            <div className="h-8 w-8 animate-spin rounded-full border-2 border-slate-400 border-t-transparent" />
          </div>
        )}
        <div style={hotTableScaleStyle}>
          <HotTable
            key={`ht-wb-${hotTableMountKey}`}
            ref={hotRef}
            {...hotTableSettings}
            manualColumnResize={!readOnly && !lightweightPerformance}
            manualRowResize={!readOnly && !lightweightPerformance}
            width={hotViewportWidth > 0 ? hotViewportWidth : "100%"}
          />
        </div>
      </div>

      {isPreviewTruncated && (
        <div className="px-2 py-1 text-xs text-amber-700 border border-amber-200 rounded bg-amber-50">
          Preview mode — workbooks are limited to at most {MAX_PREVIEW_ROWS}{" "}
          rows × {MAX_PREVIEW_COLS} columns; showing {previewRows} ×{" "}
          {previewCols} for stability.
        </div>
      )}
    </div>
  );
});

export default HandsontableWorkbook;