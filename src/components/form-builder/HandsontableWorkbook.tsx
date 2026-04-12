import React from "react";
import { HotTable } from "@handsontable/react";
import "handsontable/styles/handsontable.css";
import "handsontable/styles/ht-theme-main.css";
import { registerAllModules } from "handsontable/registry";
import { Button } from "../ui/button";
import { HyperFormula } from "hyperformula";
import ExcelJS from "exceljs";

registerAllModules();

type SheetData = {
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
}

const MAX_PREVIEW_ROWS = 220;
const MAX_PREVIEW_COLS = 80;
const FORMULAS_CONFIG = { engine: HyperFormula };

// ─── helpers ──────────────────────────────────────────────────────────────────

const toSafeGrid = (rawGrid: unknown): string[][] => {
  if (!Array.isArray(rawGrid) || rawGrid.length === 0) return [[""]];
  const rows = rawGrid.map((row) =>
    Array.isArray(row) ? row.map((c) => (c == null ? "" : String(c))) : [""],
  );
  return rows.length > 0 ? rows : [[""]];
};

const normalizeSheets = (input?: { sheets?: SheetData[] }): SheetData[] => {
  if (!Array.isArray(input?.sheets) || input.sheets.length === 0)
    return [{ name: "Sheet1", grid: [[""]] }];
  return input.sheets.map((sheet, i) => ({
    name: sheet?.name || `Sheet${i + 1}`,
    grid: toSafeGrid(sheet?.grid),
    mergeCells: Array.isArray(sheet?.mergeCells)
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
    cellMeta: Array.isArray(sheet?.cellMeta)
      ? sheet.cellMeta
          .filter(
            (m: any) => m && Number.isFinite(+m.row) && Number.isFinite(+m.col),
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
            source: Array.isArray(m.source) ? m.source.map(String) : undefined,
            strict: typeof m.strict === "boolean" ? m.strict : undefined,
          }))
      : [],
    images: Array.isArray((sheet as any)?.images)
      ? (sheet as any).images.filter(
          (img: any) =>
            img &&
            Number.isFinite(+img.row) &&
            Number.isFinite(+img.col) &&
            typeof img.dataUrl === "string" &&
            img.dataUrl.length > 0,
        )
      : [],
    colWidthsPx: Array.isArray(sheet?.colWidthsPx)
      ? sheet.colWidthsPx
      : undefined,
    rowHeightsPx: Array.isArray(sheet?.rowHeightsPx)
      ? sheet.rowHeightsPx
      : undefined,
    tabColor: sheet?.tabColor,
  }));
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

const HandsontableWorkbook: React.FC<HandsontableWorkbookProps> = ({
  data,
  onChange,
  readOnly = false,
}) => {
  const workbookRef = React.useRef<{ sheets: SheetData[] }>({
    sheets: normalizeSheets(data),
  });
  const lastIncomingSignatureRef = React.useRef(
    workbookSignature(normalizeSheets(data)),
  );

  const [activeSheetIndex, setActiveSheetIndex] = React.useState(0);
  const activeSheetIndexRef = React.useRef(0);
  React.useEffect(() => {
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
    if (!readOnly) return base;
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

  const textColorApplyTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(
    null,
  );
  const fillColorApplyTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(
    null,
  );

  const safeSheets = workbookRef.current.sheets;
  const activeSheet =
    safeSheets[Math.min(activeSheetIndex, safeSheets.length - 1)] ||
    safeSheets[0];
  const safeGrid =
    Array.isArray(activeSheet?.grid) && activeSheet.grid.length > 0
      ? activeSheet.grid
      : [[""]];
  const previewRows = readOnly
    ? Math.min(MAX_PREVIEW_ROWS, safeGrid.length)
    : safeGrid.length;
  const previewCols = readOnly
    ? Math.min(MAX_PREVIEW_COLS, safeGrid[0]?.length || 0)
    : safeGrid[0]?.length || 0;
  const renderedGrid = readOnly
    ? safeGrid
        .slice(0, previewRows)
        .map((row) => (Array.isArray(row) ? row.slice(0, previewCols) : []))
    : safeGrid;
  const isPreviewTruncated =
    readOnly &&
    (safeGrid.length > previewRows || (safeGrid[0]?.length || 0) > previewCols);

  const renderedMergeCells = (activeSheet.mergeCells || []).filter(
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
  const renderedColWidths = readOnly
    ? (activeSheet.colWidthsPx || []).slice(0, previewCols)
    : activeSheet.colWidthsPx;
  const renderedRowHeights = readOnly
    ? (activeSheet.rowHeightsPx || []).slice(0, previewRows)
    : activeSheet.rowHeightsPx;

  const currentCellCount = renderedGrid.reduce(
    (t, row) => t + (Array.isArray(row) ? row.length : 0),
    0,
  );
  const shouldUseFormulaEngine = !readOnly && currentCellCount <= 20000;

  const imageMap = React.useMemo(() => {
    const map = new Map<
      string,
      { dataUrl: string; rowspan: number; colspan: number }
    >();
    for (const img of (activeSheet as any)?.images || []) {
      if (!img?.dataUrl) continue;
      map.set(`${img.row}:${img.col}`, {
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
      map.set(`${meta.row}:${meta.col}`, meta);
    return map;
  }, [activeSheet]);

  // ─── selection helpers ──────────────────────────────────────────────────────

  /**
   * Snapshot selection from HOT at action time (getSelectedRangeLast first).
   * Use this in toolbar handlers instead of reading selection after focus moved.
   */
  const getToolbarActionRange = React.useCallback((hot: any) => {
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

    const cached =
      sheetSelectionRef.current[idx] ?? lastSelectionRef.current;
    return clamp(cached);
  }, [formatAllCells]);

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
    (includeMeta: boolean) => {
      const hot = hotRef.current?.hotInstance;
      if (!hot) return;

      const idx = activeSheetIndexRef.current;

      const nextGrid = (hot.getData?.() || []).map((row: any[]) =>
        row.map((cell) => (cell == null ? "" : String(cell))),
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

      let cellMeta =
        workbookRef.current.sheets[idx]?.cellMeta || [];
      if (includeMeta) {
        const nextMeta: NonNullable<SheetData["cellMeta"]> = [];
        const cellsMeta =
          typeof hot.getCellsMeta === "function" ? hot.getCellsMeta() : [];
        for (const meta of cellsMeta || []) {
          const useful =
            Boolean(meta?.className) ||
            Boolean(meta?.type) ||
            Boolean(meta?.dateFormat) ||
            typeof meta?.correctFormat === "boolean" ||
            Boolean(meta?.numericFormat) ||
            Array.isArray(meta?.source) ||
            typeof meta?.strict === "boolean";
          if (
            typeof meta?.row === "number" &&
            typeof meta?.col === "number" &&
            meta.row >= 0 &&
            meta.col >= 0 &&
            useful
          ) {
            nextMeta.push({
              row: meta.row,
              col: meta.col,
              className: meta.className ? String(meta.className) : undefined,
              type: meta.type ? String(meta.type) : undefined,
              dateFormat: meta.dateFormat ? String(meta.dateFormat) : undefined,
              correctFormat:
                typeof meta.correctFormat === "boolean"
                  ? meta.correctFormat
                  : undefined,
              numericFormat:
                meta.numericFormat && typeof meta.numericFormat === "object"
                  ? {
                      pattern:
                        typeof meta.numericFormat.pattern === "string"
                          ? meta.numericFormat.pattern
                          : undefined,
                      culture:
                        typeof meta.numericFormat.culture === "string"
                          ? meta.numericFormat.culture
                          : undefined,
                    }
                  : undefined,
              source: Array.isArray(meta.source)
                ? meta.source.map(String)
                : undefined,
              strict:
                typeof meta.strict === "boolean" ? meta.strict : undefined,
            });
          }
        }
        cellMeta = nextMeta;
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

      const colWidthsPx: number[] = [];
      for (let c = 0; c < colCount; c++) {
        const w =
          typeof hot.getColWidth === "function" ? hot.getColWidth(c) : undefined;
        const rounded =
          typeof w === "number" && Number.isFinite(w) ? Math.round(w) : NaN;
        colWidthsPx.push(
          Number.isFinite(rounded)
            ? rounded
            : Math.round(Number(current.colWidthsPx?.[c]) || 50),
        );
      }

      const rowHeightsPx: number[] = [];
      for (let r = 0; r < rowCount; r++) {
        const h =
          typeof hot.getRowHeight === "function" ? hot.getRowHeight(r) : undefined;
        const rounded =
          typeof h === "number" && Number.isFinite(h) ? Math.round(h) : NaN;
        rowHeightsPx.push(
          Number.isFinite(rounded)
            ? rounded
            : Math.round(Number(current.rowHeightsPx?.[r]) || 24),
        );
      }

      workbookRef.current.sheets[idx] = {
        ...current,
        grid: nextGrid,
        mergeCells,
        cellMeta,
        colWidthsPx,
        rowHeightsPx,
      };
    },
    [],
  );

  const toVisibleGrid = React.useCallback(
    (sheet?: SheetData) => {
      const base = sheet?.grid?.length ? sheet.grid : [[""]];
      if (!readOnly) return base;
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
        for (const meta of sheet.cellMeta || []) {
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
    if (targetIndex === activeSheetIndexRef.current) return;
    if (!readOnly) collectCurrentSheetFromHot(true);
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

  React.useEffect(() => () => flushPendingColorTimers(), [flushPendingColorTimers]);

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
    collectCurrentSheetFromHot(true);
    const range = getToolbarActionRange(hot);
    if (!range) return;

    applyClassToSelection("meta-fillable", false);

    const sheet = workbookRef.current.sheets[activeSheetIndexRef.current];
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
        metaByKey.set(key, {
          ...current,
          row: r,
          col: c,
          className: tokens.join(" ").trim() || undefined,
        });
      }
    }

    sheet.cellMeta = Array.from(metaByKey.values());
    onChange({ sheets: [...workbookRef.current.sheets] });
    if (hot && range) {
      hot.selectCell(
        range.startRow,
        range.startCol,
        range.endRow,
        range.endCol,
        false,
        false,
      );
    }
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
    onChange({
      sheets: workbookRef.current.sheets.map((s) => ({ ...s })),
    });
  };

  /** Persist column/row sizes to template and parent (e.g. before "Save Changes" without toolbar Save). */
  const flushLayoutToParent = React.useCallback(() => {
    if (readOnly) return;
    collectCurrentSheetFromHot(true);
    onChange({
      sheets: workbookRef.current.sheets.map((s) => ({ ...s })),
    });
  }, [readOnly, collectCurrentSheetFromHot, onChange]);

  const exportXlsx = async () => {
    const workbook = new ExcelJS.Workbook();
    safeSheets.forEach((sheet) => {
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
    const csv = safeGrid
      .map((row) =>
        row.map((v) => `"${String(v ?? "").replace(/"/g, '""')}"`).join(","),
      )
      .join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${activeSheet?.name || "sheet"}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  // ─── sheet management ────────────────────────────────────────────────────────

  const duplicateActiveSheet = () => {
    const cloned = JSON.parse(JSON.stringify(activeSheet)) as SheetData;
    cloned.name = `${activeSheet.name} Copy`;
    const nextSheets = [...safeSheets, cloned];
    workbookRef.current.sheets = nextSheets;
    setSheetTabs(
      nextSheets.map((s) => ({ name: s.name, tabColor: s.tabColor })),
    );
    setInitialGrid(toVisibleGrid(nextSheets[nextSheets.length - 1]));
    setActiveSheetIndex(nextSheets.length - 1);
  };

  const moveSheet = (direction: "left" | "right") => {
    const target = activeSheetIndex + (direction === "left" ? -1 : 1);
    if (target < 0 || target >= safeSheets.length) return;
    const next = [...safeSheets];
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
    const nextSheets = normalizeSheets(data);
    const sig = workbookSignature(nextSheets);
    if (sig === lastIncomingSignatureRef.current) return;
    lastIncomingSignatureRef.current = sig;

    const prevSheetCount = workbookRef.current.sheets.length;
    const nextSheetCount = nextSheets.length;
    const sheetCountChanged = prevSheetCount !== nextSheetCount;

    workbookRef.current = { sheets: nextSheets };
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
        setInitialGrid(first);
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
  }, [data, readOnly]);

  const incomingWorkbookKey = React.useMemo(
    () => workbookSignature(normalizeSheets(data)),
    [data],
  );

  React.useEffect(() => {
    loadSheetIntoHot(activeSheetIndex);
  }, [activeSheetIndex, loadSheetIntoHot, incomingWorkbookKey]);

  // ─── cell renderer ───────────────────────────────────────────────────────────

  const cellsCallback = React.useCallback(
    (row: number, col: number) => {
      const persistedMeta = persistedCellMetaMap.get(`${row}:${col}`);
      const cp: any = {};
      const persistedClassName = String(persistedMeta?.className || "");
      const classTokens = persistedClassName.split(" ").filter(Boolean);
      const isFillable = classTokens.includes("meta-fillable");
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

      const image = imageMap.get(`${row}:${col}`);
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
    (_row: number, _col: number, cellProps: Record<string, unknown>) => {
      if (!readOnly) {
        (cellProps as { readOnly?: boolean }).readOnly = false;
        return;
      }
      const cls = String((cellProps as { className?: string }).className || "");
      const isFillable = cls
        .split(" ")
        .filter(Boolean)
        .includes("meta-fillable");
      (cellProps as { readOnly?: boolean }).readOnly = !isFillable;
    },
    [readOnly],
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
    <div className="space-y-2">
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
            placeholder="Enter value or formula (e.g. =SUM(A1:A5))"
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

      {/* ── Sheet tabs ── */}
      <div className="relative z-10 flex flex-wrap items-center gap-2">
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

        {!readOnly && (
          <>
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
          </>
        )}
      </div>

      {/* ── Grid ── */}
      <div className="relative z-0 overflow-hidden border rounded-md">
        <HotTable
          ref={hotRef}
          data={initialGrid}
          themeName="ht-theme-main"
          rowHeaders
          colHeaders
          licenseKey="non-commercial-and-evaluation"
          readOnly={false}
          trimWhitespace={false}
          width="100%"
          stretchH={readOnly ? "all" : "none"}
          height={320}
          formulas={shouldUseFormulaEngine ? FORMULAS_CONFIG : undefined}
          mergeCells={
            renderedMergeCells.length > 0 ? renderedMergeCells : !readOnly
          }
          filters={!readOnly}
          dropdownMenu={!readOnly}
          columnSorting={!readOnly}
          hiddenRows={{ indicators: true }}
          hiddenColumns={{ indicators: true }}
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
          afterChange={() => refreshUndoRedoState()}
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
            setSelectionLabel(toRangeLabel(range));
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
};

export default HandsontableWorkbook;
