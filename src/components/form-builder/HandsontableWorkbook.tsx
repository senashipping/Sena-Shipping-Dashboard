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
  mergeCells?: Array<{ row: number; col: number; rowspan: number; colspan: number }>;
  cellMeta?: Array<{ row: number; col: number; className?: string }>;
  images?: Array<{ row: number; col: number; rowspan?: number; colspan?: number; dataUrl: string }>;
  colWidthsPx?: number[];
  rowHeightsPx?: number[];
  tabColor?: string;
};

interface HandsontableWorkbookProps {
  data: { sheets: SheetData[] };
  onChange: (next: { sheets: SheetData[] }) => void;
  readOnly?: boolean;
}

const MAX_PREVIEW_ROWS = 220;
const MAX_PREVIEW_COLS = 80;
const FORMULAS_CONFIG = { engine: HyperFormula };

const toSafeGrid = (rawGrid: unknown): string[][] => {
  if (!Array.isArray(rawGrid) || rawGrid.length === 0) return [[""]];
  const safeRows = rawGrid.map((row) => {
    if (!Array.isArray(row)) return [""];
    return row.map((cell) => (cell == null ? "" : String(cell)));
  });
  return safeRows.length > 0 ? safeRows : [[""]];
};

const normalizeSheets = (input?: { sheets?: SheetData[] }): SheetData[] => {
  if (!Array.isArray(input?.sheets) || input.sheets.length === 0) {
    return [{ name: "Sheet1", grid: [[""]] }];
  }
  return input.sheets.map((sheet, index) => ({
    name: sheet?.name || `Sheet${index + 1}`,
    grid: toSafeGrid(sheet?.grid),
    mergeCells: Array.isArray(sheet?.mergeCells)
      ? sheet.mergeCells
          .filter(
            (m: any) =>
              m &&
              Number.isFinite(Number(m.row)) &&
              Number.isFinite(Number(m.col)) &&
              Number.isFinite(Number(m.rowspan)) &&
              Number.isFinite(Number(m.colspan)) &&
              Number(m.rowspan) > 0 &&
              Number(m.colspan) > 0
          )
          .map((m: any) => ({
            row: Number(m.row),
            col: Number(m.col),
            rowspan: Number(m.rowspan),
            colspan: Number(m.colspan),
          }))
      : [],
    cellMeta: Array.isArray(sheet?.cellMeta)
      ? sheet.cellMeta
          .filter(
            (m: any) =>
              m &&
              Number.isFinite(Number(m.row)) &&
              Number.isFinite(Number(m.col))
          )
          .map((m: any) => ({ row: Number(m.row), col: Number(m.col), className: m.className }))
      : [],
    images: Array.isArray((sheet as any)?.images)
      ? (sheet as any).images.filter(
          (img: any) =>
            img &&
            Number.isFinite(Number(img.row)) &&
            Number.isFinite(Number(img.col)) &&
            typeof img.dataUrl === "string" &&
            img.dataUrl.length > 0
        )
      : [],
    colWidthsPx: Array.isArray(sheet?.colWidthsPx) ? sheet.colWidthsPx : undefined,
    rowHeightsPx: Array.isArray(sheet?.rowHeightsPx) ? sheet.rowHeightsPx : undefined,
    tabColor: sheet?.tabColor,
  }));
};
const workbookSignature = (sheets: SheetData[]): string =>
  sheets
    .map((s) => {
      const rows = s.grid?.length || 0;
      const cols = s.grid?.[0]?.length || 0;
      const merges = s.mergeCells?.length || 0;
      const meta = s.cellMeta?.length || 0;
      return `${s.name}|${rows}x${cols}|m${merges}|c${meta}|${s.tabColor || ""}`;
    })
    .join("::");



const HandsontableWorkbook: React.FC<HandsontableWorkbookProps> = ({
  data,
  onChange,
  readOnly = false,
}) => {
  const initialSheets = React.useMemo(() => normalizeSheets(data), []);
  const workbookRef = React.useRef<{ sheets: SheetData[] }>({ sheets: initialSheets });
  const lastIncomingSignatureRef = React.useRef<string>(workbookSignature(initialSheets));
  const [activeSheetIndex, setActiveSheetIndex] = React.useState(0);
  const [sheetTabs, setSheetTabs] = React.useState<Array<{ name: string; tabColor?: string }>>(
    workbookRef.current.sheets.map((s) => ({ name: s.name, tabColor: s.tabColor }))
  );
  const [initialGrid, setInitialGrid] = React.useState<string[][]>(() => {
    const first = workbookRef.current.sheets[0];
    const base = Array.isArray(first?.grid) && first.grid.length > 0 ? first.grid : [[""]];
    if (!readOnly) return base;
    const rows = Math.min(MAX_PREVIEW_ROWS, base.length);
    const cols = Math.min(MAX_PREVIEW_COLS, base[0]?.length || 0);
    return base.slice(0, rows).map((row) => (Array.isArray(row) ? row.slice(0, cols) : []));
  });
  const [renaming, setRenaming] = React.useState(false);
  const [renameValue, setRenameValue] = React.useState("");
  const [formulaInput, setFormulaInput] = React.useState("");
  const [findValue, setFindValue] = React.useState("");
  const [replaceValue, setReplaceValue] = React.useState("");
  const [dropdownSource, setDropdownSource] = React.useState("Option A,Option B");
  const [fontFamily, setFontFamily] = React.useState("Arial");
  const [fontSize, setFontSize] = React.useState("12");
  const [textColor, setTextColor] = React.useState("#111827");
  const [fillColor, setFillColor] = React.useState("#ffffff");
  const [isBoldActive, setIsBoldActive] = React.useState(false);
  const [isItalicActive, setIsItalicActive] = React.useState(false);
  const [isUnderlineActive, setIsUnderlineActive] = React.useState(false);
  const [isStrikeActive, setIsStrikeActive] = React.useState(false);
  const [selectedAlign, setSelectedAlign] = React.useState<"left" | "center" | "right" | "justify" | null>(null);
  const [selectedVAlign, setSelectedVAlign] = React.useState<"top" | "middle" | "bottom" | null>(null);
  const [formatAllCells, setFormatAllCells] = React.useState(false);
  const lastSelectionRef = React.useRef<{ startRow: number; endRow: number; startCol: number; endCol: number } | null>(null);
  const sheetSelectionRef = React.useRef<Record<number, { startRow: number; endRow: number; startCol: number; endCol: number }>>({});
  const [fixedRowsTop, setFixedRowsTop] = React.useState(0);
  const [fixedColumnsStart, setFixedColumnsStart] = React.useState(0);
  const hotRef = React.useRef<any>(null);
  const safeSheets = workbookRef.current.sheets;
  const activeSheet = safeSheets[Math.min(activeSheetIndex, safeSheets.length - 1)] || safeSheets[0];
  const safeGrid =
    Array.isArray(activeSheet?.grid) && activeSheet.grid.length > 0 ? activeSheet.grid : [[""]];
  const previewRows = readOnly ? Math.min(MAX_PREVIEW_ROWS, safeGrid.length) : safeGrid.length;
  const previewCols = readOnly ? Math.min(MAX_PREVIEW_COLS, safeGrid[0]?.length || 0) : (safeGrid[0]?.length || 0);
  const renderedGrid = readOnly
    ? safeGrid.slice(0, previewRows).map((row) => (Array.isArray(row) ? row.slice(0, previewCols) : []))
    : safeGrid;
  const isPreviewTruncated =
    readOnly && (safeGrid.length > previewRows || (safeGrid[0]?.length || 0) > previewCols);
  const renderedMergeCells = (activeSheet.mergeCells || []).filter(
    (m) =>
      m &&
      Number.isFinite(Number(m.row)) &&
      Number.isFinite(Number(m.col)) &&
      Number.isFinite(Number(m.rowspan)) &&
      Number.isFinite(Number(m.colspan)) &&
      m.row < previewRows &&
      m.col < previewCols &&
      m.row + m.rowspan <= previewRows &&
      m.col + m.colspan <= previewCols
  );
  const renderedColWidths = readOnly ? (activeSheet.colWidthsPx || []).slice(0, previewCols) : activeSheet.colWidthsPx;
  const renderedRowHeights = readOnly ? (activeSheet.rowHeightsPx || []).slice(0, previewRows) : activeSheet.rowHeightsPx;
  const currentCellCount = renderedGrid.reduce(
    (total, row) => total + (Array.isArray(row) ? row.length : 0),
    0
  );
  const shouldUseFormulaEngine = !readOnly && currentCellCount <= 20000;
  const imageMap = React.useMemo(() => {
    const map = new Map<string, { dataUrl: string; rowspan: number; colspan: number }>();
    for (const img of (activeSheet as any)?.images || []) {
      if (!img?.dataUrl) continue;
      map.set(`${img.row}:${img.col}`, {
        dataUrl: img.dataUrl,
        rowspan: Math.max(1, Number(img.rowspan) || 1),
        colspan: Math.max(1, Number(img.colspan) || 1),
      });
    }
    return map;
  }, [activeSheet]);
  const shouldApplyCellRenderer = !readOnly || imageMap.size > 0;
  const persistedCellMetaMap = React.useMemo(() => {
    const map = new Map<string, { className?: string }>();
    for (const meta of activeSheet?.cellMeta || []) {
      map.set(`${meta.row}:${meta.col}`, {
        className: meta.className,
      });
    }
    return map;
  }, [activeSheet]);

  const collectCurrentSheetFromHot = (includeMeta: boolean) => {
    const hot = hotRef.current?.hotInstance;
    if (!hot) return;
    const nextGrid = (hot.getData?.() || []).map((row: any[]) =>
      row.map((cell) => (cell == null ? "" : String(cell)))
    );
    const mergeCells =
      hot?.getPlugin?.("mergeCells")?.mergedCellsCollection?.mergedCells?.map((cell: any) => ({
        row: cell.row,
        col: cell.col,
        rowspan: cell.rowspan,
        colspan: cell.colspan,
      })) || [];
    let cellMeta = workbookRef.current.sheets[activeSheetIndex]?.cellMeta || [];
    if (includeMeta) {
      // Use Handsontable tracked meta list instead of scanning all cells.
      const nextMeta: Array<{ row: number; col: number; className?: string }> = [];
      const cellsMeta = typeof hot.getCellsMeta === "function" ? hot.getCellsMeta() : [];
      for (const meta of cellsMeta || []) {
        if (
          typeof meta?.row === "number" &&
          typeof meta?.col === "number" &&
          meta.row >= 0 &&
          meta.col >= 0 &&
          meta.className
        ) {
          nextMeta.push({
            row: meta.row,
            col: meta.col,
            className: String(meta.className),
          });
        }
      }
      cellMeta = nextMeta;
    }
    const current = workbookRef.current.sheets[activeSheetIndex] || { name: `Sheet${activeSheetIndex + 1}`, grid: [[""]] };
    workbookRef.current.sheets[activeSheetIndex] = {
      ...current,
      grid: nextGrid,
      mergeCells,
      cellMeta,
      colWidthsPx: current.colWidthsPx,
      rowHeightsPx: current.rowHeightsPx,
    };
  };

  const getSelectedRange = () => {
    const hot = hotRef.current?.hotInstance;
    if (!hot) return null;
    if (formatAllCells) {
      const sheet = workbookRef.current.sheets[activeSheetIndex];
      const rows = Math.max(1, sheet?.grid?.length || 1);
      const cols = Math.max(1, sheet?.grid?.[0]?.length || 1);
      return { startRow: 0, endRow: rows - 1, startCol: 0, endCol: cols - 1 };
    }
    if (lastSelectionRef.current) {
      return lastSelectionRef.current;
    }
    const selected = hot.getSelectedLast?.();
    if (!selected) return null;
    const [r1, c1, r2, c2] = selected;
    return {
      startRow: Math.min(r1, r2),
      endRow: Math.max(r1, r2),
      startCol: Math.min(c1, c2),
      endCol: Math.max(c1, c2),
    };
  };

  const ensureSelection = () => {
    const hot = hotRef.current?.hotInstance;
    const range = getSelectedRange();
    if (!hot || !range) return null;
    hot.selectCell(range.startRow, range.startCol, range.endRow, range.endCol, false, false);
    return range;
  };

  const loadSheetIntoHot = React.useCallback((targetIndex: number) => {
    const hot = hotRef.current?.hotInstance;
    if (!hot) return;
    const sheet = workbookRef.current.sheets[targetIndex];
    if (!sheet) return;
    const baseGrid = sheet.grid?.length ? sheet.grid : [[""]];
    const rows = readOnly ? Math.min(MAX_PREVIEW_ROWS, baseGrid.length) : baseGrid.length;
    const cols = readOnly ? Math.min(MAX_PREVIEW_COLS, baseGrid[0]?.length || 0) : (baseGrid[0]?.length || 0);
    const visibleGrid = readOnly
      ? baseGrid.slice(0, rows).map((row) => (Array.isArray(row) ? row.slice(0, cols) : []))
      : baseGrid;
    setInitialGrid(visibleGrid);
    hot.loadData(visibleGrid);
    if (!readOnly) {
      for (const meta of sheet.cellMeta || []) {
        if (meta.className) hot.setCellMeta(meta.row, meta.col, "className", meta.className);
      }
    }
    hot.render();
  }, [readOnly]);

  const handleSheetSwitch = (targetIndex: number) => {
    if (targetIndex === activeSheetIndex) return;
    if (!readOnly) {
      collectCurrentSheetFromHot(true);
    }
    lastSelectionRef.current = sheetSelectionRef.current[targetIndex] || null;
    setActiveSheetIndex(targetIndex);
    setTimeout(() => loadSheetIntoHot(targetIndex), 0);
  };

  const emitWorkbookSnapshot = () => {
    if (!readOnly) collectCurrentSheetFromHot(true);
    onChange({ sheets: workbookRef.current.sheets });
  };

  React.useEffect(() => {
    const nextSheets = normalizeSheets(data);
    const incomingSig = workbookSignature(nextSheets);
    if (incomingSig === lastIncomingSignatureRef.current) return;
    lastIncomingSignatureRef.current = incomingSig;
    workbookRef.current = { sheets: nextSheets };
    setSheetTabs(nextSheets.map((s) => ({ name: s.name, tabColor: s.tabColor })));
    setActiveSheetIndex(0);
    const first = nextSheets[0]?.grid?.length ? nextSheets[0].grid : [[""]];
    if (!readOnly) {
      setInitialGrid(first);
    } else {
      const rows = Math.min(MAX_PREVIEW_ROWS, first.length);
      const cols = Math.min(MAX_PREVIEW_COLS, first[0]?.length || 0);
      setInitialGrid(first.slice(0, rows).map((row) => (Array.isArray(row) ? row.slice(0, cols) : [])));
    }
  }, [data, readOnly]);

  const applyClassToSelection = (classToken: string, toggle = false, replacePrefix?: string) => {
    const hot = hotRef.current?.hotInstance;
    const range = ensureSelection();
    if (!hot || !range || readOnly) return;
    const tokenPrefix = replacePrefix || classToken;
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
            : [...current.filter((x: string) => !x.startsWith(tokenPrefix)), classToken];
          hot.setCellMeta(r, c, "className", next.join(" ").trim());
        }
      }
    };
    if (typeof hot.batch === "function") hot.batch(apply);
    else apply();
    hot.render();
    collectCurrentSheetFromHot(true);
  };

  const setAlignment = (align: "left" | "center" | "right" | "justify") => {
    applyClassToSelection(`meta-align-${align}`, false, "meta-align-");
    setSelectedAlign(align);
  };

  const setVerticalAlignment = (align: "top" | "middle" | "bottom") => {
    applyClassToSelection(`meta-valign-${align}`, false, "meta-valign-");
    setSelectedVAlign(align);
  };

  const setWrapText = () => {
    applyClassToSelection("meta-wrap", true);
  };

  const setFontStyle = (style: "bold" | "italic" | "underline" | "strike") => {
    applyClassToSelection(`meta-${style}`, true);
    if (style === "bold") setIsBoldActive((prev) => !prev);
    if (style === "italic") setIsItalicActive((prev) => !prev);
    if (style === "underline") setIsUnderlineActive((prev) => !prev);
    if (style === "strike") setIsStrikeActive((prev) => !prev);
  };

  const applyFontFamily = () => applyClassToSelection(`meta-font-${fontFamily.replace(/\s+/g, "_")}`, false, "meta-font-");
  const applyFontSize = () => applyClassToSelection(`meta-size-${fontSize}`, false, "meta-size-");
  const applyTextColor = () => applyClassToSelection(`meta-color-${textColor.replace("#", "")}`, false, "meta-color-");
  const applyFillColor = () => applyClassToSelection(`meta-fill-${fillColor.replace("#", "")}`, false, "meta-fill-");

  const formatSelectedAs = (kind: "number" | "currency" | "percent" | "date") => {
    const hot = hotRef.current?.hotInstance;
    const range = ensureSelection();
    if (!hot || !range || readOnly) return;
    const apply = () => {
      for (let r = range.startRow; r <= range.endRow; r++) {
        for (let c = range.startCol; c <= range.endCol; c++) {
          const raw = hot.getDataAtCell(r, c);
          if (kind === "date") {
            hot.setCellMeta(r, c, "type", "date");
            hot.setCellMeta(r, c, "dateFormat", "YYYY-MM-DD");
            hot.setCellMeta(r, c, "correctFormat", true);
            if (raw == null || raw === "") continue;
            const d = new Date(raw);
            if (!Number.isNaN(d.getTime())) hot.setDataAtCell(r, c, d.toISOString().slice(0, 10));
            continue;
          }

          hot.setCellMeta(r, c, "type", "numeric");
          if (kind === "number") {
            hot.setCellMeta(r, c, "numericFormat", { pattern: "0,0.00", culture: "en-US" });
          }
          if (kind === "currency") {
            hot.setCellMeta(r, c, "numericFormat", { pattern: "$0,0.00", culture: "en-US" });
          }
          if (kind === "percent") {
            hot.setCellMeta(r, c, "numericFormat", { pattern: "0.00%", culture: "en-US" });
          }

          if (raw == null || raw === "") continue;
          const rawText = String(raw).trim();
          let numeric = Number(rawText.replace(/[$,%\s,]/g, ""));
          if (kind === "percent" && rawText.includes("%")) {
            numeric = numeric / 100;
          }
          if (!Number.isNaN(numeric)) {
            hot.setDataAtCell(r, c, numeric);
          }
        }
      }
    };
    if (typeof hot.batch === "function") hot.batch(apply);
    else apply();
    hot.render();
    collectCurrentSheetFromHot(true);
  };

  const applyDropdownValidation = () => {
    const hot = hotRef.current?.hotInstance;
    const range = ensureSelection();
    if (!hot || !range || readOnly) return;
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
    hot.render();
    collectCurrentSheetFromHot(true);
  };

  const applyDateCellType = () => {
    const hot = hotRef.current?.hotInstance;
    const range = ensureSelection();
    if (!hot || !range || readOnly) return;
    for (let r = range.startRow; r <= range.endRow; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        hot.setCellMeta(r, c, "type", "date");
        hot.setCellMeta(r, c, "dateFormat", "YYYY-MM-DD");
        hot.setCellMeta(r, c, "correctFormat", true);
      }
    }
    hot.render();
    collectCurrentSheetFromHot(true);
  };

  const sortSelectedColumn = (order: "asc" | "desc") => {
    const hot = hotRef.current?.hotInstance;
    const range = ensureSelection();
    if (!hot || !range) return;
    const sortingPlugin = hot.getPlugin?.("columnSorting");
    if (!sortingPlugin || typeof sortingPlugin.sort !== "function") return;
    sortingPlugin.sort({ column: range.startCol, sortOrder: order });
  };

  const doFindReplace = () => {
    const hot = hotRef.current?.hotInstance;
    if (!hot || !findValue || readOnly) return;
    const data = hot.getData();
    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        const val = String(data[r][c] ?? "");
        if (val.includes(findValue)) {
          hot.setDataAtCell(r, c, val.split(findValue).join(replaceValue));
        }
      }
    }
    collectCurrentSheetFromHot(false);
  };

  const exportXlsx = async () => {
    const workbook = new ExcelJS.Workbook();
    safeSheets.forEach((sheet) => {
      const ws = workbook.addWorksheet(sheet.name || "Sheet");
      sheet.grid.forEach((row) => ws.addRow(row));
    });
    const buf = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "workbook.xlsx";
    a.click();
    URL.revokeObjectURL(url);
  };

  const exportCsv = () => {
    const csv = safeGrid.map((row) => row.map((v) => `"${String(v ?? "").split('"').join('""')}"`).join(",")).join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${activeSheet?.name || "sheet"}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const duplicateActiveSheet = () => {
    const cloned = JSON.parse(JSON.stringify(activeSheet)) as SheetData;
    cloned.name = `${activeSheet.name} Copy`;
    const nextSheets = [...safeSheets, cloned];
    workbookRef.current.sheets = nextSheets;
    setSheetTabs(nextSheets.map((s) => ({ name: s.name, tabColor: s.tabColor })));
    setActiveSheetIndex(nextSheets.length - 1);
    setTimeout(() => loadSheetIntoHot(nextSheets.length - 1), 0);
  };

  const moveSheet = (direction: "left" | "right") => {
    const idx = activeSheetIndex;
    const target = direction === "left" ? idx - 1 : idx + 1;
    if (target < 0 || target >= safeSheets.length) return;
    const next = [...safeSheets];
    const [moved] = next.splice(idx, 1);
    next.splice(target, 0, moved);
    workbookRef.current.sheets = next;
    setSheetTabs(next.map((s) => ({ name: s.name, tabColor: s.tabColor })));
    setActiveSheetIndex(target);
    setTimeout(() => loadSheetIntoHot(target), 0);
  };

  const applySheetColor = (color: string) => {
    const nextSheets = safeSheets.map((s, idx) =>
      idx === activeSheetIndex ? { ...s, tabColor: color } : s
    );
    workbookRef.current.sheets = nextSheets;
    setSheetTabs(nextSheets.map((s) => ({ name: s.name, tabColor: s.tabColor })));
  };

  const applyFormulaBar = () => {
    const hot = hotRef.current?.hotInstance;
    const range = ensureSelection();
    if (!hot || !range || readOnly) return;
    hot.setDataAtCell(range.startRow, range.startCol, formulaInput);
    collectCurrentSheetFromHot(false);
  };

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
        [class*="meta-font-"] { font-family: Arial, sans-serif; }
      `}</style>
      {!readOnly && (
        <div className="flex flex-wrap items-center gap-1 p-2 border rounded-md bg-slate-50">
          <Button type="button" size="sm" variant={isBoldActive ? "default" : "outline"} onClick={() => setFontStyle("bold")}>B</Button>
          <Button type="button" size="sm" variant={isItalicActive ? "default" : "outline"} onClick={() => setFontStyle("italic")}>I</Button>
          <Button type="button" size="sm" variant={isUnderlineActive ? "default" : "outline"} onClick={() => setFontStyle("underline")}>U</Button>
          <Button type="button" size="sm" variant={isStrikeActive ? "default" : "outline"} onClick={() => setFontStyle("strike")}>S</Button>
          <select value={fontFamily} onChange={(e) => setFontFamily(e.target.value)} className="h-8 px-2 text-sm border rounded">
            <option>Arial</option><option>Calibri</option><option>Times New Roman</option><option>Verdana</option>
          </select>
          <input value={fontSize} onChange={(e) => setFontSize(e.target.value)} className="w-14 h-8 px-2 text-sm border rounded" />
          <Button type="button" size="sm" variant="outline" onClick={applyFontFamily}>Font</Button>
          <Button type="button" size="sm" variant="outline" onClick={applyFontSize}>Size</Button>
          <input type="color" value={textColor} onChange={(e) => setTextColor(e.target.value)} className="w-8 h-8 p-0 border rounded" />
          <Button type="button" size="sm" variant="outline" onClick={applyTextColor}>Text</Button>
          <input type="color" value={fillColor} onChange={(e) => setFillColor(e.target.value)} className="w-8 h-8 p-0 border rounded" />
          <Button type="button" size="sm" variant="outline" onClick={applyFillColor}>Fill</Button>
          <Button
            type="button"
            size="sm"
            variant={selectedAlign === "left" ? "default" : "outline"}
            onClick={() => setAlignment("left")}
          >
            Left
          </Button>
          <Button
            type="button"
            size="sm"
            variant={selectedAlign === "center" ? "default" : "outline"}
            onClick={() => setAlignment("center")}
          >
            Center
          </Button>
          <Button
            type="button"
            size="sm"
            variant={selectedAlign === "right" ? "default" : "outline"}
            onClick={() => setAlignment("right")}
          >
            Right
          </Button>
          <Button
            type="button"
            size="sm"
            variant={selectedAlign === "justify" ? "default" : "outline"}
            onClick={() => setAlignment("justify")}
          >
            Justify
          </Button>
          <Button
            type="button"
            size="sm"
            variant={selectedVAlign === "top" ? "default" : "outline"}
            onClick={() => setVerticalAlignment("top")}
          >
            Top
          </Button>
          <Button
            type="button"
            size="sm"
            variant={selectedVAlign === "middle" ? "default" : "outline"}
            onClick={() => setVerticalAlignment("middle")}
          >
            Middle
          </Button>
          <Button
            type="button"
            size="sm"
            variant={selectedVAlign === "bottom" ? "default" : "outline"}
            onClick={() => setVerticalAlignment("bottom")}
          >
            Bottom
          </Button>
          <Button type="button" size="sm" variant="outline" onClick={setWrapText}>Wrap</Button>
          <Button type="button" size="sm" variant="outline" onClick={() => formatSelectedAs("number")}>123</Button>
          <Button type="button" size="sm" variant="outline" onClick={() => formatSelectedAs("currency")}>$</Button>
          <Button type="button" size="sm" variant="outline" onClick={() => formatSelectedAs("percent")}>%</Button>
          <Button type="button" size="sm" variant="outline" onClick={() => formatSelectedAs("date")}>Date</Button>
          <Button type="button" size="sm" variant="outline" onClick={() => sortSelectedColumn("asc")}>A→Z</Button>
          <Button type="button" size="sm" variant="outline" onClick={() => sortSelectedColumn("desc")}>Z→A</Button>
          <input value={findValue} onChange={(e) => setFindValue(e.target.value)} placeholder="Find" className="h-8 px-2 text-sm border rounded" />
          <input value={replaceValue} onChange={(e) => setReplaceValue(e.target.value)} placeholder="Replace" className="h-8 px-2 text-sm border rounded" />
          <Button type="button" size="sm" variant="outline" onClick={doFindReplace}>Replace</Button>
          <Button
            type="button"
            size="sm"
            variant={formatAllCells ? "default" : "outline"}
            onClick={() => setFormatAllCells((prev) => !prev)}
            title="Apply toolbar formatting to all cells in active sheet"
          >
            All Cells
          </Button>
          <Button type="button" size="sm" variant="default" onClick={emitWorkbookSnapshot}>Save Workbook</Button>
          <Button type="button" size="sm" variant="outline" onClick={exportXlsx}>Export .xlsx</Button>
          <Button type="button" size="sm" variant="outline" onClick={exportCsv}>CSV</Button>
        </div>
      )}

      <div className="flex items-center gap-2 p-2 border rounded-md">
        <span className="text-xs text-gray-500">Formula</span>
        <input
          value={formulaInput}
          onChange={(e) => setFormulaInput(e.target.value)}
          className="flex-1 h-8 px-2 text-sm border rounded"
          placeholder="Type formula/value for active cell (e.g. =SUM(A1:A5))"
        />
        {!readOnly && (
          <Button type="button" size="sm" variant="outline" onClick={applyFormulaBar}>
            Apply
          </Button>
        )}
      </div>

      <div className="flex flex-wrap items-center gap-2">
        {sheetTabs.map((sheet, index) => (
          <Button
            key={`${sheet.name}-${index}`}
            type="button"
            variant={index === activeSheetIndex ? "default" : "outline"}
            size="sm"
            style={sheet.tabColor ? { backgroundColor: sheet.tabColor, color: "#111827" } : undefined}
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
                className="h-8 px-2 text-sm border rounded"
                placeholder="Sheet name"
              />
              <Button
                type="button"
                variant="outline"
                size="sm"
                onClick={() => {
                  const nextName = renameValue.trim();
                  if (!nextName) return;
                  const nextSheets = safeSheets.map((sheet, index) =>
                    index === activeSheetIndex ? { ...sheet, name: nextName } : sheet
                  );
                  workbookRef.current.sheets = nextSheets;
                  setSheetTabs(nextSheets.map((s) => ({ name: s.name, tabColor: s.tabColor })));
                  setRenaming(false);
                }}
              >
                Save
              </Button>
              <Button
                type="button"
                variant="outline"
                size="sm"
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
              className="text-yellow-900 border-yellow-500 bg-yellow-300 hover:bg-yellow-400 hover:border-yellow-600"
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
            className="text-green-900 border-green-500 bg-green-300 hover:bg-green-400 hover:border-green-600"
            onClick={() => {
              const nextSheets = [
                ...safeSheets,
                { name: `Sheet${safeSheets.length + 1}`, grid: [[""]] },
              ];
              workbookRef.current.sheets = nextSheets;
              setSheetTabs(nextSheets.map((s) => ({ name: s.name, tabColor: s.tabColor })));
              setActiveSheetIndex(nextSheets.length - 1);
              setTimeout(() => loadSheetIntoHot(nextSheets.length - 1), 0);
            }}
          >
            + Add Sheet
          </Button>
          <Button type="button" variant="outline" size="sm" onClick={duplicateActiveSheet}>
            Duplicate
          </Button>
          <Button type="button" variant="outline" size="sm" onClick={() => moveSheet("left")}>
            Move Left
          </Button>
          <Button type="button" variant="outline" size="sm" onClick={() => moveSheet("right")}>
            Move Right
          </Button>
          <input type="color" className="w-8 h-8 p-0 border rounded" onChange={(e) => applySheetColor(e.target.value)} />
          <input
            value={dropdownSource}
            onChange={(e) => setDropdownSource(e.target.value)}
            className="h-8 px-2 text-sm border rounded"
            placeholder="Dropdown: A,B,C"
          />
          <Button type="button" variant="outline" size="sm" onClick={applyDropdownValidation}>
            Set Dropdown
          </Button>
          <Button type="button" variant="outline" size="sm" onClick={applyDateCellType}>
            Set Date Cell
          </Button>
          <span className="text-xs text-gray-500">Freeze</span>
          <input
            value={fixedRowsTop}
            onChange={(e) => setFixedRowsTop(Math.max(0, Number(e.target.value) || 0))}
            className="w-12 h-8 px-2 text-sm border rounded"
            title="Rows"
          />
          <input
            value={fixedColumnsStart}
            onChange={(e) => setFixedColumnsStart(Math.max(0, Number(e.target.value) || 0))}
            className="w-12 h-8 px-2 text-sm border rounded"
            title="Columns"
          />
          </>
        )}
      </div>
      <div className="overflow-hidden border rounded-md">
      <HotTable
        ref={hotRef}
        data={initialGrid}
        themeName="ht-theme-main"
        rowHeaders
        colHeaders
        licenseKey="non-commercial-and-evaluation"
        readOnly={readOnly}
        width="100%"
        stretchH="all"
        height={320}
        formulas={shouldUseFormulaEngine ? FORMULAS_CONFIG : undefined}
        mergeCells={readOnly ? renderedMergeCells : (renderedMergeCells.length > 0 ? renderedMergeCells : true)}
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
        cells={
          shouldApplyCellRenderer
            ? (row, col) => {
                const persistedMeta = persistedCellMetaMap.get(`${row}:${col}`);
                const className = String(persistedMeta?.className || "");
                const cp: any = {};
                // Force editable cells in non-readonly mode.
                if (!readOnly) {
                  cp.readOnly = false;
                }
                if (className) cp.className = className;
                const image = imageMap.get(`${row}:${col}`);
                if (image || !readOnly) {
                  cp.renderer = (
                    instance: any,
                    td: HTMLTableCellElement,
                    rowIndex: number,
                    colIndex: number,
                    prop: any,
                    value: any,
                    cellProperties: any
                  ) => {
                    const base = (window as any).Handsontable?.renderers?.TextRenderer;
                    if (base) base(instance, td, rowIndex, colIndex, prop, value, cellProperties);
                    else td.textContent = value == null ? "" : String(value);
                    const tokenClassName = String(cellProperties?.className || "");
                    const tokens = tokenClassName.split(" ").filter(Boolean);
                    const isBold = tokens.includes("meta-bold");
                    const isItalic = tokens.includes("meta-italic");
                    const isUnderline = tokens.includes("meta-underline");
                    const isStrike = tokens.includes("meta-strike");
                    const fontToken = tokens.find((t: string) => t.startsWith("meta-font-"));
                    const sizeToken = tokens.find((t: string) => t.startsWith("meta-size-"));
                    const colorToken = tokens.find((t: string) => t.startsWith("meta-color-"));
                    const fillToken = tokens.find((t: string) => t.startsWith("meta-fill-"));
                    if (isBold) td.style.fontWeight = "700";
                    if (isItalic) td.style.fontStyle = "italic";
                    if (isUnderline || isStrike) {
                      const decorations = [
                        isUnderline ? "underline" : "",
                        isStrike ? "line-through" : "",
                      ].filter(Boolean).join(" ");
                      td.style.textDecoration = decorations;
                    }
                    if (fontToken) td.style.fontFamily = fontToken.replace("meta-font-", "").split("_").join(" ");
                    if (sizeToken) td.style.fontSize = `${sizeToken.replace("meta-size-", "")}px`;
                    if (colorToken) td.style.color = `#${colorToken.replace("meta-color-", "")}`;
                    if (fillToken) td.style.backgroundColor = `#${fillToken.replace("meta-fill-", "")}`;
                    if (image?.dataUrl) {
                      const colWidths = Array.isArray(renderedColWidths) ? renderedColWidths : [];
                      const rowHeights = Array.isArray(renderedRowHeights) ? renderedRowHeights : [];
                      let imageWidthPx = 0;
                      for (let cx = colIndex; cx < colIndex + image.colspan; cx++) {
                        imageWidthPx += Number(colWidths[cx] || 80);
                      }
                      let imageHeightPx = 0;
                      for (let rx = rowIndex; rx < rowIndex + image.rowspan; rx++) {
                        imageHeightPx += Number(rowHeights[rx] || 24);
                      }
                      td.style.padding = "0";
                      td.style.position = "relative";
                      td.style.overflow = "visible";
                      td.textContent = "";
                      const img = document.createElement("img");
                      img.src = image.dataUrl;
                      img.style.position = "absolute";
                      img.style.left = "0";
                      img.style.top = "0";
                      img.style.width = `${Math.max(16, imageWidthPx)}px`;
                      img.style.height = `${Math.max(16, imageHeightPx)}px`;
                      img.style.objectFit = "fill";
                      img.style.display = "block";
                      img.style.pointerEvents = "none";
                      img.style.zIndex = "3";
                      td.appendChild(img);
                    }
                    return td;
                  };
                }
                return cp;
              }
            : undefined
        }
        afterChange={() => {
          // Keep Handsontable fully in charge during editing.
          // We only sync to ref on explicit save or sheet switch.
        }}
        afterSelectionEnd={(r, c) => {
          const hot = hotRef.current?.hotInstance;
          if (!hot) return;
          const hasValidCell = Number.isInteger(r) && Number.isInteger(c) && r >= 0 && c >= 0;
          if (!hasValidCell) {
            // Keep last valid selection so toolbar actions still apply
            // after focus moves from the grid to toolbar controls.
            return;
          }
          const selected = hot.getSelectedLast?.();
          if (selected) {
            const [r1, c1, r2, c2] = selected;
            lastSelectionRef.current = {
              startRow: Math.min(r1, r2),
              endRow: Math.max(r1, r2),
              startCol: Math.min(c1, c2),
              endCol: Math.max(c1, c2),
            };
            sheetSelectionRef.current[activeSheetIndex] = lastSelectionRef.current;
          } else {
            lastSelectionRef.current = { startRow: r, endRow: r, startCol: c, endCol: c };
            sheetSelectionRef.current[activeSheetIndex] = lastSelectionRef.current;
          }
          const v = hot.getDataAtCell(r, c);
          setFormulaInput(v == null ? "" : String(v));
          const cls = String(hot.getCellMeta(r, c)?.className || "");
          const fontToken = cls.split(" ").find((x: string) => x.startsWith("meta-font-"));
          const sizeToken = cls.split(" ").find((x: string) => x.startsWith("meta-size-"));
          const colorToken = cls.split(" ").find((x: string) => x.startsWith("meta-color-"));
          const fillToken = cls.split(" ").find((x: string) => x.startsWith("meta-fill-"));
          setIsBoldActive(cls.includes("meta-bold"));
          setIsItalicActive(cls.includes("meta-italic"));
          setIsUnderlineActive(cls.includes("meta-underline"));
          setIsStrikeActive(cls.includes("meta-strike"));
          if (fontToken) {
            setFontFamily(fontToken.replace("meta-font-", "").split("_").join(" "));
          }
          if (sizeToken) {
            setFontSize(sizeToken.replace("meta-size-", ""));
          }
          if (colorToken) {
            setTextColor(`#${colorToken.replace("meta-color-", "")}`);
          } else {
            setTextColor("#111827");
          }
          if (fillToken) {
            setFillColor(`#${fillToken.replace("meta-fill-", "")}`);
          } else {
            setFillColor("#ffffff");
          }
          if (cls.includes("meta-align-left")) setSelectedAlign("left");
          else if (cls.includes("meta-align-center")) setSelectedAlign("center");
          else if (cls.includes("meta-align-right")) setSelectedAlign("right");
          else if (cls.includes("meta-align-justify")) setSelectedAlign("justify");
          else setSelectedAlign(null);
          if (cls.includes("meta-valign-top")) setSelectedVAlign("top");
          else if (cls.includes("meta-valign-middle")) setSelectedVAlign("middle");
          else if (cls.includes("meta-valign-bottom")) setSelectedVAlign("bottom");
          else setSelectedVAlign(null);
        }}
        afterMergeCells={() => {
          if (readOnly) return;
          collectCurrentSheetFromHot(true);
        }}
        afterUnmergeCells={() => {
          if (readOnly) return;
          collectCurrentSheetFromHot(true);
        }}
        afterCreateRow={() => {}}
        afterCreateCol={() => {}}
        afterRemoveRow={() => {}}
        afterRemoveCol={() => {}}
      />
    </div>
    {isPreviewTruncated && (
      <div className="px-2 py-1 text-xs text-amber-700 border border-amber-200 rounded bg-amber-50">
        Preview mode showing first {previewRows} rows x {previewCols} columns for stability.
      </div>
    )}
    </div>
  );
};

export default HandsontableWorkbook;
