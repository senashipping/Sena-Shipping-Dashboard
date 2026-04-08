import React from "react";
import { HotTable } from "@handsontable/react";
import "handsontable/styles/handsontable.css";
import "handsontable/styles/ht-theme-main.css";
import { registerAllModules } from "handsontable/registry";
import { Button } from "../ui/button";
import { HyperFormula } from "hyperformula";
import ExcelJS from "exceljs";
import * as XLSX from "xlsx";

registerAllModules();

type SheetData = {
  name: string;
  grid: string[][];
  mergeCells?: Array<{ row: number; col: number; rowspan: number; colspan: number }>;
  cellMeta?: Array<{ row: number; col: number; className?: string; readOnly?: boolean }>;
  colWidthsPx?: number[];
  rowHeightsPx?: number[];
  tabColor?: string;
};

interface HandsontableWorkbookProps {
  data: { sheets: SheetData[] };
  onChange: (next: { sheets: SheetData[] }) => void;
  readOnly?: boolean;
}

const argbToHex = (argb?: string): string | null => {
  if (!argb || typeof argb !== "string") return null;
  if (argb.length === 8) return `#${argb.slice(2)}`;
  if (argb.length === 6) return `#${argb}`;
  return null;
};

const safeToText = (value: unknown): string => {
  if (value == null) return "";
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return `${value}`;
  }
  if (value instanceof Date) return value.toISOString();
  try {
    return JSON.stringify(value);
  } catch {
    return "";
  }
};

const cellValueToString = (value: any): string => {
  if (value == null) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "object") {
    if (Array.isArray(value.richText)) {
      return value.richText.map((x: any) => x?.text ?? "").join("");
    }
    if (typeof value.text === "string") return value.text;
    if (typeof value.hyperlink === "string" && typeof value.text === "string") return value.text;
    if ("result" in value) return safeToText(value.result);
    if ("formula" in value) return value.formula ? `=${safeToText(value.formula)}` : safeToText(value.result);
    if ("error" in value) return safeToText(value.error);
  }
  return safeToText(value);
};

const decodeA1 = (a1: string) => {
  const clean = a1.replace(/\$/g, "");
  const m = clean.match(/^([A-Z]+)(\d+)$/i);
  if (!m) return { row: 1, col: 1 };
  const letters = m[1].toUpperCase();
  let col = 0;
  for (let i = 0; i < letters.length; i++) col = col * 26 + (letters.charCodeAt(i) - 64);
  return { row: Number(m[2]), col };
};

const mergeStringToRegion = (merge: string) => {
  const range = merge.includes("!") ? merge.split("!").pop() || merge : merge;
  const parts = range.split(":");
  const start = decodeA1(parts[0]);
  const end = decodeA1(parts[1] || parts[0]);
  return {
    row: start.row - 1,
    col: start.col - 1,
    rowspan: Math.max(1, end.row - start.row + 1),
    colspan: Math.max(1, end.col - start.col + 1),
  };
};

const importWithXlsxFallback = async (buffer: ArrayBuffer): Promise<SheetData[]> => {
  const wb = XLSX.read(buffer, { type: "array", cellDates: true });
  const sheets: SheetData[] = [];
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    if (!ws) continue;
    const ref = ws["!ref"] || "A1";
    const range = XLSX.utils.decode_range(ref);
    const rows = Math.max(1, range.e.r - range.s.r + 1);
    const cols = Math.max(1, range.e.c - range.s.c + 1);
    const grid = Array.from({ length: rows }, (_, r) =>
      Array.from({ length: cols }, (_, c) => {
        const addr = XLSX.utils.encode_cell({ r: r + range.s.r, c: c + range.s.c });
        const cell = ws[addr] as any;
        if (!cell) return "";
        if (cell.w != null) return String(cell.w);
        if (cell.v == null) return "";
        return String(cell.v);
      })
    );
    const merges = Array.isArray(ws["!merges"])
      ? ws["!merges"].map((m: any) => ({
          row: m.s.r - range.s.r,
          col: m.s.c - range.s.c,
          rowspan: m.e.r - m.s.r + 1,
          colspan: m.e.c - m.s.c + 1,
        }))
      : [];
    const colWidthsPx = Array.from({ length: cols }, (_, i) => {
      const col = (ws["!cols"] || [])[i + range.s.c] as any;
      if (col?.wpx) return Math.round(col.wpx);
      if (col?.wch) return Math.round(col.wch * 7 + 8);
      return 80;
    });
    const rowHeightsPx = Array.from({ length: rows }, (_, i) => {
      const row = (ws["!rows"] || [])[i + range.s.r] as any;
      if (row?.hpx) return Math.round(row.hpx);
      if (row?.hpt) return Math.round((row.hpt * 96) / 72);
      return 24;
    });
    sheets.push({
      name: name || `Sheet${sheets.length + 1}`,
      grid,
      mergeCells: merges,
      colWidthsPx,
      rowHeightsPx,
      cellMeta: [],
    });
  }
  return sheets.length ? sheets : [{ name: "Sheet1", grid: [[""]] }];
};

const getClassToken = (prefix: string, raw: unknown): string | null => {
  const text = safeToText(raw).trim();
  if (!text) return null;
  return `${prefix}${text.replace(/\s+/g, "_")}`;
};

const HandsontableWorkbook: React.FC<HandsontableWorkbookProps> = ({
  data,
  onChange,
  readOnly = false,
}) => {
  const importInputId = React.useId();
  const importInputRef = React.useRef<HTMLInputElement | null>(null);
  const [activeSheetIndex, setActiveSheetIndex] = React.useState(0);
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
  const [fixedRowsTop, setFixedRowsTop] = React.useState(0);
  const [fixedColumnsStart, setFixedColumnsStart] = React.useState(0);
  const [toolbarError, setToolbarError] = React.useState("");
  const hotRef = React.useRef<any>(null);
  const safeSheets =
    Array.isArray(data?.sheets) && data.sheets.length > 0
      ? data.sheets
      : [{ name: "Sheet1", grid: [[""]] }];
  const activeSheet = safeSheets[Math.min(activeSheetIndex, safeSheets.length - 1)];
  const safeGrid =
    Array.isArray(activeSheet?.grid) && activeSheet.grid.length > 0 ? activeSheet.grid : [[""]];
  const cellMetaMap = React.useMemo(() => {
    const map = new Map<string, { className?: string; readOnly?: boolean }>();
    for (const meta of activeSheet?.cellMeta || []) {
      map.set(`${meta.row}:${meta.col}`, { className: meta.className, readOnly: meta.readOnly });
    }
    return map;
  }, [activeSheet?.cellMeta]);

  const emitWorkbook = (nextSheet: SheetData) => {
    const nextSheets = safeSheets.map((sheet, index) =>
      index === activeSheetIndex ? nextSheet : sheet
    );
    onChange({ sheets: nextSheets });
  };

  const collectAndEmitMeta = (nextGrid: string[][], includeMeta: boolean) => {
    const hot = hotRef.current?.hotInstance;
    const mergeCells =
      hot?.getPlugin?.("mergeCells")?.mergedCellsCollection?.mergedCells?.map((cell: any) => ({
        row: cell.row,
        col: cell.col,
        rowspan: cell.rowspan,
        colspan: cell.colspan,
      })) || [];
    let cellMeta = activeSheet.cellMeta;
    if (includeMeta) {
      const nextMeta: Array<{ row: number; col: number; className?: string; readOnly?: boolean }> = [];
      for (let r = 0; r < nextGrid.length; r++) {
        for (let c = 0; c < (nextGrid[r]?.length || 0); c++) {
          const meta = hot?.getCellMeta?.(r, c);
          if (meta?.className || meta?.readOnly) {
            nextMeta.push({
              row: r,
              col: c,
              className: meta.className,
              readOnly: meta.readOnly,
            });
          }
        }
      }
      cellMeta = nextMeta;
    }
    emitWorkbook({
      ...activeSheet,
      grid: nextGrid,
      mergeCells,
      cellMeta,
      colWidthsPx: hot?.getColHeader?.() ? activeSheet.colWidthsPx : activeSheet.colWidthsPx,
      rowHeightsPx: activeSheet.rowHeightsPx,
    });
  };

  const syncFromHot = (includeMeta = false) => {
    const hot = hotRef.current?.hotInstance;
    if (!hot) return;
    const next = (hot.getData?.() || safeGrid).map((row: any[]) =>
      row.map((cell) => (cell == null ? "" : String(cell)))
    );
    collectAndEmitMeta(next, includeMeta);
  };

  const getSelectedRange = () => {
    const hot = hotRef.current?.hotInstance;
    if (!hot) return null;
    const selected = hot.getSelectedLast?.();
    if (!selected) {
      return { startRow: 0, endRow: 0, startCol: 0, endCol: 0 };
    }
    const [r1, c1, r2, c2] = selected;
    return {
      startRow: Math.min(r1, r2),
      endRow: Math.max(r1, r2),
      startCol: Math.min(c1, c2),
      endCol: Math.max(c1, c2),
    };
  };

  const applyClassToSelection = (classToken: string, toggle = false) => {
    const hot = hotRef.current?.hotInstance;
    const range = getSelectedRange();
    if (!hot || !range || readOnly) return;
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
          : [...current.filter((x: string) => !x.startsWith(classToken.split("-")[0] + "-")), classToken];
        hot.setCellMeta(r, c, "className", next.join(" ").trim());
      }
    }
    hot.render();
    syncFromHot(true);
  };

  const setAlignment = (align: "left" | "center" | "right") => {
    applyClassToSelection(`meta-align-${align}`);
  };

  const setWrapText = () => {
    applyClassToSelection("meta-wrap", true);
  };

  const setFontStyle = (style: "bold" | "italic" | "underline" | "strike") => {
    applyClassToSelection(`meta-${style}`, true);
  };

  const applyFontFamily = () => applyClassToSelection(`meta-font-${fontFamily.replace(/\s+/g, "_")}`);
  const applyFontSize = () => applyClassToSelection(`meta-size-${fontSize}`);
  const applyTextColor = () => applyClassToSelection(`meta-color-${textColor.replace("#", "")}`);
  const applyFillColor = () => applyClassToSelection(`meta-fill-${fillColor.replace("#", "")}`);

  const formatSelectedAs = (kind: "number" | "currency" | "percent" | "date") => {
    const hot = hotRef.current?.hotInstance;
    const range = getSelectedRange();
    if (!hot || !range || readOnly) return;
    for (let r = range.startRow; r <= range.endRow; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        const raw = hot.getDataAtCell(r, c);
        if (raw == null || raw === "") continue;
        const num = Number(raw);
        if (kind === "date") {
          const d = new Date(raw);
          if (!Number.isNaN(d.getTime())) hot.setDataAtCell(r, c, d.toISOString().slice(0, 10));
        } else if (!Number.isNaN(num)) {
          if (kind === "number") hot.setDataAtCell(r, c, String(num));
          if (kind === "currency") hot.setDataAtCell(r, c, `$${num.toFixed(2)}`);
          if (kind === "percent") hot.setDataAtCell(r, c, `${(num * 100).toFixed(2)}%`);
        }
      }
    }
    syncFromHot(true);
  };

  const applyDropdownValidation = () => {
    const hot = hotRef.current?.hotInstance;
    const range = getSelectedRange();
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
    syncFromHot(true);
  };

  const applyDateCellType = () => {
    const hot = hotRef.current?.hotInstance;
    const range = getSelectedRange();
    if (!hot || !range || readOnly) return;
    for (let r = range.startRow; r <= range.endRow; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        hot.setCellMeta(r, c, "type", "date");
        hot.setCellMeta(r, c, "dateFormat", "YYYY-MM-DD");
        hot.setCellMeta(r, c, "correctFormat", true);
      }
    }
    hot.render();
    syncFromHot(true);
  };

  const sortSelectedColumn = (order: "asc" | "desc") => {
    const hot = hotRef.current?.hotInstance;
    const range = getSelectedRange();
    if (!hot || !range) return;
    hot.getPlugin("columnSorting").sort({ column: range.startCol, sortOrder: order });
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
    syncFromHot();
  };

  const importXlsx = async (file: File) => {
    setToolbarError("");
    try {
      if (!/\.xlsx$/i.test(file.name)) {
        setToolbarError("Only .xlsx files are supported.");
        return;
      }
      const buffer = await file.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      let nextSheets: SheetData[] = [];
      try {
        await workbook.xlsx.load(buffer);
        nextSheets = workbook.worksheets.map((ws, idx) => {
      const rows = Math.max(ws.rowCount || 1, 1);
      const cols = Math.max(ws.columnCount || 1, 1);
      const grid = Array.from({ length: rows }, (_, r) =>
        Array.from({ length: cols }, (_, c) => {
          const cell = ws.getCell(r + 1, c + 1);
          // Prefer Excel's rendered text to avoid [object Object] and preserve display value
          const displayed = typeof cell.text === "string" ? cell.text : "";
          if (displayed) return displayed;
          return cellValueToString(cell.value);
        })
      );
      const mergeCells = (Array.isArray((ws.model as any)?.merges) ? (ws.model as any).merges : [])
        .filter((m: unknown) => typeof m === "string" && m.length > 0)
        .map((m: string) => mergeStringToRegion(m));
      const cellMeta: Array<{ row: number; col: number; className?: string; readOnly?: boolean }> = [];
      for (let r = 1; r <= rows; r++) {
        for (let c = 1; c <= cols; c++) {
          const cell = ws.getCell(r, c);
          const tokens: string[] = [];
          const font = cell.font as any;
          const fill = cell.fill as any;
          const alignment = cell.alignment as any;
          if (font?.bold) tokens.push("meta-bold");
          if (font?.italic) tokens.push("meta-italic");
          if (font?.underline) tokens.push("meta-underline");
          if (font?.strike) tokens.push("meta-strike");
          const fontToken = getClassToken("meta-font-", font?.name);
          if (fontToken) tokens.push(fontToken);
          const sizeToken = getClassToken("meta-size-", font?.size);
          if (sizeToken) tokens.push(sizeToken);
          const fontHex = argbToHex(font?.color?.argb);
          if (fontHex) tokens.push(`meta-color-${fontHex.replace("#", "").toLowerCase()}`);
          const fillHex = argbToHex(fill?.fgColor?.argb || fill?.bgColor?.argb);
          if (fillHex) tokens.push(`meta-fill-${fillHex.replace("#", "").toLowerCase()}`);
          if (alignment?.horizontal === "left") tokens.push("meta-align-left");
          if (alignment?.horizontal === "center") tokens.push("meta-align-center");
          if (alignment?.horizontal === "right") tokens.push("meta-align-right");
          if (alignment?.wrapText) tokens.push("meta-wrap");
          if (tokens.length) {
            cellMeta.push({
              row: r - 1,
              col: c - 1,
              className: tokens.join(" "),
            });
          }
        }
      }
      const colWidthsPx = Array.from({ length: cols }, (_, i) => {
        const w = ws.getColumn(i + 1).width;
        return typeof w === "number" && w > 0 ? Math.round(w * 7 + 8) : 80;
      });
      const rowHeightsPx = Array.from({ length: rows }, (_, i) => {
        const h = ws.getRow(i + 1).height;
        return typeof h === "number" && h > 0 ? Math.round((h * 96) / 72) : 24;
      });
      return { name: ws.name || `Sheet${idx + 1}`, grid, mergeCells, cellMeta, colWidthsPx, rowHeightsPx };
        });
      } catch {
        // Fallback parser for workbooks ExcelJS cannot load.
        nextSheets = await importWithXlsxFallback(buffer);
      }
      onChange({ sheets: nextSheets.length ? nextSheets : [{ name: "Sheet1", grid: [[""]] }] });
      setActiveSheetIndex(0);
    } catch (error: any) {
      const rawMessage =
        typeof error?.message === "string"
          ? error.message
          : "Unknown parse error";
      setToolbarError(`Import failed: ${rawMessage}`);
    }
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
    onChange({ sheets: nextSheets });
    setActiveSheetIndex(nextSheets.length - 1);
  };

  const moveSheet = (direction: "left" | "right") => {
    const idx = activeSheetIndex;
    const target = direction === "left" ? idx - 1 : idx + 1;
    if (target < 0 || target >= safeSheets.length) return;
    const next = [...safeSheets];
    const [moved] = next.splice(idx, 1);
    next.splice(target, 0, moved);
    onChange({ sheets: next });
    setActiveSheetIndex(target);
  };

  const applySheetColor = (color: string) => {
    const nextSheets = safeSheets.map((s, idx) =>
      idx === activeSheetIndex ? { ...s, tabColor: color } : s
    );
    onChange({ sheets: nextSheets });
  };

  const applyFormulaBar = () => {
    const hot = hotRef.current?.hotInstance;
    const range = getSelectedRange();
    if (!hot || !range || readOnly) return;
    hot.setDataAtCell(range.startRow, range.startCol, formulaInput);
    syncFromHot();
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
        [class*="meta-font-"] { font-family: Arial, sans-serif; }
      `}</style>
      {!readOnly && (
        <div className="flex flex-wrap items-center gap-1 p-2 border rounded-md bg-slate-50">
          <Button type="button" size="sm" variant="outline" onClick={() => setFontStyle("bold")}>B</Button>
          <Button type="button" size="sm" variant="outline" onClick={() => setFontStyle("italic")}>I</Button>
          <Button type="button" size="sm" variant="outline" onClick={() => setFontStyle("underline")}>U</Button>
          <Button type="button" size="sm" variant="outline" onClick={() => setFontStyle("strike")}>S</Button>
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
          <Button type="button" size="sm" variant="outline" onClick={() => setAlignment("left")}>Left</Button>
          <Button type="button" size="sm" variant="outline" onClick={() => setAlignment("center")}>Center</Button>
          <Button type="button" size="sm" variant="outline" onClick={() => setAlignment("right")}>Right</Button>
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
          <input
            ref={importInputRef}
            type="file"
            accept=".xlsx"
            className="hidden"
            id={importInputId}
            onChange={(e) => {
              const file = e.target.files?.[0];
              if (file) importXlsx(file);
              e.currentTarget.value = "";
            }}
          />
          <Button
            type="button"
            size="sm"
            variant="outline"
            onClick={() => importInputRef.current?.click()}
          >
            Import .xlsx
          </Button>
          <Button type="button" size="sm" variant="outline" onClick={exportXlsx}>Export .xlsx</Button>
          <Button type="button" size="sm" variant="outline" onClick={exportCsv}>CSV</Button>
        </div>
      )}
      {!!toolbarError && (
        <div className="px-2 py-1 text-xs text-red-700 border border-red-200 rounded bg-red-50">
          {toolbarError}
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
        {safeSheets.map((sheet, index) => (
          <Button
            key={`${sheet.name}-${index}`}
            type="button"
            variant={index === activeSheetIndex ? "default" : "outline"}
            size="sm"
            style={sheet.tabColor ? { backgroundColor: sheet.tabColor, color: "#111827" } : undefined}
            onClick={() => setActiveSheetIndex(index)}
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
                  onChange({ sheets: nextSheets });
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
              onChange({ sheets: nextSheets });
              setActiveSheetIndex(nextSheets.length - 1);
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
        data={safeGrid}
        themeName="ht-theme-main"
        rowHeaders
        colHeaders
        licenseKey="non-commercial-and-evaluation"
        readOnly={readOnly}
        width="100%"
        stretchH="all"
        height={320}
        formulas={{ engine: HyperFormula }}
        mergeCells={activeSheet.mergeCells || true}
        filters
        dropdownMenu
        columnSorting
        hiddenRows={{ indicators: true }}
        hiddenColumns={{ indicators: true }}
        multiColumnSorting
        manualColumnFreeze
        autoColumnSize
        autoRowSize
        fillHandle
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
                  make_read_only: {},
                  undo: {},
                  redo: {},
                },
              }
        }
        className="ht-theme-main"
        manualRowResize
        manualColumnResize
        colWidths={activeSheet.colWidthsPx as any}
        rowHeights={activeSheet.rowHeightsPx as any}
        wordWrap
        autoWrapRow
        autoWrapCol
        cells={(row, col) => {
          const meta = cellMetaMap.get(`${row}:${col}`);
          const cp: any = {};
          if (meta?.className) cp.className = meta.className;
          if (meta?.readOnly) cp.readOnly = true;
          const tokens = (meta?.className || "").split(" ").filter(Boolean);
          const fontToken = tokens.find((t) => t.startsWith("meta-font-"));
          const sizeToken = tokens.find((t) => t.startsWith("meta-size-"));
          const colorToken = tokens.find((t) => t.startsWith("meta-color-"));
          const fillToken = tokens.find((t) => t.startsWith("meta-fill-"));
          if (fontToken || sizeToken || colorToken || fillToken) {
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
              if (base) {
                base(instance, td, rowIndex, colIndex, prop, value, cellProperties);
              } else {
                td.textContent = value == null ? "" : String(value);
              }
              if (fontToken) td.style.fontFamily = fontToken.replace("meta-font-", "").split("_").join(" ");
              if (sizeToken) td.style.fontSize = `${sizeToken.replace("meta-size-", "")}px`;
              if (colorToken) td.style.color = `#${colorToken.replace("meta-color-", "")}`;
              if (fillToken) td.style.backgroundColor = `#${fillToken.replace("meta-fill-", "")}`;
              return td;
            };
          }
          return cp;
        }}
        afterChange={(changes, source) => {
          if (!changes || source === "loadData" || readOnly) return;
          syncFromHot(false);
        }}
        afterSelectionEnd={(r, c) => {
          const hot = hotRef.current?.hotInstance;
          if (!hot) return;
          const v = hot.getDataAtCell(r, c);
          setFormulaInput(v == null ? "" : String(v));
        }}
        afterMergeCells={() => {
          if (readOnly) return;
          syncFromHot(true);
        }}
        afterUnmergeCells={() => {
          if (readOnly) return;
          syncFromHot(true);
        }}
        afterCreateRow={() => {
          if (readOnly) return;
          syncFromHot(false);
        }}
        afterCreateCol={() => {
          if (readOnly) return;
          syncFromHot(false);
        }}
        afterRemoveRow={() => {
          if (readOnly) return;
          syncFromHot(false);
        }}
        afterRemoveCol={() => {
          if (readOnly) return;
          syncFromHot(false);
        }}
      />
    </div>
    </div>
  );
};

export default HandsontableWorkbook;
