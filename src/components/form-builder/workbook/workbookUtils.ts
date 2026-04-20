import { SheetData } from "./workbookTypes";

export const FORMULAS_CONFIG = { engine: null as any };
export const EMPTY_GRID_PLACEHOLDER: string[][] = [[""]];
export const YES_NO_PAIR_TOKEN_PREFIX = "meta-yesno-pair-";
export const SINGLE_CHECKBOX_CLASS = "meta-single-checkbox";

type CellMetaEntry = NonNullable<SheetData["cellMeta"]>[number];

export const cellCoordKey = (row: number, col: number) => `${row}:${col}`;

export const classNameHasFillable = (className?: string) =>
  String(className || "")
    .split(/\s+/)
    .filter(Boolean)
    .includes("meta-fillable");

export const extractYesNoPairToken = (className?: string) =>
  String(className || "")
    .split(/\s+/)
    .filter(Boolean)
    .find((token) => token.startsWith(YES_NO_PAIR_TOKEN_PREFIX));

export const toCheckboxChecked = (value: unknown) => {
  if (value === true || value === 1) return true;
  const normalized = String(value ?? "")
    .trim()
    .toLowerCase();
  return normalized === "true" || normalized === "1" || normalized === "yes";
};

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

const toSafeGrid = (rawGrid: unknown): string[][] => {
  if (!Array.isArray(rawGrid) || rawGrid.length === 0) return [[""]];
  const rows = rawGrid.map((row) =>
    Array.isArray(row) ? row.map((c) => (c == null ? "" : String(c))) : [""],
  );
  return rows.length > 0 ? rows : [[""]];
};

export const cloneEditableGrid = (rawGrid: unknown): string[][] => {
  const g = toSafeGrid(rawGrid);
  return g.map((row) => (Array.isArray(row) ? [...row] : [""]));
};

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

export const dedupeImagesByAnchor = (
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

export const dedupeCellMetaByCoordinate = (
  list: NonNullable<SheetData["cellMeta"]>,
): NonNullable<SheetData["cellMeta"]> => {
  const map = new Map<string, CellMetaEntry>();
  for (const raw of list) {
    if (!raw || !Number.isFinite(+raw.row) || !Number.isFinite(+raw.col)) continue;
    const row = +raw.row;
    const col = +raw.col;
    const key = cellCoordKey(row, col);
    const next: CellMetaEntry = {
      row,
      col,
      formula: typeof (raw as any).formula === "string" ? (raw as any).formula : undefined,
      formulaCachedValue:
        typeof (raw as any).formulaCachedValue === "string"
          ? (raw as any).formulaCachedValue
          : undefined,
      formulaWarning:
        typeof (raw as any).formulaWarning === "boolean"
          ? (raw as any).formulaWarning
          : undefined,
      className: raw.className ? String(raw.className) : undefined,
      type: typeof raw.type === "string" ? raw.type : undefined,
      checkedTemplate:
        typeof raw.checkedTemplate === "string" ? raw.checkedTemplate : undefined,
      uncheckedTemplate:
        typeof raw.uncheckedTemplate === "string"
          ? raw.uncheckedTemplate
          : undefined,
      dateFormat: typeof raw.dateFormat === "string" ? raw.dateFormat : undefined,
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
      formula: next.formula ?? prev.formula,
      formulaCachedValue: next.formulaCachedValue ?? prev.formulaCachedValue,
      formulaWarning:
        typeof next.formulaWarning === "boolean"
          ? next.formulaWarning
          : prev.formulaWarning,
      type: next.type ?? prev.type,
      checkedTemplate: next.checkedTemplate ?? prev.checkedTemplate,
      uncheckedTemplate: next.uncheckedTemplate ?? prev.uncheckedTemplate,
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

export const deepCloneSheet = (s: SheetData): SheetData => ({
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
      formula: typeof (m as any).formula === "string" ? (m as any).formula : undefined,
      formulaCachedValue:
        typeof (m as any).formulaCachedValue === "string"
          ? (m as any).formulaCachedValue
          : undefined,
      formulaWarning:
        typeof (m as any).formulaWarning === "boolean"
          ? (m as any).formulaWarning
          : undefined,
      className: m.className,
      type: m.type,
      checkedTemplate: m.checkedTemplate,
      uncheckedTemplate: m.uncheckedTemplate,
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

export const normalizeSheets = (input?: { sheets?: SheetData[] }): SheetData[] => {
  if (!Array.isArray(input?.sheets) || input.sheets.length === 0)
    return [{ name: "Sheet1", grid: [[""]] }];
  return input.sheets.map((sheet, i) => {
    const grid = toSafeGrid(sheet?.grid);
    const gridRows = grid.length;
    const gridCols = Math.max(
      1,
      grid.reduce((w, row) => Math.max(w, Array.isArray(row) ? row.length : 0), 0),
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
                formula: typeof m.formula === "string" ? m.formula : undefined,
                formulaCachedValue:
                  typeof m.formulaCachedValue === "string"
                    ? m.formulaCachedValue
                    : undefined,
                formulaWarning:
                  typeof m.formulaWarning === "boolean"
                    ? m.formulaWarning
                    : undefined,
                className:
                  typeof m.className === "string" ? m.className : undefined,
                type: typeof m.type === "string" ? m.type : undefined,
                checkedTemplate:
                  typeof m.checkedTemplate === "string"
                    ? m.checkedTemplate
                    : undefined,
                uncheckedTemplate:
                  typeof m.uncheckedTemplate === "string"
                    ? m.uncheckedTemplate
                    : undefined,
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

const dimListSignature = (arr?: number[]) => {
  if (!arr?.length) return "";
  if (arr.length > 400) {
    let sum = 0;
    for (let i = 0; i < arr.length; i++) sum += Number(arr[i]) || 0;
    return `${arr.length}:sum${sum}:a${arr[0]}:z${arr[arr.length - 1]}`;
  }
  return arr.join(",");
};

export const workbookSignature = (sheets: SheetData[]) =>
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

export const hotTableMountSignature = (sheets: SheetData[]) =>
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

export const mergeFillableMetaFromPrevSheet = (
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
    metaByKey.set(cellCoordKey(+m.row, +m.col), { ...m, row: +m.row, col: +m.col });
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

export const toRangeLabel = (
  range: { startRow: number; endRow: number; startCol: number; endCol: number } | null,
) => {
  if (!range) return "A1";
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
  const start = `${toColumnLabel(range.startCol)}${range.startRow + 1}`;
  const end = `${toColumnLabel(range.endCol)}${range.endRow + 1}`;
  return start === end ? start : `${start}:${end}`;
};

export const buildYesNoOppositeMap = (
  cellMeta?: NonNullable<SheetData["cellMeta"]>,
) => {
  const pairBuckets = new Map<string, Array<{ row: number; col: number }>>();
  for (const meta of cellMeta || []) {
    const pairToken = extractYesNoPairToken(meta.className);
    if (!pairToken) continue;
    const list = pairBuckets.get(pairToken) || [];
    list.push({ row: meta.row, col: meta.col });
    pairBuckets.set(pairToken, list);
  }
  const oppositeCellByKey = new Map<string, { row: number; col: number }>();
  for (const entries of pairBuckets.values()) {
    if (entries.length !== 2) continue;
    const a = entries[0];
    const b = entries[1];
    oppositeCellByKey.set(cellCoordKey(a.row, a.col), b);
    oppositeCellByKey.set(cellCoordKey(b.row, b.col), a);
  }
  return oppositeCellByKey;
};

