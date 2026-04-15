import React from "react";
import { View, Text, Image, StyleSheet } from "@react-pdf/renderer";

/** Persisted workbook shape (matches builder / submission payload). */
export type EmbeddedExcelPdfSheet = {
  name: string;
  grid: unknown[][];
  mergeCells?: Array<{ row: number; col: number; rowspan: number; colspan: number }>;
  cellMeta?: Array<{ row: number; col: number; className?: string }>;
  images?: Array<{
    row: number;
    col: number;
    rowspan?: number;
    colspan?: number;
    dataUrl: string;
  }>;
};

export type EmbeddedExcelPdfWorkbook = { sheets: EmbeddedExcelPdfSheet[] };

type EmbeddedSheetImage = NonNullable<EmbeddedExcelPdfSheet["images"]>[number];

const C = {
  border: "#cdd5dc",
  text: "#1e2a38",
  muted: "#5a6a7a",
  fillableBg: "#fffbe6",
  white: "#ffffff",
  sheetTitle: "#0f2341",
};

const MAX_ROWS = 260;
const MAX_COLS = 56;
const BASE_ROW_MIN_PT = 11;
const BASE_CELL_FONT_PT = 5.5;
const BASE_CELL_PADDING_PT = 1.5;
const BASE_TITLE_FONT_PT = 7;
const BASE_TITLE_PAD_H = 4;
const BASE_TITLE_PAD_V = 3;
const BASE_IMAGE_W = 36;
const BASE_IMAGE_H = 14;
const FIT_TARGET_ROWS = 46;
const FIT_TARGET_COLS = 16;
const FIT_MIN_SCALE = 0.44;

const es = StyleSheet.create({
  sheetBlock: {
    marginBottom: 10,
    borderWidth: 1,
    borderColor: C.border,
    borderRadius: 3,
    overflow: "hidden",
    backgroundColor: C.white,
  },
  sheetTitle: {
    fontSize: 7,
    fontFamily: "Helvetica-Bold",
    color: C.sheetTitle,
    paddingHorizontal: 4,
    paddingVertical: 3,
    backgroundColor: "#f4f6f8",
    borderBottomWidth: 1,
    borderBottomColor: C.border,
  },
  table: {
    width: "100%",
  },
  row: {
    flexDirection: "row",
    width: "100%",
    alignItems: "stretch",
  },
  cellInner: {
    fontSize: BASE_CELL_FONT_PT,
    color: C.text,
    lineHeight: 1.15,
  },
  truncateNote: {
    fontSize: 6,
    color: C.muted,
    marginTop: 4,
    paddingHorizontal: 2,
  },
});

type SheetScaleMetrics = {
  rowMinPt: number;
  cellFontPt: number;
  cellPaddingPt: number;
  titleFontPt: number;
  titlePadH: number;
  titlePadV: number;
  imageW: number;
  imageH: number;
};

function clamp(n: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, n));
}

/**
 * Heuristic fit: many rows and many columns both increase page pressure.
 * Scale down smoothly so dense sheets still fit a single PDF page.
 */
function computeSheetScaleMetrics(rows: number, cols: number): SheetScaleMetrics {
  const rowScale = rows > FIT_TARGET_ROWS ? FIT_TARGET_ROWS / rows : 1;
  const colScale = cols > FIT_TARGET_COLS ? FIT_TARGET_COLS / cols : 1;
  const scale = clamp(Math.min(rowScale, colScale), FIT_MIN_SCALE, 1);
  return {
    rowMinPt: BASE_ROW_MIN_PT * scale,
    cellFontPt: BASE_CELL_FONT_PT * scale,
    cellPaddingPt: BASE_CELL_PADDING_PT * scale,
    titleFontPt: BASE_TITLE_FONT_PT * scale,
    titlePadH: BASE_TITLE_PAD_H * scale,
    titlePadV: BASE_TITLE_PAD_V * scale,
    imageW: BASE_IMAGE_W * scale,
    imageH: BASE_IMAGE_H * scale,
  };
}

function toSafeGrid(raw: unknown): string[][] {
  if (!Array.isArray(raw) || raw.length === 0) return [[""]];
  const rows = raw.map((row) =>
    Array.isArray(row) ? row.map((c) => (c == null ? "" : String(c))) : [""],
  );
  return rows.length > 0 ? rows : [[""]];
}

function clipMerges(
  merges: NonNullable<EmbeddedExcelPdfSheet["mergeCells"]>,
  gridRows: number,
  gridCols: number,
): NonNullable<EmbeddedExcelPdfSheet["mergeCells"]> {
  if (!merges.length || gridRows < 1 || gridCols < 1) return [];
  const maxR = gridRows - 1;
  const maxC = gridCols - 1;
  const out: NonNullable<EmbeddedExcelPdfSheet["mergeCells"]> = [];
  for (const m of merges) {
    if (!m) continue;
    const r0 = +m.row;
    const c0 = +m.col;
    const rs = +m.rowspan;
    const cs = +m.colspan;
    if (![r0, c0, rs, cs].every((n) => Number.isFinite(n)) || rs < 1 || cs < 1) continue;
    const r1 = r0 + rs - 1;
    const c1 = c0 + cs - 1;
    const cr0 = Math.max(0, Math.min(maxR, r0));
    const cc0 = Math.max(0, Math.min(maxC, c0));
    const cr1 = Math.max(0, Math.min(maxR, r1));
    const cc1 = Math.max(0, Math.min(maxC, c1));
    if (cr1 < cr0 || cc1 < cc0) continue;
    const rowspan = cr1 - cr0 + 1;
    const colspan = cc1 - cc0 + 1;
    if (rowspan > 0 && colspan > 0) out.push({ row: cr0, col: cc0, rowspan, colspan });
  }
  return out;
}

function normalizeSheet(sheet: EmbeddedExcelPdfSheet, index: number): {
  name: string;
  grid: string[][];
  merges: NonNullable<EmbeddedExcelPdfSheet["mergeCells"]>;
  fillable: Set<string>;
  imagesByKey: Map<string, EmbeddedSheetImage>;
} {
  const grid = toSafeGrid(sheet?.grid);
  const gridRows = grid.length;
  const gridCols = Math.max(
    1,
    grid.reduce((w, row) => Math.max(w, Array.isArray(row) ? row.length : 0), 0),
  );
  const merges = clipMerges(Array.isArray(sheet?.mergeCells) ? sheet.mergeCells : [], gridRows, gridCols);
  const fillable = new Set<string>();
  for (const m of sheet?.cellMeta || []) {
    if (!m || !Number.isFinite(+m.row) || !Number.isFinite(+m.col)) continue;
    const cls = typeof m.className === "string" ? m.className : "";
    if (cls.includes("meta-fillable")) fillable.add(`${+m.row},${+m.col}`);
  }
  const imagesByKey = new Map<string, EmbeddedSheetImage>();
  for (const im of sheet?.images || []) {
    if (!im || !Number.isFinite(+im.row) || !Number.isFinite(+im.col) || !im.dataUrl) continue;
    imagesByKey.set(`${+im.row},${+im.col}`, im);
  }
  return {
    name: sheet?.name || `Sheet${index + 1}`,
    grid,
    merges,
    fillable,
    imagesByKey,
  };
}

function mergeAt(r: number, c: number, merges: NonNullable<EmbeddedExcelPdfSheet["mergeCells"]>) {
  for (const m of merges) {
    if (m.row === r && m.col === c) return m;
  }
  return null;
}

function isHSkip(
  r: number,
  c: number,
  merges: NonNullable<EmbeddedExcelPdfSheet["mergeCells"]>,
): boolean {
  for (const m of merges) {
    if (r !== m.row) continue;
    if (c > m.col && c < m.col + m.colspan) return true;
  }
  return false;
}

function verticalTailMerge(
  r: number,
  c: number,
  merges: NonNullable<EmbeddedExcelPdfSheet["mergeCells"]>,
) {
  for (const m of merges) {
    if (m.rowspan <= 1) continue;
    if (r <= m.row || r >= m.row + m.rowspan) continue;
    if (c === m.col) return m;
  }
  return null;
}

function cellStyle(
  widthPct: number,
  rowspan: number,
  isFillable: boolean,
  borders: { left: boolean; top: boolean },
  metrics: SheetScaleMetrics,
) {
  return {
    width: `${widthPct}%`,
    minHeight: metrics.rowMinPt * rowspan,
    borderLeftWidth: borders.left ? 1 : 0,
    borderTopWidth: borders.top ? 1 : 0,
    borderRightWidth: 1,
    borderBottomWidth: 1,
    borderColor: C.border,
    padding: metrics.cellPaddingPt,
    backgroundColor: isFillable ? C.fillableBg : C.white,
    justifyContent: "flex-start" as const,
  };
}

/** Continuation row under a rowspan merge — no top border so it meets the cell above. */
function vContCellStyle(widthPct: number, metrics: SheetScaleMetrics) {
  return {
    width: `${widthPct}%`,
    minHeight: metrics.rowMinPt,
    borderLeftWidth: 1,
    borderTopWidth: 0,
    borderRightWidth: 1,
    borderBottomWidth: 1,
    borderColor: C.border,
    padding: 0,
    backgroundColor: C.white,
  };
}

function SheetTable({
  name,
  grid,
  merges,
  fillable,
  imagesByKey,
  truncated,
  metrics,
}: {
  name: string;
  grid: string[][];
  merges: NonNullable<EmbeddedExcelPdfSheet["mergeCells"]>;
  fillable: Set<string>;
  imagesByKey: Map<string, EmbeddedSheetImage>;
  truncated: boolean;
  metrics: SheetScaleMetrics;
}) {
  const rows = grid.length;
  const cols = Math.max(
    1,
    grid.reduce((w, row) => Math.max(w, row.length), 0),
  );

  const body: React.ReactNode[] = [];
  for (let r = 0; r < rows; r++) {
    const segments: React.ReactNode[] = [];
    let c = 0;
    while (c < cols) {
      const vt = verticalTailMerge(r, c, merges);
      if (vt) {
        const w = (vt.colspan / cols) * 100;
        segments.push(<View key={`vt-${r}-${c}`} style={vContCellStyle(w, metrics)} />);
        c += vt.colspan;
        continue;
      }
      if (isHSkip(r, c, merges)) {
        c += 1;
        continue;
      }
      const anch = mergeAt(r, c, merges);
      if (anch) {
        const w = (anch.colspan / cols) * 100;
        const txt = grid[r]?.[c] ?? "";
        const fill = fillable.has(`${r},${c}`);
        const img = imagesByKey.get(`${r},${c}`);
        segments.push(
          <View
            key={`a-${r}-${c}`}
            style={cellStyle(w, anch.rowspan, fill, { left: c === 0, top: r === 0 }, metrics)}
          >
            {img?.dataUrl?.startsWith("data:") ? (
              <Image
                src={img.dataUrl}
                style={{ width: metrics.imageW, height: metrics.imageH, objectFit: "contain" }}
              />
            ) : null}
            <Text style={[es.cellInner, { fontSize: metrics.cellFontPt }]} wrap>
              {txt}
            </Text>
          </View>,
        );
        c += anch.colspan;
        continue;
      }
      const w = (1 / cols) * 100;
      const txt = grid[r]?.[c] ?? "";
      const fill = fillable.has(`${r},${c}`);
      const img = imagesByKey.get(`${r},${c}`);
      segments.push(
        <View
          key={`c-${r}-${c}`}
          style={cellStyle(w, 1, fill, { left: c === 0, top: r === 0 }, metrics)}
        >
          {img?.dataUrl?.startsWith("data:") ? (
            <Image
              src={img.dataUrl}
              style={{ width: metrics.imageW, height: metrics.imageH, objectFit: "contain" }}
            />
          ) : null}
          <Text style={[es.cellInner, { fontSize: metrics.cellFontPt }]} wrap>
            {txt}
          </Text>
        </View>,
      );
      c += 1;
    }
    body.push(
      <View key={`row-${r}`} style={es.row} wrap={false}>
        {segments}
      </View>,
    );
  }

  return (
    <View style={es.sheetBlock} wrap={false}>
      <Text
        style={[
          es.sheetTitle,
          {
            fontSize: metrics.titleFontPt,
            paddingHorizontal: metrics.titlePadH,
            paddingVertical: metrics.titlePadV,
          },
        ]}
      >
        {name}
      </Text>
      <View style={es.table}>{body}</View>
      {truncated ? (
        <Text style={es.truncateNote}>
          Grid truncated for PDF ({MAX_ROWS} rows × {MAX_COLS} columns max).
        </Text>
      ) : null}
    </View>
  );
}

export type EmbeddedExcelWorkbookPdfProps = {
  workbook: EmbeddedExcelPdfWorkbook | null | undefined;
};

/**
 * Renders embedded Excel workbook sheets as bordered tables in React-PDF (submission / template export).
 * Respects merge regions and `meta-fillable` cell highlighting.
 */
export const EmbeddedExcelWorkbookPdf: React.FC<EmbeddedExcelWorkbookPdfProps> = ({ workbook }) => {
  const sheets = workbook?.sheets;
  if (!Array.isArray(sheets) || sheets.length === 0) {
    return (
      <View style={[es.sheetBlock, { padding: 8 }]}>
        <Text style={es.cellInner}>—</Text>
      </View>
    );
  }

  return (
    <View wrap>
      {sheets.map((raw, idx) => {
        const n = normalizeSheet(raw, idx);
        let grid = n.grid;
        let truncated = false;
        const origMaxCol = grid.reduce((w, row) => Math.max(w, row.length), 0);
        if (grid.length > MAX_ROWS || origMaxCol > MAX_COLS) {
          truncated = true;
          grid = grid.slice(0, MAX_ROWS).map((row) => [...row].slice(0, MAX_COLS));
        }
        const gridRows = grid.length;
        let gridCols = Math.max(
          1,
          grid.reduce((w, row) => Math.max(w, row.length), 0),
        );
        const metrics = computeSheetScaleMetrics(gridRows, gridCols);
        grid = grid.map((row) => {
          const r = [...row];
          while (r.length < gridCols) r.push("");
          return r;
        });
        const merges = clipMerges(n.merges, gridRows, gridCols);
        const fill = new Set<string>();
        for (const key of n.fillable) {
          const [rs, cs] = key.split(",").map(Number);
          if (rs < gridRows && cs < gridCols) fill.add(key);
        }
        const imgs = new Map<string, EmbeddedSheetImage>();
        for (const [k, v] of n.imagesByKey) {
          const [rs, cs] = k.split(",").map(Number);
          if (rs < gridRows && cs < gridCols) imgs.set(k, v);
        }
        return (
          <View key={`${n.name}-${idx}`} wrap={false} break={idx > 0}>
            <SheetTable
              name={n.name}
              grid={grid}
              merges={merges}
              fillable={fill}
              imagesByKey={imgs}
              truncated={truncated}
              metrics={metrics}
            />
          </View>
        );
      })}
    </View>
  );
};
