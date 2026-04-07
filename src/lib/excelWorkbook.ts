/**
 * Load .xlsx workbooks into an editable grid using ExcelJS (values, merges, column widths,
 * cell styling). Bounds include merges and column definitions so rows/columns are not dropped.
 */

import type { CSSProperties } from "react";
import type { Border, Cell, Color, Fill, Font, Alignment, Worksheet } from "exceljs";
import type { RawCellContent } from "hyperformula";
import type { HyperFormula } from "hyperformula";
import {
  buildHyperFormulaEngine,
  cellToHFContent,
  formatHyperFormulaDisplay,
  hfSheetKey,
} from "./excelHyperFormula";

/** Serializable CSS-oriented style for one cell (master cell for merges). */
export interface EmbeddedExcelCellStyle {
  fontFamily?: string;
  fontSizePt?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  bgColor?: string;
  borderTop?: string;
  borderRight?: string;
  borderBottom?: string;
  borderLeft?: string;
  textAlign?: "left" | "center" | "right" | "justify";
  verticalAlign?: "top" | "middle" | "bottom";
  wrapText?: boolean;
  /** Excel textRotation (degrees; 255 = stacked vertical). */
  textRotation?: number;
}

/** 1-based Excel extents used for HyperFormula (full `hfFullSheet` matrix). */
export interface ExcelSheetBounds {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

export interface EmbeddedExcelSheetData {
  name: string;
  grid: string[][];
  colWidthsPx?: number[];
  /** Row height in px when Excel specifies one (optional per row). */
  rowHeightsPx?: (number | undefined)[];
  mergeRegions?: Array<{ r: number; c: number; rowspan: number; colspan: number }>;
  /** Parallel to `grid`; only meaningful on merge master / normal cells (slaves may be empty). */
  cellStyles?: (EmbeddedExcelCellStyle | undefined)[][];
  /**
   * Full worksheet-sized matrix for HyperFormula (`bottom`×`right`), 1-based Excel addressing.
   * Omitted when saved to API — rebuilt from the file on load.
   */
  hfFullSheet?: RawCellContent[][];
  /** Same bounds as `hfFullSheet` dimensions. */
  excelBounds?: ExcelSheetBounds;
  /** Parallel to `grid`: cell contains an Excel formula (saved submissions do not overwrite these). */
  formulaCells?: boolean[][];
  /** Worksheet images/logos positioned over the rendered grid. */
  images?: Array<{ src: string; left: number; top: number; width: number; height: number }>;
}

export interface EmbeddedExcelFieldValue {
  sheets: EmbeddedExcelSheetData[];
  sourceUrl?: string;
}

/** Remove large runtime-only fields before persisting submission / form state. */
export function stripEmbeddedExcelRuntimeFields(data: EmbeddedExcelFieldValue): EmbeddedExcelFieldValue {
  return {
    ...data,
    sheets: data.sheets.map((s) => {
      const { hfFullSheet: _h, ...rest } = s;
      return rest;
    }),
  };
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

function mediaEntryToDataUrl(entry: unknown): string | undefined {
  if (!entry || typeof entry !== "object") return undefined;
  const e = entry as {
    base64?: string;
    extension?: string;
    buffer?: ArrayBuffer | Uint8Array | { data?: number[] };
  };
  const ext = (e.extension ?? "png").toLowerCase();
  if (typeof e.base64 === "string" && e.base64.length) {
    const base = e.base64.startsWith("data:") ? e.base64.split(",").pop() ?? "" : e.base64;
    return `data:image/${ext};base64,${base}`;
  }
  const buf = e.buffer;
  if (buf instanceof Uint8Array) {
    return `data:image/${ext};base64,${bytesToBase64(buf)}`;
  }
  if (buf instanceof ArrayBuffer) {
    return `data:image/${ext};base64,${bytesToBase64(new Uint8Array(buf))}`;
  }
  if (buf && typeof buf === "object" && Array.isArray((buf as { data?: number[] }).data)) {
    return `data:image/${ext};base64,${bytesToBase64(new Uint8Array((buf as { data: number[] }).data))}`;
  }
  return undefined;
}

function mediaEntryIdCandidates(entry: unknown, index: number): number[] {
  const out = new Set<number>([index, index + 1]);
  if (entry && typeof entry === "object") {
    const e = entry as { index?: unknown; imageId?: unknown; id?: unknown };
    const ids = [e.index, e.imageId, e.id];
    for (const id of ids) {
      if (typeof id === "number" && Number.isFinite(id)) out.add(id);
      if (typeof id === "string" && /^\d+$/.test(id)) out.add(Number(id));
    }
  }
  return Array.from(out);
}

function sizePrefixSums(sizes: number[], fallback: number): number[] {
  const out: number[] = [0];
  for (let i = 0; i < sizes.length; i++) out.push(out[i] + (sizes[i] ?? fallback));
  return out;
}

function positionFromCoord(coord: number, offsets: number[]): number {
  const base = Math.floor(coord);
  const frac = coord - base;
  if (base <= 0) return coord * (offsets[1] - offsets[0] || 0);
  if (base >= offsets.length - 1) {
    const last = offsets[offsets.length - 1] ?? 0;
    return last + frac * ((offsets[offsets.length - 1] ?? 0) - (offsets[offsets.length - 2] ?? 0));
  }
  return offsets[base] + frac * (offsets[base + 1] - offsets[base]);
}

function extractWorksheetImages(
  ws: Worksheet,
  resolveImage: (imageId: number) => string | undefined,
  leftExcelCol: number,
  topExcelRow: number,
  numCols: number,
  numRows: number,
  colWidthsPx: number[],
  rowHeightsPx: (number | undefined)[],
): Array<{ src: string; left: number; top: number; width: number; height: number }> {
  const getImages = (ws as unknown as { getImages?: () => Array<{ imageId: number; range: unknown }> }).getImages;
  if (!getImages) return [];
  const images = getImages.call(ws) ?? [];
  if (!images.length) return [];

  const colOffsets = sizePrefixSums(colWidthsPx, 72);
  const rowOffsets = sizePrefixSums(rowHeightsPx.map((h) => h ?? 22), 22);
  const out: Array<{ src: string; left: number; top: number; width: number; height: number }> = [];

  for (const img of images) {
    const src = resolveImage(img.imageId);
    if (!src) continue;
    const r = img.range as
      | {
          tl?: { nativeCol?: number; nativeRow?: number; col?: number; row?: number };
          br?: { nativeCol?: number; nativeRow?: number; col?: number; row?: number };
          ext?: { width?: number; height?: number };
        }
      | undefined;
    const tlCol = r?.tl?.nativeCol ?? r?.tl?.col ?? 0;
    const tlRow = r?.tl?.nativeRow ?? r?.tl?.row ?? 0;
    const brCol = r?.br?.nativeCol ?? r?.br?.col;
    const brRow = r?.br?.nativeRow ?? r?.br?.row;

    const relTlCol = tlCol - (leftExcelCol - 1);
    const relTlRow = tlRow - (topExcelRow - 1);
    const relBrCol = brCol != null ? brCol - (leftExcelCol - 1) : undefined;
    const relBrRow = brRow != null ? brRow - (topExcelRow - 1) : undefined;

    const leftPx = positionFromCoord(relTlCol, colOffsets);
    const topPx = positionFromCoord(relTlRow, rowOffsets);
    let widthPx = (r?.ext?.width ?? 0);
    let heightPx = (r?.ext?.height ?? 0);
    if (relBrCol != null && relBrRow != null) {
      widthPx = Math.max(8, positionFromCoord(relBrCol, colOffsets) - leftPx);
      heightPx = Math.max(8, positionFromCoord(relBrRow, rowOffsets) - topPx);
    }

    const maxW = colOffsets[numCols] ?? colOffsets[colOffsets.length - 1] ?? 0;
    const maxH = rowOffsets[numRows] ?? rowOffsets[rowOffsets.length - 1] ?? 0;
    if (leftPx > maxW || topPx > maxH) continue;
    out.push({
      src,
      left: Math.max(0, leftPx),
      top: Math.max(0, topPx),
      width: Math.max(8, widthPx),
      height: Math.max(8, heightPx),
    });
  }
  return out;
}

/** Default Office theme colors (approx. Excel 2013+); used when only `theme` index is stored. */
const THEME_SRGB: Record<number, string> = {
  0: "#FFFFFF",
  1: "#000000",
  2: "#E7E6E6",
  3: "#44546A",
  4: "#5B9BD5",
  5: "#ED7D31",
  6: "#A5A5A5",
  7: "#FFC000",
  8: "#4472C4",
  9: "#70AD47",
  10: "#0563C1",
  11: "#954F72",
};

/** Excel legacy indexed palette (subset; 0–63). */
const INDEXED_SRGB: Record<number, string> = {
  0: "#000000",
  1: "#FFFFFF",
  2: "#FF0000",
  3: "#00FF00",
  4: "#0000FF",
  5: "#FFFF00",
  6: "#FF00FF",
  7: "#00FFFF",
  8: "#000000",
  9: "#FFFFFF",
  10: "#FF0000",
  11: "#00FF00",
  12: "#0000FF",
  13: "#FFFF00",
  14: "#FF00FF",
  15: "#00FFFF",
  16: "#800000",
  17: "#008000",
  18: "#000080",
  19: "#808000",
  20: "#800080",
  21: "#008080",
  22: "#C0C0C0",
  23: "#808080",
  24: "#9999FF",
  25: "#993366",
  26: "#FFFFCC",
  27: "#CCFFFF",
  28: "#660066",
  29: "#FF8080",
  30: "#0066CC",
  31: "#CCCCFF",
  32: "#000080",
  33: "#FF00FF",
  34: "#FFFF00",
  35: "#00FFFF",
  36: "#800080",
  37: "#800000",
  38: "#008080",
  39: "#0000FF",
  40: "#00CCFF",
  41: "#CCFFFF",
  42: "#CCFFCC",
  43: "#FFFF99",
  44: "#99CCFF",
  45: "#FF99CC",
  46: "#CC99FF",
  47: "#FFCC99",
  48: "#3366FF",
  49: "#33CCCC",
  50: "#99CC00",
  51: "#FFCC00",
  52: "#FF9900",
  53: "#FF6600",
  54: "#666699",
  55: "#969696",
  56: "#003366",
  57: "#339966",
  58: "#003300",
  59: "#333300",
  60: "#993300",
  61: "#993366",
  62: "#333399",
  63: "#333333",
};

function argbToCss(c?: Partial<Color>): string | undefined {
  if (!c?.argb || typeof c.argb !== "string") return undefined;
  const a = c.argb;
  if (a.length === 8) return `#${a.slice(2)}`;
  if (a.length === 6) return `#${a}`;
  return undefined;
}

/** ECMA-376 tint: -1..1 applied to RGB. */
function applyTintToRgb(hex: string, tint: number): string {
  const h = hex.replace("#", "");
  if (h.length !== 6) return hex;
  let r = parseInt(h.slice(0, 2), 16);
  let g = parseInt(h.slice(2, 4), 16);
  let b = parseInt(h.slice(4, 6), 16);
  const adjust = (v: number) => {
    if (tint < 0) return Math.round(v * (1 + tint));
    return Math.round(v * (1 - tint) + 255 * tint);
  };
  r = Math.min(255, Math.max(0, adjust(r)));
  g = Math.min(255, Math.max(0, adjust(g)));
  b = Math.min(255, Math.max(0, adjust(b)));
  return `#${r.toString(16).padStart(2, "0")}${g.toString(16).padStart(2, "0")}${b.toString(16).padStart(2, "0")}`;
}

function readTint(c: Partial<Color> & { tint?: number }): number | undefined {
  if (typeof c.tint === "number" && c.tint !== 0) return c.tint;
  return undefined;
}

function resolveColor(c?: Partial<Color>): string | undefined {
  if (!c) return undefined;
  const direct = argbToCss(c);
  if (direct) return direct;
  const tint = readTint(c);
  if (c.theme !== undefined && c.theme !== null) {
    const base = THEME_SRGB[c.theme] ?? THEME_SRGB[1];
    return tint !== undefined ? applyTintToRgb(base, tint) : base;
  }
  const idx = (c as Partial<Color> & { indexed?: number }).indexed;
  if (idx !== undefined && INDEXED_SRGB[idx]) {
    return INDEXED_SRGB[idx];
  }
  return undefined;
}

function borderSide(b?: Partial<Border>): string | undefined {
  if (!b?.style) return undefined;
  const color = resolveColor(b.color) ?? "#000000";
  if (b.style === "double") {
    return `3px double ${color}`;
  }
  const w = b.style === "thick" || b.style === "medium" ? "2px" : "1px";
  return `${w} solid ${color}`;
}

const PATTERN_BG: Partial<Record<string, string>> = {
  darkGray: "#808080",
  mediumGray: "#C0C0C0",
  lightGray: "#D3D3D3",
  gray125: "#FCFCFC",
  gray0625: "#F2F2F2",
};

function fillToBg(f?: Fill): string | undefined {
  if (!f) return undefined;
  if (f.type === "gradient" && "stops" in f && f.stops?.length) {
    const c = f.stops[0].color;
    return resolveColor(c) ?? argbToCss(c);
  }
  if (f.type !== "pattern") return undefined;
  if (f.pattern === "none" || f.pattern === undefined) return undefined;
  const solid = resolveColor(f.fgColor) ?? resolveColor(f.bgColor) ?? argbToCss(f.fgColor) ?? argbToCss(f.bgColor);
  if (solid) return solid;
  if (f.pattern && PATTERN_BG[f.pattern]) return PATTERN_BG[f.pattern];
  return undefined;
}

function mapHorizontal(
  h?: Alignment["horizontal"],
): EmbeddedExcelCellStyle["textAlign"] | undefined {
  if (!h) return undefined;
  if (h === "center" || h === "centerContinuous") return "center";
  if (h === "distributed" || h === "fill") return "justify";
  return h as EmbeddedExcelCellStyle["textAlign"];
}

function mapVertical(
  v?: Alignment["vertical"],
): EmbeddedExcelCellStyle["verticalAlign"] | undefined {
  if (!v) return undefined;
  if (v === "distributed" || v === "justify") return "middle";
  return v as EmbeddedExcelCellStyle["verticalAlign"];
}

function mapTextRotation(a?: Partial<Alignment>): number | undefined {
  if (a?.textRotation === undefined || a?.textRotation === null) return undefined;
  const tr = a.textRotation;
  if (tr === "vertical") return 255;
  if (typeof tr === "number") {
    if (tr === 0) return undefined;
    return tr;
  }
  return undefined;
}

function extractCellStyle(cell: Cell): EmbeddedExcelCellStyle {
  const c = cell.isMerged ? cell.master : cell;
  const font = c.font as Partial<Font> | undefined;
  const fill = c.fill as Fill | undefined;
  const border = c.border;
  const alignment = c.alignment as Partial<Alignment> | undefined;

  const fontFamily = font?.name ? (font.name.includes(" ") ? `"${font.name}"` : font.name) : undefined;

  return {
    fontFamily,
    fontSizePt: typeof font?.size === "number" ? font.size : undefined,
    bold: font?.bold === true,
    italic: font?.italic === true,
    underline: font?.underline === true || font?.underline === "single",
    color: resolveColor(font?.color),
    bgColor: fillToBg(fill),
    borderTop: borderSide(border?.top),
    borderRight: borderSide(border?.right),
    borderBottom: borderSide(border?.bottom),
    borderLeft: borderSide(border?.left),
    textAlign: mapHorizontal(alignment?.horizontal),
    verticalAlign: mapVertical(alignment?.vertical),
    wrapText: alignment?.wrapText === true,
    textRotation: mapTextRotation(alignment),
  };
}

function cellDisplayText(cell: Cell): string {
  const c = cell.isMerged ? cell.master : cell;
  const t = c.text;
  if (t != null && t !== "") return t;
  const v = c.value;
  if (v == null || v === "") return "";
  if (typeof v === "object" && v !== null && "richText" in v) {
    const rt = v as { richText?: Array<{ text: string }> };
    return rt.richText?.map((x) => x.text).join("") ?? "";
  }
  if (typeof v === "object" && v !== null && "result" in v) {
    const r = (v as { result?: unknown }).result;
    return r != null ? String(r) : "";
  }
  if (v instanceof Date) return v.toLocaleString();
  return String(v);
}

/** Decode A1 → 1-based row and column (Excel). */
export function decodeA1(a1: string): { row: number; col: number } {
  const s = a1.replace(/\$/g, "").trim();
  const m = s.match(/^([A-Za-z]+)(\d+)$/);
  if (!m) return { row: 1, col: 1 };
  const letters = m[1].toUpperCase();
  const row = parseInt(m[2], 10);
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return { row, col };
}

function boundsFromMergeStrings(merges: string[] | undefined): {
  top: number;
  left: number;
  bottom: number;
  right: number;
} | null {
  if (!merges?.length) return null;
  let top = Infinity;
  let left = Infinity;
  let bottom = 0;
  let right = 0;
  for (const raw of merges) {
    const clean = raw.includes("!") ? raw.split("!").pop()! : raw;
    const parts = clean.includes(":") ? clean.split(":") : [clean, clean];
    const tl = decodeA1(parts[0].trim());
    const br = decodeA1(parts[1].trim());
    top = Math.min(top, tl.row, br.row);
    left = Math.min(left, tl.col, br.col);
    bottom = Math.max(bottom, tl.row, br.row);
    right = Math.max(right, tl.col, br.col);
  }
  if (top === Infinity) return null;
  return { top, left, bottom, right };
}

/**
 * Column `<cols>` in xlsx often spans 1–16384 even when the form is small. Only widen up to
 * this many columns past the union of cell dimensions + merges so we do not allocate millions of cells.
 */
const COL_METADATA_PADDING = 12;

/**
 * `ws.rowCount` can be the last row index in the sheet (very large). Only extend a few rows past real content.
 */
const ROW_TAIL_PADDING = 12;

/** Prevent browser freeze: cap grid size (product) for synchronous parse + DOM. */
const MAX_IMPORT_CELLS = 1_200_000;

function boundsFromColumnDefinitions(
  ws: Worksheet,
  contentRight: number,
): { right: number } | null {
  const cols = ws.columns;
  if (!cols?.length) return null;
  let right = 0;
  for (const col of cols) {
    if (col && typeof col.number === "number") {
      right = Math.max(right, col.number);
    }
  }
  if (right <= 0) return null;
  const capRight = contentRight + COL_METADATA_PADDING;
  return {
    right: Math.min(right, capRight),
  };
}

async function yieldToBrowser(): Promise<void> {
  await new Promise<void>((resolve) => {
    setTimeout(resolve, 0);
  });
}

/**
 * Union of ExcelJS dimensions, merges, and capped `<cols>` / row tail.
 * Avoids full `eachRow` scans and uncapped column ranges that freeze the UI.
 */
function computeExpandedBounds(ws: Worksheet): { top: number; left: number; bottom: number; right: number } | null {
  const dim = ws.dimensions;
  let top = dim.top;
  let left = dim.left;
  let bottom = dim.bottom;
  let right = dim.right;

  const mb = boundsFromMergeStrings((ws.model.merges ?? []) as string[]);
  const contentTop = mb ? Math.min(dim.top, mb.top) : dim.top;
  const contentLeft = mb ? Math.min(dim.left, mb.left) : dim.left;
  const contentBottom = mb ? Math.max(dim.bottom, mb.bottom) : dim.bottom;
  const contentRight = mb ? Math.max(dim.right, mb.right) : dim.right;
  top = contentTop;
  left = contentLeft;
  bottom = contentBottom;
  right = contentRight;

  const cb = boundsFromColumnDefinitions(ws, contentRight);
  if (cb) {
    right = Math.max(right, cb.right);
  }

  if (typeof ws.rowCount === "number" && ws.rowCount > 0) {
    const cappedTail = Math.min(ws.rowCount, contentBottom + ROW_TAIL_PADDING);
    bottom = Math.max(bottom, cappedTail);
  }

  if (bottom < top || right < left) return null;

  let numRows = bottom - top + 1;
  let numCols = right - left + 1;
  const cells = numRows * numCols;
  if (cells > MAX_IMPORT_CELLS) {
    const ratio = Math.sqrt(MAX_IMPORT_CELLS / cells);
    const newRows = Math.max(1, Math.floor(numRows * ratio));
    const newCols = Math.max(1, Math.floor(numCols * ratio));
    bottom = top + newRows - 1;
    right = left + newCols - 1;
  }

  return { top, left, bottom, right };
}

function parseMergeRangeString(
  mergeRange: string,
  sheetStartRow: number,
  sheetStartCol: number,
  numRows: number,
  numCols: number,
): { r: number; c: number; rowspan: number; colspan: number } | null {
  const clean = mergeRange.includes("!") ? mergeRange.split("!").pop()! : mergeRange;
  const part = clean.includes(":") ? clean.split(":") : [clean, clean];
  const tl = decodeA1(part[0].trim());
  const br = decodeA1(part[1].trim());
  const r = tl.row - sheetStartRow;
  const c = tl.col - sheetStartCol;
  if (r < 0 || c < 0 || r >= numRows || c >= numCols) return null;
  let rowspan = br.row - tl.row + 1;
  let colspan = br.col - tl.col + 1;
  rowspan = Math.min(rowspan, numRows - r);
  colspan = Math.min(colspan, numCols - c);
  if (rowspan < 1 || colspan < 1) return null;
  return { r, c, rowspan, colspan };
}

function mergeRegionsFromWorksheet(
  merges: string[] | undefined,
  sheetStartRow: number,
  sheetStartCol: number,
  numRows: number,
  numCols: number,
): Array<{ r: number; c: number; rowspan: number; colspan: number }> {
  const out: Array<{ r: number; c: number; rowspan: number; colspan: number }> = [];
  if (!merges?.length) return out;
  for (const m of merges) {
    const parsed = parseMergeRangeString(m, sheetStartRow, sheetStartCol, numRows, numCols);
    if (parsed) out.push(parsed);
  }
  return out;
}

function hasMeaningfulStyle(s?: EmbeddedExcelCellStyle): boolean {
  if (!s) return false;
  return Boolean(
    s.fontFamily ||
      s.fontSizePt ||
      s.bold ||
      s.italic ||
      s.underline ||
      s.color ||
      s.bgColor ||
      s.borderTop ||
      s.borderRight ||
      s.borderBottom ||
      s.borderLeft ||
      s.textAlign ||
      s.verticalAlign ||
      s.wrapText ||
      s.textRotation,
  );
}

export function columnLetter(zeroBasedIndex: number): string {
  let n = zeroBasedIndex + 1;
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function cellStyleToCss(style: EmbeddedExcelCellStyle | undefined): CSSProperties {
  if (!style) return {};
  const rot = style.textRotation;
  let transform: string | undefined;
  if (rot === 255) {
    transform = "rotate(-90deg)";
  } else if (typeof rot === "number" && rot > 0 && rot <= 180) {
    transform = `rotate(${-rot}deg)`;
  }
  return {
    fontFamily: style.fontFamily,
    fontSize: style.fontSizePt != null ? `${style.fontSizePt}pt` : undefined,
    fontWeight: style.bold ? 700 : undefined,
    fontStyle: style.italic ? "italic" : undefined,
    textDecoration: style.underline ? "underline" : undefined,
    color: style.color,
    backgroundColor: style.bgColor,
    borderTop: style.borderTop,
    borderRight: style.borderRight,
    borderBottom: style.borderBottom,
    borderLeft: style.borderLeft,
    textAlign: style.textAlign,
    verticalAlign: style.verticalAlign,
    whiteSpace: style.wrapText ? "pre-wrap" : undefined,
    transform,
  };
}

export async function parseWorkbookArrayBuffer(buf: ArrayBuffer): Promise<EmbeddedExcelFieldValue> {
  const ExcelJS = (await import("exceljs")).default;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buf);

  const sheets: EmbeddedExcelSheetData[] = [];
  const mediaById = new Map<number, string>();
  const media = (workbook as unknown as { model?: { media?: unknown[] } }).model?.media ?? [];
  for (let i = 0; i < media.length; i++) {
    const src = mediaEntryToDataUrl(media[i]);
    if (!src) continue;
    const ids = mediaEntryIdCandidates(media[i], i);
    for (const id of ids) mediaById.set(id, src);
  }
  const wbGetImage = (workbook as unknown as { getImage?: (id: number) => unknown }).getImage;
  const imageCache = new Map<number, string | undefined>();
  const resolveImage = (imageId: number): string | undefined => {
    const ids = [imageId, imageId - 1, imageId + 1];
    for (const id of ids) {
      if (imageCache.has(id)) {
        const hit = imageCache.get(id);
        if (hit) return hit;
        continue;
      }
      const direct = mediaById.get(id);
      if (direct) {
        imageCache.set(id, direct);
        return direct;
      }
      let byApi: string | undefined;
      if (wbGetImage) {
        try {
          byApi = mediaEntryToDataUrl(wbGetImage.call(workbook, id));
        } catch {
          byApi = undefined;
        }
      }
      imageCache.set(id, byApi);
      if (byApi) return byApi;
    }
    return undefined;
  };

  for (const ws of workbook.worksheets) {
    const name = ws.name;
    const bounds = computeExpandedBounds(ws);
    if (!bounds) {
      sheets.push({ name, grid: [[""]], cellStyles: [[undefined]] });
      continue;
    }

    const { top, left, bottom, right } = bounds;
    const numRows = bottom - top + 1;
    const numCols = right - left + 1;

    if (numRows < 1 || numCols < 1) {
      sheets.push({ name, grid: [[""]], cellStyles: [[undefined]] });
      continue;
    }

    const cellStyles: (EmbeddedExcelCellStyle | undefined)[][] = [];
    const colWidthsPx: number[] = [];
    const rowHeightsPx: (number | undefined)[] = [];

    const fullRows = bottom;
    const fullCols = right;
    const fullHf: RawCellContent[][] = Array.from({ length: fullRows }, () =>
      Array(fullCols).fill(null) as RawCellContent[],
    );
    const formulaCells: boolean[][] = Array.from({ length: numRows }, () => Array(numCols).fill(false));

    for (let gc = 0; gc < numCols; gc++) {
      const col = ws.getColumn(left + gc);
      const w = col.width;
      if (typeof w === "number" && w > 0) {
        colWidthsPx.push(Math.round(w * 7 + 8));
      } else {
        colWidthsPx.push(72);
      }
    }

    const yieldEvery = numRows > 2000 ? 600 : 0;

    for (let gr = 0; gr < numRows; gr++) {
      if (yieldEvery > 0 && gr > 0 && gr % yieldEvery === 0) {
        await yieldToBrowser();
      }

      const excelRow = top + gr;
      const row = ws.getRow(excelRow);
      const h = row.height;
      if (typeof h === "number" && h > 0) {
        rowHeightsPx.push(Math.round((h * 96) / 72));
      } else {
        rowHeightsPx.push(undefined);
      }

      const styleRow: (EmbeddedExcelCellStyle | undefined)[] = [];

      for (let gc = 0; gc < numCols; gc++) {
        const excelCol = left + gc;
        const cell = ws.getCell(excelRow, excelCol);
        const { content, isFormula } = cellToHFContent(cell);
        fullHf[excelRow - 1][excelCol - 1] = content;

        if (cell.isMerged && cell.address !== cell.master.address) {
          styleRow.push(undefined);
          continue;
        }
        formulaCells[gr][gc] = isFormula;
        styleRow.push(extractCellStyle(cell));
      }
      cellStyles.push(styleRow);
    }

    const key = hfSheetKey(name);
    let displayGrid: string[][] = [];
    let hf: HyperFormula | undefined;
    try {
      hf = buildHyperFormulaEngine({ [key]: fullHf });
      const sid = hf.getSheetId(key);
      if (sid !== undefined) {
        const vals = hf.getSheetValues(sid);
        for (let gr = 0; gr < numRows; gr++) {
          const excelRow = top + gr;
          const rowOut: string[] = [];
          for (let gc = 0; gc < numCols; gc++) {
            const excelCol = left + gc;
            const raw = vals[excelRow - 1]?.[excelCol - 1];
            const shown = formatHyperFormulaDisplay(raw ?? null);
            if (formulaCells[gr]?.[gc] && shown.startsWith("#")) {
              // If HyperFormula cannot evaluate a valid Excel formula, keep Excel's cached display.
              rowOut.push(cellDisplayText(ws.getCell(excelRow, excelCol)) || shown);
            } else {
              rowOut.push(shown);
            }
          }
          displayGrid.push(rowOut);
        }
      }
    } catch {
      displayGrid = [];
    } finally {
      hf?.destroy();
    }

    if (!displayGrid.length) {
      for (let gr = 0; gr < numRows; gr++) {
        const excelRow = top + gr;
        const textRow: string[] = [];
        for (let gc = 0; gc < numCols; gc++) {
          const cell = ws.getCell(excelRow, left + gc);
          if (cell.isMerged && cell.address !== cell.master.address) {
            textRow.push("");
          } else {
            textRow.push(cellDisplayText(cell));
          }
        }
        displayGrid.push(textRow);
      }
    }

    const mergeRegions = mergeRegionsFromWorksheet(
      (ws.model.merges ?? []) as string[],
      top,
      left,
      numRows,
      numCols,
    );

    const baseGrid = displayGrid.length ? displayGrid : [[""]];
    let trimTop = 0;
    let trimLeft = 0;
    let trimBottom = numRows - 1;
    let trimRight = numCols - 1;

    const rowIsEmpty = (ri: number): boolean => {
      for (let ci = 0; ci < numCols; ci++) {
        if ((baseGrid[ri]?.[ci] ?? "") !== "") return false;
        if (formulaCells[ri]?.[ci]) return false;
        if (hasMeaningfulStyle(cellStyles[ri]?.[ci])) return false;
      }
      return true;
    };
    const colIsEmpty = (ci: number): boolean => {
      for (let ri = 0; ri < numRows; ri++) {
        if ((baseGrid[ri]?.[ci] ?? "") !== "") return false;
        if (formulaCells[ri]?.[ci]) return false;
        if (hasMeaningfulStyle(cellStyles[ri]?.[ci])) return false;
      }
      return true;
    };

    while (trimTop < numRows - 1 && rowIsEmpty(trimTop)) trimTop++;
    while (trimLeft < numCols - 1 && colIsEmpty(trimLeft)) trimLeft++;
    while (trimBottom > trimTop && rowIsEmpty(trimBottom)) trimBottom--;
    while (trimRight > trimLeft && colIsEmpty(trimRight)) trimRight--;

    const grid = baseGrid.slice(trimTop, trimBottom + 1).map((r) => r.slice(trimLeft, trimRight + 1));
    const trimmedStyles = cellStyles
      .slice(trimTop, trimBottom + 1)
      .map((r) => r.slice(trimLeft, trimRight + 1));
    const trimmedFormulaCells = formulaCells
      .slice(trimTop, trimBottom + 1)
      .map((r) => r.slice(trimLeft, trimRight + 1));
    const trimmedColWidths = colWidthsPx.slice(trimLeft, trimRight + 1);
    const trimmedRowHeights = rowHeightsPx.slice(trimTop, trimBottom + 1);
    const trimmedMerges = mergeRegions
      .map((m) => ({ ...m, r: m.r - trimTop, c: m.c - trimLeft }))
      .filter(
        (m) =>
          m.r >= 0 &&
          m.c >= 0 &&
          m.r < grid.length &&
          m.c < (grid[0]?.length ?? 0),
      );

    const images = extractWorksheetImages(
      ws,
      resolveImage,
      left + trimLeft,
      top + trimTop,
      grid[0]?.length ?? 0,
      grid.length,
      trimmedColWidths,
      trimmedRowHeights,
    );

    sheets.push({
      name,
      grid: grid.length ? grid : [[""]],
      colWidthsPx: trimmedColWidths,
      rowHeightsPx: trimmedRowHeights,
      mergeRegions: trimmedMerges,
      cellStyles: trimmedStyles.length ? trimmedStyles : [[undefined]],
      hfFullSheet: fullHf,
      excelBounds: {
        top: top + trimTop,
        left: left + trimLeft,
        bottom: top + trimBottom,
        right: left + trimRight,
      },
      formulaCells: trimmedFormulaCells,
      images,
    });
  }

  return { sheets: sheets.length ? sheets : [{ name: "Sheet1", grid: [[""]], cellStyles: [[undefined]] }] };
}

export async function loadWorkbookFromUrl(url: string): Promise<EmbeddedExcelFieldValue> {
  const fetchUrl = url.startsWith("http://") || url.startsWith("https://") ? url : encodeURI(url);
  const res = await fetch(fetchUrl);
  if (!res.ok) {
    throw new Error(`Could not load file (${res.status}). Put the .xlsx in the public folder or check the path.`);
  }
  const buf = await res.arrayBuffer();
  const parsed = await parseWorkbookArrayBuffer(buf);
  return { ...parsed, sourceUrl: url };
}

export function dataUrlToArrayBuffer(dataUrl: string): ArrayBuffer {
  const i = dataUrl.indexOf(",");
  if (i === -1) throw new Error("Invalid data URL");
  const base64 = dataUrl.slice(i + 1);
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let j = 0; j < binary.length; j++) bytes[j] = binary.charCodeAt(j);
  return bytes.buffer;
}

export async function loadEmbeddedExcelSource(source: string): Promise<EmbeddedExcelFieldValue> {
  const s = source.trim();
  if (!s) {
    throw new Error("No Excel source");
  }
  if (s.startsWith("data:")) {
    const buf = dataUrlToArrayBuffer(s);
    const parsed = await parseWorkbookArrayBuffer(buf);
    return { ...parsed, sourceUrl: "embedded" };
  }
  return loadWorkbookFromUrl(s);
}

export function mergeEmbeddedExcelValue(
  loaded: EmbeddedExcelFieldValue,
  saved: EmbeddedExcelFieldValue | null | undefined,
): EmbeddedExcelFieldValue {
  if (!saved?.sheets?.length) return loaded;
  const sheets = loaded.sheets.map((sh, si) => {
    const savedSh = saved.sheets.find((s) => s.name === sh.name) ?? saved.sheets[si];
    if (!savedSh?.grid?.length) return sh;
    const grid = sh.grid.map((row, ri) =>
      row.map((cell, ci) => {
        if (sh.formulaCells?.[ri]?.[ci]) return cell;
        const sv = savedSh.grid[ri]?.[ci];
        if (sv === undefined || sv === "") return cell;
        return String(sv);
      }),
    );
    return {
      ...sh,
      grid,
      colWidthsPx: sh.colWidthsPx,
      rowHeightsPx: sh.rowHeightsPx,
      mergeRegions: sh.mergeRegions,
      cellStyles: sh.cellStyles,
      hfFullSheet: sh.hfFullSheet,
      excelBounds: sh.excelBounds,
      formulaCells: sh.formulaCells,
      images: savedSh.images?.length ? savedSh.images : sh.images,
    };
  });
  return { sheets, sourceUrl: loaded.sourceUrl ?? saved.sourceUrl };
}

export function getContainingMergeRegion(
  regions: Array<{ r: number; c: number; rowspan: number; colspan: number }> | undefined,
  r: number,
  c: number,
): { r: number; c: number; rowspan: number; colspan: number } | null {
  if (!regions?.length) return null;
  for (const m of regions) {
    if (r >= m.r && c >= m.c && r < m.r + m.rowspan && c < m.c + m.colspan) {
      return m;
    }
  }
  return null;
}

export function isMergeSkipCell(
  regions: Array<{ r: number; c: number; rowspan: number; colspan: number }> | undefined,
  r: number,
  c: number,
): boolean {
  const m = getContainingMergeRegion(regions, r, c);
  return m != null && (r !== m.r || c !== m.c);
}

export function getMergeSpanIfMaster(
  regions: Array<{ r: number; c: number; rowspan: number; colspan: number }> | undefined,
  r: number,
  c: number,
): { rowspan: number; colspan: number } | null {
  if (!regions?.length) return null;
  for (const m of regions) {
    if (m.r === r && m.c === c) return { rowspan: m.rowspan, colspan: m.colspan };
  }
  return null;
}
