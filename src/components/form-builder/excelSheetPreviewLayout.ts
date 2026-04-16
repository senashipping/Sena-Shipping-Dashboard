/**
 * Single layout for every read-only / preview Handsontable (form preview, submission page, admin submission modal).
 */

export const EXCEL_PREVIEW_HOT_HEIGHT_MIN_PX = 480;
export const EXCEL_PREVIEW_HOT_HEIGHT_MAX_PX = 860;
export const EXCEL_PREVIEW_VIEWPORT_HEIGHT_RATIO = 0.66;
/** When `window` is unavailable (SSR) or first paint. */
export const EXCEL_PREVIEW_HOT_HEIGHT_SSR_FALLBACK_PX = 620;

export function getExcelPreviewHotHeightPx(): number {
  if (typeof window === "undefined") return EXCEL_PREVIEW_HOT_HEIGHT_SSR_FALLBACK_PX;
  return Math.min(
    EXCEL_PREVIEW_HOT_HEIGHT_MAX_PX,
    Math.max(
      EXCEL_PREVIEW_HOT_HEIGHT_MIN_PX,
      Math.round(window.innerHeight * EXCEL_PREVIEW_VIEWPORT_HEIGHT_RATIO),
    ),
  );
}

/** Dialog shell when the form (or submission) includes an embedded Excel workbook. */
export const EXCEL_PREVIEW_DIALOG_CONTENT_CLASS =
  "mx-4 max-h-[calc(100dvh_-_1.25rem)] w-[min(100rem,calc(100vw_-_2rem))] max-w-[min(100rem,calc(100vw_-_2rem))] overflow-y-auto sm:mx-6";

/** Direct wrapper around read-only Handsontable — full width of host, horizontal scroll if needed. */
export const EXCEL_PREVIEW_SHEET_FRAME_CLASS = "w-full min-w-0 overflow-x-auto";

/**
 * Same outer scroll shell as the admin Edit Workbook dialog (`PropertiesPanel`).
 * Use with `embeddedExcelMatchEditorViewport` on `HandsontableWorkbook` (320px HOT height).
 */
export const EXCEL_RUNTIME_MATCH_EDITOR_FRAME_CLASS =
  "max-h-[65vh] min-h-0 overflow-y-auto";

export function formDefinitionHasEmbeddedExcel(
  fields: { type: string }[] | undefined,
  sections: { type?: string; fields?: { type: string }[] }[] | undefined,
): boolean {
  if (fields?.some((f) => f.type === "embedded_excel")) return true;
  return (sections ?? []).some(
    (s) => s.type === "fields" && (s.fields ?? []).some((f) => f.type === "embedded_excel"),
  );
}
