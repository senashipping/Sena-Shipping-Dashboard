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
    checkedTemplate?: string;
    uncheckedTemplate?: string;
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

export interface HandsontableWorkbookProps {
  data: { sheets: SheetData[] };
  onChange: (next: { sheets: SheetData[] }) => void;
  readOnly?: boolean;
  strictViewOnly?: boolean;
  readOnlyHotHeight?: number;
  lightweightPerformance?: boolean;
  /**
   * Use the same fixed HOT height (320px) and virtualization as the admin
   * "Edit Workbook" dialog instead of a tall read-only viewport.
   * Intended for form preview, submission page, and submission detail modal.
   */
  embeddedExcelMatchEditorViewport?: boolean;
}

export type HandsontableWorkbookRef = {
  getWorkbookSnapshot: () => { sheets: SheetData[] } | null;
};

export const MAX_PREVIEW_ROWS = 500;
export const MAX_PREVIEW_COLS = 100;

